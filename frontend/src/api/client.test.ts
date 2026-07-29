import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { server } from '@/test/server'
import api, {
  downloadAuthenticated,
  getAccessToken,
  setAccessToken,
  setSessionExpiredHandler,
} from '@/api/client'

describe('client', () => {
  beforeEach(() => {
    setAccessToken(null)
    setSessionExpiredHandler(null)
  })

  afterEach(() => {
    setAccessToken(null)
    setSessionExpiredHandler(null)
  })

  describe('downloadAuthenticated', () => {
    beforeEach(() => {
      // jsdom não implementa createObjectURL/revokeObjectURL para Blob.
      vi.stubGlobal('URL', Object.assign(URL, {
        createObjectURL: vi.fn(() => 'blob:mock-url'),
        revokeObjectURL: vi.fn(),
      }))
    })

    afterEach(() => {
      vi.unstubAllGlobals()
      vi.useRealTimers()
    })

    it('monta a URL certa e dispara o download quando o caminho já vem com prefixo /api', async () => {
      setAccessToken('tok-123')
      let recebido: { auth: string | null; url: string } | null = null
      server.use(
        http.get('/api/documents/files/:categoria/:arquivo', ({ request, params }) => {
          recebido = {
            auth: request.headers.get('Authorization'),
            url: `/documents/files/${params.categoria}/${params.arquivo}`,
          }
          return HttpResponse.text('conteudo-do-arquivo', {
            headers: { 'Content-Type': 'application/pdf' },
          })
        })
      )

      const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})

      await downloadAuthenticated('/api/documents/files/remocao/nota.pdf', 'nota.pdf')

      expect(recebido).toEqual({ auth: 'Bearer tok-123', url: '/documents/files/remocao/nota.pdf' })
      expect(URL.createObjectURL).toHaveBeenCalledTimes(1)
      expect(clickSpy).toHaveBeenCalledTimes(1)

      clickSpy.mockRestore()
    })

    it('monta a URL certa quando o caminho vem sem o prefixo /api (não duplica o baseURL)', async () => {
      setAccessToken('tok-456')
      let chamadas = 0
      server.use(
        http.get('/api/documents/files/termos/termo.pdf', ({ request }) => {
          chamadas += 1
          expect(request.headers.get('Authorization')).toBe('Bearer tok-456')
          return HttpResponse.text('termo')
        })
      )

      const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})

      await downloadAuthenticated('/documents/files/termos/termo.pdf', 'termo.pdf')

      expect(chamadas).toBe(1)
      clickSpy.mockRestore()
    })

    it('cria o link com o nome de arquivo correto e insere/remove do DOM (fluxo <a download> real não envia Authorization, por isso passa pelo axios)', async () => {
      setAccessToken('tok-789')
      server.use(http.get('/api/documents/files/x/y.pdf', () => HttpResponse.text('bytes')))

      let anchorCapturado: HTMLAnchorElement | null = null
      const originalCreateElement = document.createElement.bind(document)
      const createElementSpy = vi.spyOn(document, 'createElement').mockImplementation((tag: string) => {
        const el = originalCreateElement(tag)
        if (tag === 'a') anchorCapturado = el as HTMLAnchorElement
        return el
      })
      const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {})

      await downloadAuthenticated('/api/documents/files/x/y.pdf', 'meu-arquivo.pdf')

      expect(anchorCapturado).not.toBeNull()
      expect(anchorCapturado!.download).toBe('meu-arquivo.pdf')
      expect(anchorCapturado!.href).toContain('blob:mock-url')

      createElementSpy.mockRestore()
      clickSpy.mockRestore()
    })
  })

  describe('interceptor de 401', () => {
    it('um 401 comum dispara refresh e reexecuta a requisição original com o token novo', async () => {
      setAccessToken('token-velho')
      let refreshCalls = 0

      server.use(
        http.post('/api/auth/refresh', () => {
          refreshCalls += 1
          return HttpResponse.json({ access_token: 'token-novo' })
        }),
        http.get('/api/recurso-protegido', ({ request }) => {
          const auth = request.headers.get('Authorization')
          if (auth !== 'Bearer token-novo') {
            return new HttpResponse(null, { status: 401 })
          }
          return HttpResponse.json({ ok: true })
        })
      )

      const res = await api.get('/recurso-protegido')

      expect(res.data).toEqual({ ok: true })
      expect(refreshCalls).toBe(1)
      expect(getAccessToken()).toBe('token-novo')
    })

    it('um 401 em /auth/login não dispara refresh nem "sessão expirada" (senha errada é resposta de negócio)', async () => {
      let refreshCalls = 0
      const sessionExpiredHandler = vi.fn()
      setSessionExpiredHandler(sessionExpiredHandler)

      server.use(
        http.post('/api/auth/refresh', () => {
          refreshCalls += 1
          return HttpResponse.json({ access_token: 'nao-devia-existir' })
        }),
        http.post('/api/auth/login', () =>
          HttpResponse.json({ detail: 'Usuário ou senha inválidos.' }, { status: 401 })
        )
      )

      await expect(api.post('/auth/login', { username: 'a', password: 'errada' })).rejects.toMatchObject({
        response: { status: 401 },
      })

      expect(refreshCalls).toBe(0)
      expect(sessionExpiredHandler).not.toHaveBeenCalled()
    })

    it('falha no refresh aciona o callback de sessão expirada e limpa o token', async () => {
      setAccessToken('token-velho')
      const sessionExpiredHandler = vi.fn()
      setSessionExpiredHandler(sessionExpiredHandler)

      server.use(
        http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 })),
        http.get('/api/recurso-protegido', () => new HttpResponse(null, { status: 401 }))
      )

      await expect(api.get('/recurso-protegido')).rejects.toMatchObject({
        response: { status: 401 },
      })

      expect(sessionExpiredHandler).toHaveBeenCalledTimes(1)
      expect(getAccessToken()).toBeNull()
    })

    it('requisições concorrentes durante o refresh entram na fila e são reexecutadas com o token novo', async () => {
      setAccessToken('token-velho')
      let refreshCalls = 0

      server.use(
        http.post('/api/auth/refresh', async () => {
          refreshCalls += 1
          return HttpResponse.json({ access_token: 'token-concorrente' })
        }),
        http.get('/api/recurso-a', ({ request }) =>
          request.headers.get('Authorization') === 'Bearer token-concorrente'
            ? HttpResponse.json({ recurso: 'a' })
            : new HttpResponse(null, { status: 401 })
        ),
        http.get('/api/recurso-b', ({ request }) =>
          request.headers.get('Authorization') === 'Bearer token-concorrente'
            ? HttpResponse.json({ recurso: 'b' })
            : new HttpResponse(null, { status: 401 })
        )
      )

      const [resA, resB] = await Promise.all([api.get('/recurso-a'), api.get('/recurso-b')])

      expect(resA.data).toEqual({ recurso: 'a' })
      expect(resB.data).toEqual({ recurso: 'b' })
      // O ponto central do teste: só UM refresh deve ter sido disparado para as duas
      // requisições concorrentes, não um por requisição.
      expect(refreshCalls).toBe(1)
      expect(getAccessToken()).toBe('token-concorrente')
    })
  })
})
