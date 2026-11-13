import { beforeEach, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { act, screen, waitFor, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import App from '@/App'
import { ThemeProvider } from '@/lib/theme'

const mocks = vi.hoisted(() => ({ iniciarLeitor: vi.fn() }))

vi.mock('@/features/inventario/scanner', async (original) => ({
  ...(await original<typeof import('@/features/inventario/scanner')>()),
  iniciarLeitor: mocks.iniciarLeitor,
}))

import { ErroCamera } from '@/features/inventario/scanner'

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'teste', role: 'Técnico' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Técnico'),
  }),
}))

const resumo = {
  esperados: 10,
  lidos: 2,
  encontrados: 2,
  ok: 2,
  local_errado: 0,
  usuario_diferente: 0,
  outra_unidade: 0,
  nao_cadastrado: 0,
  nao_encontrados: 8,
}

function inventario(status = 'aberto') {
  return {
    id: 12,
    unidade_id: 1,
    unidade_nome: 'Juazeiro',
    status,
    iniciado_por: 'tecnico.teste',
    iniciado_em: '2026-11-10T10:00:00',
    total_leituras: 2,
    resumo,
  }
}

function leitura(extra: Record<string, unknown> = {}) {
  return {
    id: 100,
    inventario_id: 12,
    item_id: 12,
    codigo_lido: '#12',
    lido_por: 'tecnico.teste',
    lido_em: '2026-11-10T10:05:00',
    resultado: 'ok',
    detalhe: null,
    usuario_conferido: false,
    usuario_esperado: 'Maria Souza',
    item_tipo: 'Notebook',
    item_brand: 'Dell',
    item_model: 'Latitude',
    duplicada: false,
    ...extra,
  }
}

let corposPost: unknown[] = []
let status = 'aberto'

beforeEach(() => {
  corposPost = []
  status = 'aberto'
  mocks.iniciarLeitor.mockReset()
  Object.defineProperty(window, 'isSecureContext', { value: true, configurable: true })
  Object.defineProperty(document, 'visibilityState', { value: 'visible', configurable: true })
  server.use(
    http.get('/api/inventarios/12', () => HttpResponse.json(inventario(status))),
    http.get('/api/inventarios/12/leituras', () => HttpResponse.json({ items: [], total: 0 })),
    http.get('/api/locais', () =>
      HttpResponse.json([
        { id: 3, unidade_id: 1, nome: 'Recepção', ativo: true },
        { id: 4, unidade_id: 1, nome: 'Sala TI', ativo: true },
      ])
    ),
    http.post('/api/inventarios/12/leituras', async ({ request }) => {
      corposPost.push(await request.json())
      return HttpResponse.json(leitura(), { status: 201 })
    })
  )
})

function renderRota() {
  return renderWithClient(
    <ThemeProvider>
      <MemoryRouter initialEntries={['/inventarios/12/leitura']}>
        <App />
      </MemoryRouter>
    </ThemeProvider>
  )
}

const CAMPO = 'Código ou número do patrimônio'

describe('InventarioLeituraPage', () => {
  it('digitação registra com o local atual', async () => {
    const user = userEvent.setup()
    renderRota()
    expect(await screen.findByText('Inventário #12 · JUAZEIRO')).toBeInTheDocument()
    expect(screen.getByText('Encontrados 2 de 10 · Faltam 8')).toBeInTheDocument()
    await screen.findByRole('option', { name: 'Recepção' })
    // O nome acessível do seletor é o mesmo do rótulo visível.
    await user.selectOptions(screen.getByRole('combobox', { name: 'Estou no local' }), 'Recepção')
    const campo = screen.getByLabelText(CAMPO)

    await user.type(campo, '#12{Enter}')

    expect(await screen.findByText('Conferido')).toBeInTheDocument()
    expect(screen.getByText(/#12 · Notebook/)).toBeInTheDocument()
    expect(corposPost).toEqual([{ codigo: '#12', local_id_encontrado: 3 }])
    await waitFor(() => expect(campo).toHaveValue(''))
    expect(campo).toHaveFocus()
  })

  it('o botão Registrar também envia', async () => {
    const user = userEvent.setup()
    renderRota()
    await user.type(await screen.findByLabelText(CAMPO), 'SN-9')
    await user.click(screen.getByRole('button', { name: 'Registrar' }))

    expect(await screen.findByText('Conferido')).toBeInTheDocument()
    expect(corposPost).toEqual([{ codigo: 'SN-9', local_id_encontrado: null }])
  })

  it('leitura duplicada avisa que não contou de novo', async () => {
    server.use(
      http.post('/api/inventarios/12/leituras', () =>
        HttpResponse.json(leitura({ duplicada: true }), { status: 200 })
      )
    )
    const user = userEvent.setup()
    renderRota()

    await user.type(await screen.findByLabelText(CAMPO), '#12{Enter}')

    expect(await screen.findByText(/não contou de novo/)).toBeInTheDocument()
    expect(screen.getByText(/Já lido em .* por tecnico\.teste/)).toBeInTheDocument()
  })

  it('câmera negada mostra a mensagem e foca a digitação', async () => {
    mocks.iniciarLeitor.mockRejectedValue(new ErroCamera('negada'))
    const user = userEvent.setup()
    renderRota()

    await user.click(await screen.findByRole('button', { name: 'Ler com a câmera' }))

    expect(await screen.findByRole('alert')).toHaveTextContent('Permissão da câmera negada')
    expect(screen.getByLabelText(CAMPO)).toHaveFocus()
    // A câmera foi desligada: volta a oferecer "Ler com a câmera".
    expect(screen.getByRole('button', { name: 'Ler com a câmera' })).toBeInTheDocument()
  })

  it('código lido pela câmera é registrado, e "Parar câmera" desliga', async () => {
    let entregar: (t: string) => void = () => {}
    const parar = vi.fn()
    mocks.iniciarLeitor.mockImplementation(async (_video, cb) => {
      entregar = cb
      return { parar }
    })
    const user = userEvent.setup()
    renderRota()

    await user.click(await screen.findByRole('button', { name: 'Ler com a câmera' }))
    await screen.findByLabelText('Imagem da câmera')
    await act(async () => entregar('https://p.revalle.com.br/e/12'))

    expect(await screen.findByText('Conferido')).toBeInTheDocument()
    expect(corposPost).toEqual([{ codigo: 'https://p.revalle.com.br/e/12', local_id_encontrado: null }])

    await user.click(screen.getByRole('button', { name: 'Parar câmera' }))
    expect(parar).toHaveBeenCalled()
    expect(screen.queryByLabelText('Imagem da câmera')).not.toBeInTheDocument()
  })

  it('aba escondida pausa a câmera e explica', async () => {
    mocks.iniciarLeitor.mockResolvedValue({ parar: vi.fn() })
    const user = userEvent.setup()
    renderRota()
    await user.click(await screen.findByRole('button', { name: 'Ler com a câmera' }))
    await screen.findByLabelText('Imagem da câmera')

    act(() => {
      Object.defineProperty(document, 'visibilityState', { value: 'hidden', configurable: true })
      document.dispatchEvent(new Event('visibilitychange'))
    })

    expect(await screen.findByText(/Câmera pausada/)).toBeInTheDocument()
    expect(screen.queryByLabelText('Imagem da câmera')).not.toBeInTheDocument()
  })

  it('"Ninguém usando" corrige a leitura', async () => {
    let corpo: unknown = null
    server.use(
      http.patch('/api/inventarios/12/leituras/100', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(
          leitura({ resultado: 'usuario_diferente', usuario_conferido: true, detalhe: 'Usuário cadastrado: Maria Souza · encontrado com: ninguém.' })
        )
      })
    )
    const user = userEvent.setup()
    renderRota()
    await user.type(await screen.findByLabelText(CAMPO), '#12{Enter}')
    await screen.findByText('Conferido')

    await user.click(screen.getByRole('button', { name: 'Ninguém usando' }))

    expect(await screen.findByText('Usuário diferente')).toBeInTheDocument()
    expect(corpo).toEqual({ usuario_conferido: true, colaborador_id_encontrado: null })
    expect(screen.getByText(/encontrado com: ninguém/)).toBeInTheDocument()
  })

  it('"Informar usuário" escolhe o colaborador e corrige a leitura', async () => {
    let corpo: unknown = null
    server.use(
      http.get('/api/colaboradores', () =>
        HttpResponse.json({
          items: [
            {
              id: 5, cpf: '12345678909', nome: 'JOSE LIMA', situacao: 'ativo',
              email_manual: false, telefone_manual: false, origem: 'rh',
            },
          ],
          total: 1,
        })
      ),
      http.patch('/api/inventarios/12/leituras/100', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json(leitura({ resultado: 'usuario_diferente', usuario_conferido: true }))
      })
    )
    const user = userEvent.setup()
    renderRota()
    await user.type(await screen.findByLabelText(CAMPO), '#12{Enter}')
    await screen.findByText('Conferido')

    await user.click(screen.getByRole('button', { name: 'Informar usuário' }))
    await user.click(await screen.findByRole('button', { name: /quem está usando/i }))
    await user.click(await screen.findByText('JOSE LIMA'))

    expect(await screen.findByText('Usuário diferente')).toBeInTheDocument()
    expect(corpo).toEqual({ colaborador_id_encontrado: 5 })
  })

  it('desfazer leitura chama DELETE', async () => {
    let apagou = false
    server.use(
      http.delete('/api/inventarios/12/leituras/100', () => {
        apagou = true
        return HttpResponse.json({ detail: 'Leitura excluída.' })
      })
    )
    const user = userEvent.setup()
    renderRota()
    await user.type(await screen.findByLabelText(CAMPO), '#12{Enter}')
    await screen.findByText('Conferido')

    await user.click(screen.getByRole('button', { name: 'Desfazer leitura' }))
    const dialogo = await screen.findByRole('alertdialog')
    expect(within(dialogo).getByText('Desfazer esta leitura?')).toBeInTheDocument()
    await user.click(within(dialogo).getByRole('button', { name: 'Desfazer' }))

    await waitFor(() => expect(apagou).toBe(true))
    await waitFor(() => expect(screen.queryByText('Conferido')).not.toBeInTheDocument())
  })

  it('inventário encerrado não oferece leitura', async () => {
    status = 'encerrado'
    renderRota()

    expect(await screen.findByText('Este inventário não está aberto.')).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Ver relatório' })).toHaveAttribute('href', '/inventarios/12')
    expect(screen.queryByLabelText(CAMPO)).not.toBeInTheDocument()
  })

  it('sem HTTPS a câmera fica desabilitada e a digitação continua', async () => {
    Object.defineProperty(window, 'isSecureContext', { value: false, configurable: true })
    const user = userEvent.setup()
    renderRota()

    expect(await screen.findByRole('button', { name: 'Ler com a câmera' })).toBeDisabled()
    expect(screen.getByText(/só funciona em conexão segura \(HTTPS\)/)).toBeInTheDocument()
    await user.type(screen.getByLabelText(CAMPO), '#12{Enter}')
    expect(await screen.findByText('Conferido')).toBeInTheDocument()
  })

  it('código B visto enquanto A grava não se perde: é enviado uma vez quando A termina', async () => {
    let abrir: () => void = () => {}
    const espera = new Promise<void>((r) => (abrir = r))
    server.use(
      http.post('/api/inventarios/12/leituras', async ({ request }) => {
        const corpo = (await request.json()) as { codigo: string }
        corposPost.push(corpo)
        if (corpo.codigo === 'A') await espera
        return HttpResponse.json(leitura({ codigo_lido: corpo.codigo }), { status: 201 })
      })
    )
    let entregar: (t: string) => boolean | void = () => {}
    mocks.iniciarLeitor.mockImplementation(async (_video, cb) => {
      entregar = cb
      return { parar: vi.fn() }
    })
    const user = userEvent.setup()
    renderRota()
    await user.click(await screen.findByRole('button', { name: 'Ler com a câmera' }))
    await screen.findByLabelText('Imagem da câmera')

    await act(async () => {
      entregar('A')
    })
    await waitFor(() => expect(corposPost).toHaveLength(1))
    await act(async () => {
      entregar('B') // A ainda está gravando: B é recusado, sem POST
    })
    expect(corposPost).toHaveLength(1)

    await act(async () => abrir())
    await screen.findByText('Conferido')
    await act(async () => {
      entregar('B') // B continua à vista: agora passa
      entregar('B') // e o quadro seguinte é só repetição
    })

    await waitFor(() => expect(corposPost.map((c) => (c as { codigo: string }).codigo)).toEqual(['A', 'B']))
  })

  it('409 (encerrado por outra pessoa) mostra que o inventário não está aberto e para de ler', async () => {
    server.use(
      http.post('/api/inventarios/12/leituras', () => {
        status = 'encerrado'
        return HttpResponse.json(
          { detail: 'Este inventário não está aberto: não aceita leituras.' },
          { status: 409 }
        )
      })
    )
    const user = userEvent.setup()
    renderRota()

    await user.type(await screen.findByLabelText(CAMPO), '#12{Enter}')

    expect(await screen.findByText('Este inventário não está aberto.')).toBeInTheDocument()
    expect(screen.queryByLabelText(CAMPO)).not.toBeInTheDocument()
  })

  it('o campo esvazia ao enviar, mesmo quando a gravação falha', async () => {
    server.use(
      http.post('/api/inventarios/12/leituras', () => HttpResponse.json({ detail: 'Falha ao gravar.' }, { status: 500 }))
    )
    const user = userEvent.setup()
    renderRota()
    const campo = await screen.findByLabelText(CAMPO)

    await user.type(campo, 'SN-1{Enter}')

    expect(campo).toHaveValue('')
    expect(await screen.findByText('Falha ao gravar.')).toBeInTheDocument()
    expect(campo).toHaveFocus()
  })

  it('depois de uma leitura pela câmera o foco não vai para o campo (não abre o teclado sobre o vídeo)', async () => {
    let entregar: (t: string) => boolean | void = () => {}
    mocks.iniciarLeitor.mockImplementation(async (_video, cb) => {
      entregar = cb
      return { parar: vi.fn() }
    })
    const user = userEvent.setup()
    renderRota()
    await user.click(await screen.findByRole('button', { name: 'Ler com a câmera' }))
    await screen.findByLabelText('Imagem da câmera')

    await act(async () => {
      entregar('https://p.revalle.com.br/e/12')
    })

    expect(await screen.findByText('Conferido')).toBeInTheDocument()
    expect(screen.getByLabelText(CAMPO)).not.toHaveFocus()
  })

  it('segundo código digitado/lido pelo USB durante a gravação entra na fila: campo limpo e envio em ordem', async () => {
    let abrir: () => void = () => {}
    const espera = new Promise<void>((r) => (abrir = r))
    server.use(
      http.post('/api/inventarios/12/leituras', async ({ request }) => {
        const corpo = (await request.json()) as { codigo: string }
        corposPost.push(corpo)
        if (corpo.codigo === 'A') await espera
        return HttpResponse.json(leitura({ codigo_lido: corpo.codigo }), { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderRota()
    const campo = await screen.findByLabelText(CAMPO)

    await user.type(campo, 'A{Enter}')
    await waitFor(() => expect(corposPost).toHaveLength(1))
    await user.type(campo, 'B{Enter}')
    expect(campo).toHaveValue('')
    await user.type(campo, 'B{Enter}') // repetição exata do último da fila: ignorada
    expect(campo).toHaveValue('')
    expect(corposPost).toHaveLength(1)

    await act(async () => abrir())

    await waitFor(() => expect(corposPost.map((c) => (c as { codigo: string }).codigo)).toEqual(['A', 'B']))
    await new Promise((r) => setTimeout(r, 100))
    expect(corposPost).toHaveLength(2)
  })

  it('409 na gravação em andamento descarta a fila', async () => {
    let abrir: () => void = () => {}
    const espera = new Promise<void>((r) => (abrir = r))
    server.use(
      http.post('/api/inventarios/12/leituras', async ({ request }) => {
        const corpo = (await request.json()) as { codigo: string }
        corposPost.push(corpo)
        if (corpo.codigo === 'A') {
          await espera
          return HttpResponse.json({ detail: 'Conflito.' }, { status: 409 })
        }
        return HttpResponse.json(leitura({ codigo_lido: corpo.codigo }), { status: 201 })
      })
    )
    const user = userEvent.setup()
    renderRota()
    const campo = await screen.findByLabelText(CAMPO)

    await user.type(campo, 'A{Enter}')
    await waitFor(() => expect(corposPost).toHaveLength(1))
    await user.type(campo, 'B{Enter}')
    await act(async () => abrir())
    expect(await screen.findByText('Conflito.')).toBeInTheDocument()
    await new Promise((r) => setTimeout(r, 100))
    expect(corposPost.map((c) => (c as { codigo: string }).codigo)).toEqual(['A'])

    // A fila não ficou travada: um novo código segue normalmente.
    await user.type(campo, 'C{Enter}')
    await waitFor(() => expect(corposPost.map((c) => (c as { codigo: string }).codigo)).toEqual(['A', 'C']))
  })
})
