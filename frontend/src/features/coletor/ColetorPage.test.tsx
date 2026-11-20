import { StrictMode } from 'react'
import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes, useLocation } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import ColetorPage from '@/features/coletor/ColetorPage'
import { lerEstadoColetor } from '@/features/coletor/estadoColetor'

const logout = vi.fn(() => Promise.resolve())

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'tecnico.teste', role: 'Técnico' },
    isLoading: false,
    logout,
  }),
}))

const COMANDO =
  'powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=\'Tls12\';irm \'https://patrimonio.revalle.com.br/api/coletor/abc123\'|iex"'

const CAMPOS = { brand: 'Dell', model: 'Latitude 3420', serial: 'ABC123' }

function mockCriar(onPost?: () => void) {
  let n = 0
  server.use(
    http.post('/api/coletor', () => {
      n += 1
      onPost?.()
      return HttpResponse.json(
        {
          id: n,
          comando: COMANDO,
          url_arquivo: '/api/coletor/abc123/coletor.cmd',
          expira_em: '2026-11-20T15:00:00Z',
        },
        { status: 201 }
      )
    })
  )
}

function mockAndamento(corpo: Record<string, unknown>) {
  server.use(
    http.get('/api/coletor/:id/andamento', () =>
      HttpResponse.json({ expira_em: '2026-11-20T15:00:00Z', ...corpo })
    )
  )
}

function Destino({ rotulo }: { rotulo: string }) {
  const { state } = useLocation()
  return (
    <div>
      <p>{rotulo}</p>
      <pre data-testid="estado">{JSON.stringify(state)}</pre>
    </div>
  )
}

function renderPage(estrito = false) {
  const Raiz = estrito ? StrictMode : ({ children }: { children: React.ReactNode }) => <>{children}</>
  return renderWithClient(
    <Raiz>
    <MemoryRouter initialEntries={['/register/coletor']}>
      <Routes>
        <Route path="/register/coletor" element={<ColetorPage />} />
        <Route path="/register" element={<Destino rotulo="Tela de cadastro" />} />
        <Route path="/edit/:id" element={<Destino rotulo="Tela de edição" />} />
        <Route path="/login" element={<p>Tela de login</p>} />
      </Routes>
    </MemoryRouter>
    </Raiz>
  )
}

const ativa = {
  id: 123,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude 3420',
  status: 'Disponível',
  revenda: 'RV-JZ',
  ativo: true,
  motivos: ['mac'],
}

describe('ColetorPage', () => {
  it('mostra o comando e copia', async () => {
    mockCriar()
    mockAndamento({ status: 'aguardando' })
    const user = userEvent.setup()
    const escrever = vi.spyOn(navigator.clipboard, 'writeText')
    renderPage()

    expect(await screen.findByDisplayValue(COMANDO)).toHaveAttribute('readonly')
    expect(screen.getByText('⏳ Aguardando os dados da máquina…')).toBeInTheDocument()
    expect(screen.getByText(/Aperte/).textContent).toBe(
      'Aperte Windows + R, cole com Ctrl+V e dê Enter.'
    )
    expect(screen.getByRole('link', { name: 'Baixar o coletor' })).toHaveAttribute(
      'href',
      '/api/coletor/abc123/coletor.cmd'
    )

    // A região viva já existe (vazia) antes do aviso: o texto entra nela e é anunciado.
    const regiao = screen.getByRole('status')
    expect(regiao).toBeEmptyDOMElement()

    await user.click(screen.getByRole('button', { name: 'Copiar comando' }))

    expect(escrever).toHaveBeenCalledWith(COMANDO)
    await waitFor(() => expect(regiao).toHaveTextContent('Comando copiado.'))
    expect(screen.getByRole('status')).toBe(regiao)
  })

  it('sem clipboard seleciona o texto e orienta a copiar com Ctrl+C', async () => {
    mockCriar()
    mockAndamento({ status: 'aguardando' })
    const user = userEvent.setup()
    vi.spyOn(navigator.clipboard, 'writeText').mockRejectedValue(new Error('negado'))
    renderPage()

    await user.click(await screen.findByRole('button', { name: 'Copiar comando' }))

    expect(await screen.findByText('Selecione e copie com Ctrl+C')).toBeInTheDocument()
  })

  it('cria a coleta uma única vez, mesmo com o StrictMode', async () => {
    let posts = 0
    mockCriar(() => {
      posts += 1
    })
    mockAndamento({ status: 'aguardando' })
    renderPage(true)

    await screen.findByDisplayValue(COMANDO)
    expect(posts).toBe(1)
  })

  it('sem correspondência navega para o cadastro com os campos', async () => {
    mockCriar()
    mockAndamento({ status: 'recebido', campos: CAMPOS, tipo_sugerido: 'Notebook', correspondencias: [] })
    renderPage()

    expect(await screen.findByText('Tela de cadastro')).toBeInTheDocument()
    const estado = JSON.parse(screen.getByTestId('estado').textContent ?? 'null')
    expect(estado.campos.brand).toBe('Dell')
    expect(estado).toEqual({ origem: 'coletor', coletaId: 1, campos: CAMPOS, tipoSugerido: 'Notebook' })
  })

  it('correspondência ativa oferece atualizar', async () => {
    mockCriar()
    mockAndamento({
      status: 'recebido',
      campos: CAMPOS,
      tipo_sugerido: 'Notebook',
      correspondencias: [ativa],
    })
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText('#123 Notebook Dell Latitude 3420 · RV-JZ · mesmo MAC')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Atualizar este cadastro' }))

    expect(await screen.findByText('Tela de edição')).toBeInTheDocument()
    expect(JSON.parse(screen.getByTestId('estado').textContent ?? 'null')).toEqual({
      origem: 'coletor',
      coletaId: 1,
      campos: CAMPOS,
      tipoSugerido: 'Notebook',
    })
  })

  it('correspondência baixada não oferece atualizar', async () => {
    mockCriar()
    mockAndamento({
      status: 'recebido',
      campos: CAMPOS,
      tipo_sugerido: null,
      correspondencias: [{ ...ativa, ativo: false, status: 'Baixado', motivos: ['serial', 'anydesk'] }],
    })
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText(/mesmo número de série, mesmo AnyDesk/)).toBeInTheDocument()
    expect(screen.getByText('Consta como baixada')).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Atualizar este cadastro' })).not.toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Ver registro' })).toHaveAttribute('href', '/items/123')

    await user.click(screen.getByRole('button', { name: 'Cadastrar como novo mesmo assim' }))
    expect(await screen.findByText('Tela de cadastro')).toBeInTheDocument()
    expect(JSON.parse(screen.getByTestId('estado').textContent ?? 'null').tipoSugerido).toBeNull()
  })

  it('código vencido oferece gerar novo', async () => {
    let posts = 0
    mockCriar(() => {
      posts += 1
    })
    mockAndamento({ status: 'vencido' })
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText('O código venceu.')).toBeInTheDocument()
    expect(posts).toBe(1)
    await user.click(screen.getByRole('button', { name: 'Gerar novo código' }))

    await waitFor(() => expect(posts).toBe(2))
  })

  it('andamento com erro mostra aviso e permite gerar novo código', async () => {
    let posts = 0
    mockCriar(() => {
      posts += 1
    })
    server.use(
      http.get('/api/coletor/:id/andamento', () =>
        HttpResponse.json({ detail: 'Coleta não encontrada.' }, { status: 404 })
      )
    )
    const user = userEvent.setup()
    renderPage()

    expect(await screen.findByText('Não foi possível acompanhar a coleta.')).toBeInTheDocument()
    expect(screen.queryByText('⏳ Aguardando os dados da máquina…')).not.toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Gerar novo código' }))

    await waitFor(() => expect(posts).toBe(2))
  })

  it('continua consultando enquanto aguarda e para depois de receber', async () => {
    mockCriar()
    let chamadas = 0
    server.use(
      http.get('/api/coletor/:id/andamento', () => {
        chamadas += 1
        return HttpResponse.json(
          chamadas === 1
            ? { status: 'aguardando', expira_em: '2026-11-20T15:00:00Z' }
            : {
                status: 'recebido',
                expira_em: '2026-11-20T15:00:00Z',
                campos: CAMPOS,
                tipo_sugerido: 'Notebook',
                correspondencias: [],
              }
        )
      })
    )
    renderPage()

    expect(await screen.findByText('⏳ Aguardando os dados da máquina…')).toBeInTheDocument()
    expect(await screen.findByText('Tela de cadastro', {}, { timeout: 6000 })).toBeInTheDocument()
    expect(chamadas).toBe(2)

    await new Promise((resolve) => setTimeout(resolve, 2500))
    expect(chamadas).toBe(2)
  }, 15000)

  it('gerar novo código troca o comando e consulta o andamento da nova coleta', async () => {
    const COMANDO_2 = COMANDO.replace('abc123', 'novo456')
    const consultados: string[] = []
    let n = 0
    server.use(
      http.post('/api/coletor', () => {
        n += 1
        return HttpResponse.json(
          {
            id: n * 10,
            comando: n === 1 ? COMANDO : COMANDO_2,
            url_arquivo: `/api/coletor/codigo${n}/coletor.cmd`,
            expira_em: '2026-11-20T15:00:00Z',
          },
          { status: 201 }
        )
      }),
      http.get('/api/coletor/:id/andamento', ({ params }) => {
        consultados.push(String(params.id))
        return HttpResponse.json({
          status: params.id === '10' ? 'vencido' : 'aguardando',
          expira_em: '2026-11-20T15:00:00Z',
        })
      })
    )
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: 'Gerar novo código' }))

    expect(await screen.findByDisplayValue(COMANDO_2)).toBeInTheDocument()
    expect(screen.getByRole('link', { name: 'Baixar o coletor' })).toHaveAttribute(
      'href',
      '/api/coletor/codigo2/coletor.cmd'
    )
    await waitFor(() => expect(consultados).toContain('20'))
    expect(consultados[0]).toBe('10')
  })

  it('Sair deste computador encerra a sessão e vai para o login', async () => {
    mockCriar()
    mockAndamento({ status: 'aguardando' })
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: 'Sair deste computador' }))

    expect(await screen.findByText('Tela de login')).toBeInTheDocument()
    expect(logout).toHaveBeenCalled()
  })

  it('se o logout falhar, avisa e continua na página (a sessão segue aberta)', async () => {
    mockCriar()
    mockAndamento({ status: 'aguardando' })
    logout.mockRejectedValueOnce(new Error('500'))
    const user = userEvent.setup()
    renderPage()

    await user.click(await screen.findByRole('button', { name: 'Sair deste computador' }))

    expect(await screen.findByRole('alert')).toHaveTextContent('Não foi possível sair da conta. Tente de novo.')
    expect(screen.queryByText('Tela de login')).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Sair deste computador' })).toBeEnabled()
  })

  it('avisa quando não consegue criar o código', async () => {
    server.use(http.post('/api/coletor', () => HttpResponse.json({ detail: 'Falhou.' }, { status: 500 })))
    renderPage()

    expect(await screen.findByRole('alert')).toBeInTheDocument()
  })
})

describe('lerEstadoColetor', () => {
  const valido = { origem: 'coletor', coletaId: 4, campos: { brand: 'Dell' }, tipoSugerido: null }

  it('aceita o estado válido', () => {
    expect(lerEstadoColetor(valido)).toEqual(valido)
    expect(lerEstadoColetor({ ...valido, tipoSugerido: 'Notebook' })?.tipoSugerido).toBe('Notebook')
  })

  it('recusa estado sem origem coletor ou com forma errada', () => {
    expect(lerEstadoColetor(null)).toBeNull()
    expect(lerEstadoColetor('x')).toBeNull()
    expect(lerEstadoColetor({ ...valido, origem: undefined })).toBeNull()
    expect(lerEstadoColetor({ ...valido, origem: 'outro' })).toBeNull()
    expect(lerEstadoColetor({ ...valido, coletaId: '4' })).toBeNull()
    expect(lerEstadoColetor({ ...valido, campos: { brand: 1 } })).toBeNull()
    expect(lerEstadoColetor({ ...valido, campos: null })).toBeNull()
    expect(lerEstadoColetor({ ...valido, tipoSugerido: 3 })).toBeNull()
  })
})
