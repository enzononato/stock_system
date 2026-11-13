import { afterEach, describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { server } from '@/test/server'
import { setAccessToken } from '@/api/client'
import { AuthProvider } from '@/contexts/AuthContext'
import { ToastContainer } from '@/components/ui/toast'
import LoginPage from '@/pages/LoginPage'

function renderLoginPage() {
  return render(
    <MemoryRouter initialEntries={['/login']}>
      <AuthProvider>
        <Routes>
          <Route path="/login" element={<LoginPage />} />
          <Route path="/" element={<p>Início</p>} />
          <Route path="/esqueci-senha" element={<p>Tela esqueci</p>} />
        </Routes>
      </AuthProvider>
      <ToastContainer />
    </MemoryRouter>
  )
}

const ME = { id: 1, username: 'gestor', role: 'Gestor' }
const semSessao = http.post('/api/auth/refresh', () => new HttpResponse(null, { status: 401 }))
const DETAIL_MUITAS = 'Muitas tentativas. Aguarde um minuto e tente de novo.'

async function preencherSenhaEEntrar(user: ReturnType<typeof userEvent.setup>) {
  await user.type(screen.getByLabelText('Usuário'), 'gestor')
  await user.type(screen.getByLabelText('Senha'), 'senha-certa')
  await user.click(screen.getByRole('button', { name: /Entrar/i }))
}

afterEach(() => setAccessToken(null))

describe('LoginPage', () => {
  it('credencial inválida mostra a mensagem de erro e não deixa a tela em estado de carregando', async () => {
    server.use(
      // AuthProvider tenta um refresh silencioso ao montar — sem sessão prévia, 401 é o esperado.
      semSessao,
      http.post('/api/auth/login', () =>
        HttpResponse.json({ detail: 'Usuário ou senha inválidos.' }, { status: 401 })
      )
    )

    const user = userEvent.setup()
    renderLoginPage()

    await user.type(screen.getByLabelText('Usuário'), 'usuario.teste')
    await user.type(screen.getByLabelText('Senha'), 'senha-errada')

    const botao = screen.getByRole('button', { name: /Entrar/i })
    await user.click(botao)

    expect(await screen.findByText('Usuário ou senha inválidos.')).toBeInTheDocument()

    // Não fica "preso" em carregando: o botão volta ao texto normal e fica habilitado.
    await waitFor(() => expect(screen.getByRole('button', { name: 'Entrar' })).toBeInTheDocument())
    expect(screen.getByRole('button', { name: 'Entrar' })).not.toBeDisabled()
  })

  it('credencial válida navega para fora da tela de login (não mostra erro)', async () => {
    server.use(
      semSessao,
      http.post('/api/auth/login', () =>
        HttpResponse.json({ etapa: 'ok', access_token: 'tok-abc', token_type: 'bearer' })
      ),
      http.get('/api/auth/me', () => HttpResponse.json(ME))
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)

    expect(await screen.findByText('Início')).toBeInTheDocument()
    expect(screen.queryByText('Usuário ou senha inválidos.')).not.toBeInTheDocument()
  })

  it('etapa mfa pede o código e entra', async () => {
    let corpoVerificar: unknown = null
    server.use(
      semSessao,
      http.post('/api/auth/login', () =>
        HttpResponse.json({ etapa: 'mfa', mfa_token: 'mfa-1', expira_em_segundos: 300 })
      ),
      http.post('/api/auth/mfa/verificar', async ({ request }) => {
        corpoVerificar = await request.json()
        return HttpResponse.json({ etapa: 'ok', access_token: 'tok-1', token_type: 'bearer' })
      }),
      http.get('/api/auth/me', () => HttpResponse.json(ME))
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)

    expect(await screen.findByText('Verificação em duas etapas')).toBeInTheDocument()
    expect(
      screen.getByText(
        'Digite o código de 6 dígitos do aplicativo autenticador ou um código de recuperação.'
      )
    ).toBeInTheDocument()
    await user.type(screen.getByLabelText('Código'), '123456')
    await user.click(screen.getByRole('button', { name: 'Verificar' }))

    expect(await screen.findByText('Início')).toBeInTheDocument()
    expect(corpoVerificar).toEqual({ mfa_token: 'mfa-1', codigo: '123456' })
  })

  it('código errado mostra a mensagem do servidor e continua na etapa', async () => {
    server.use(
      semSessao,
      http.post('/api/auth/login', () => HttpResponse.json({ etapa: 'mfa', mfa_token: 'mfa-1' })),
      http.post('/api/auth/mfa/verificar', () =>
        HttpResponse.json({ detail: 'Código inválido.' }, { status: 400 })
      )
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)
    await user.type(await screen.findByLabelText('Código'), '000000')
    await user.click(screen.getByRole('button', { name: 'Verificar' }))

    expect(await screen.findByText('Código inválido.')).toBeInTheDocument()
    expect(screen.getByLabelText('Código')).toBeInTheDocument()
    expect(screen.queryByText('Início')).not.toBeInTheDocument()
  })

  it('voltar retorna ao formulário de senha', async () => {
    server.use(
      semSessao,
      http.post('/api/auth/login', () => HttpResponse.json({ etapa: 'mfa', mfa_token: 'mfa-1' }))
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)
    await user.click(await screen.findByRole('button', { name: 'Voltar' }))

    expect(screen.getByRole('button', { name: 'Entrar' })).toBeInTheDocument()
    expect(screen.getByLabelText('Senha')).toHaveValue('')
  })

  it('poucos códigos de recuperação avisa', async () => {
    server.use(
      semSessao,
      http.post('/api/auth/login', () => HttpResponse.json({ etapa: 'mfa', mfa_token: 'mfa-1' })),
      http.post('/api/auth/mfa/verificar', () =>
        HttpResponse.json({
          etapa: 'ok',
          access_token: 'tok-1',
          token_type: 'bearer',
          codigos_recuperacao_restantes: 2,
        })
      ),
      http.get('/api/auth/me', () => HttpResponse.json(ME))
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)
    await user.type(await screen.findByLabelText('Código'), 'abcd-efgh-ijkl-mnop')
    await user.click(screen.getByRole('button', { name: 'Verificar' }))

    expect(await screen.findByText(/Poucos códigos de recuperação restantes \(2\)/)).toBeInTheDocument()
    expect(await screen.findByText('Início')).toBeInTheDocument()
  })

  it('cadastro obrigatório mostra QR, confirma e exibe os códigos antes de entrar', async () => {
    let chamadasMe = 0
    const codigos = Array.from({ length: 10 }, (_, i) => `aaaa-bbbb-cccc-dd${String(i).padStart(2, '2')}`)
    server.use(
      semSessao,
      http.post('/api/auth/login', () =>
        HttpResponse.json({ etapa: 'cadastro_2fa', mfa_token: 'cad-1', expira_em_segundos: 900 })
      ),
      http.post('/api/auth/mfa/cadastro/iniciar', () =>
        HttpResponse.json({
          segredo: 'JBSWY3DPEHPK3PXP',
          otpauth_uri: 'otpauth://totp/Revalle:gestor?secret=JBSWY3DPEHPK3PXP',
        })
      ),
      http.post('/api/auth/mfa/cadastro/confirmar', () =>
        HttpResponse.json({
          etapa: 'ok',
          access_token: 'tok-9',
          token_type: 'bearer',
          codigos_recuperacao: codigos,
        })
      ),
      http.get('/api/auth/me', () => {
        chamadasMe += 1
        return HttpResponse.json(ME)
      })
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)

    expect(await screen.findByText('Cadastre o autenticador')).toBeInTheDocument()
    expect(await screen.findByAltText('QR code para o aplicativo autenticador')).toBeInTheDocument()
    await user.type(screen.getByLabelText('Código de 6 dígitos'), '123456')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    expect(await screen.findByText('Códigos de recuperação')).toBeInTheDocument()
    codigos.forEach((c) => expect(screen.getByText(c)).toBeInTheDocument())
    expect(chamadasMe).toBe(0)

    await user.click(screen.getByLabelText('Guardei os códigos em local seguro'))
    await user.click(screen.getByRole('button', { name: 'Entrar no sistema' }))

    expect(await screen.findByText('Início')).toBeInTheDocument()
    expect(chamadasMe).toBe(1)
  })

  it('muitas tentativas', async () => {
    server.use(
      semSessao,
      http.post('/api/auth/login', () =>
        HttpResponse.json({ error: 'Rate limit exceeded: 5 per 1 minute' }, { status: 429 })
      )
    )

    const user = userEvent.setup()
    renderLoginPage()
    await preencherSenhaEEntrar(user)

    expect(await screen.findByText(DETAIL_MUITAS)).toBeInTheDocument()
  })

  it('link esqueci minha senha', async () => {
    server.use(semSessao)

    const user = userEvent.setup()
    renderLoginPage()
    await user.click(screen.getByRole('link', { name: 'Esqueci minha senha' }))

    expect(await screen.findByText('Tela esqueci')).toBeInTheDocument()
  })
})
