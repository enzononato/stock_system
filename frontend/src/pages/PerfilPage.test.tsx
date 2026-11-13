import { describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { getVisibleNavItems } from '@/components/layout/navigation'
import PerfilPage from '@/pages/PerfilPage'
import type { Perfil } from '@/api/perfil'

const tecnico: Perfil = {
  id: 2,
  username: 'joao.silva',
  role: 'Técnico',
  email: 'joao@revalle.com.br',
  totp_ativo: false,
  mfa_obrigatorio: false,
  mfa_disponivel: true,
  codigos_recuperacao_restantes: 0,
}

const CODIGOS = Array.from({ length: 10 }, (_, i) => `abcd-efgh-ijkl-mn${String(i).padStart(2, '2')}`)

function servirPerfil(perfil: Perfil) {
  server.use(http.get('/api/perfil', () => HttpResponse.json(perfil)))
}

function renderPagina() {
  return renderWithClient(
    <MemoryRouter>
      <PerfilPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

describe('PerfilPage', () => {
  it('mostra usuário, função, e-mail e a nota de alteração', async () => {
    servirPerfil(tecnico)
    renderPagina()

    expect(await screen.findByText('joao.silva')).toBeInTheDocument()
    expect(screen.getByText('Técnico')).toBeInTheDocument()
    expect(screen.getByText('joao@revalle.com.br')).toBeInTheDocument()
    expect(screen.getByText('Para alterar o e-mail, fale com um Gestor.')).toBeInTheDocument()
  })

  it('e-mail nulo aparece como "Não cadastrado"', async () => {
    servirPerfil({ ...tecnico, email: null })
    renderPagina()

    expect(await screen.findByText('Não cadastrado')).toBeInTheDocument()
  })

  it('Técnico sem 2FA ativa: QR, código, códigos de recuperação e conclusão', async () => {
    let ativo = false
    server.use(
      http.get('/api/perfil', () =>
        HttpResponse.json({
          ...tecnico,
          totp_ativo: ativo,
          codigos_recuperacao_restantes: ativo ? 10 : 0,
        })
      ),
      http.post('/api/perfil/2fa/iniciar', () =>
        HttpResponse.json({
          segredo: 'JBSWY3DPEHPK3PXP',
          otpauth_uri: 'otpauth://totp/Revalle:joao?secret=JBSWY3DPEHPK3PXP',
        })
      ),
      http.post('/api/perfil/2fa/confirmar', async ({ request }) => {
        expect(await request.json()).toEqual({ codigo: '123456' })
        ativo = true
        return HttpResponse.json({ codigos_recuperacao: CODIGOS })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    expect(await screen.findByText('Inativo')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Ativar verificação em duas etapas' }))
    expect(await screen.findByAltText('QR code para o aplicativo autenticador')).toBeInTheDocument()
    await user.type(screen.getByLabelText('Código de 6 dígitos'), '123456')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    CODIGOS.forEach((c) => expect(screen.getByText(c)).toBeInTheDocument())
    await user.click(screen.getByLabelText('Guardei os códigos em local seguro'))
    await user.click(screen.getByRole('button', { name: 'Concluir' }))

    expect(await screen.findByText('Ativo')).toBeInTheDocument()
    expect(await screen.findByText('Verificação em duas etapas ativada.')).toBeInTheDocument()
    expect(screen.getByText('Códigos de recuperação restantes: 10')).toBeInTheDocument()
  })

  it('Técnico com 2FA ativa desativa confirmando a senha', async () => {
    let corpo: unknown = null
    let ativo = true
    server.use(
      http.get('/api/perfil', () =>
        HttpResponse.json({ ...tecnico, totp_ativo: ativo, codigos_recuperacao_restantes: 8 })
      ),
      http.post('/api/perfil/2fa/desativar', async ({ request }) => {
        corpo = await request.json()
        ativo = false
        return HttpResponse.json({ detail: 'Verificação em duas etapas desativada.' })
      })
    )
    const user = userEvent.setup()
    renderPagina()

    await user.click(await screen.findByRole('button', { name: 'Desativar' }))
    const dialogo = await screen.findByRole('dialog')
    expect(within(dialogo).getByText('Confirme sua senha')).toBeInTheDocument()
    await user.type(within(dialogo).getByLabelText('Senha'), 'minha-senha')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))

    expect(await screen.findByText('Verificação em duas etapas desativada.')).toBeInTheDocument()
    expect(corpo).toEqual({ senha: 'minha-senha' })
    expect(await screen.findByText('Inativo')).toBeInTheDocument()
  })

  it('senha incorreta mostra a mensagem do servidor no diálogo', async () => {
    servirPerfil({ ...tecnico, totp_ativo: true, codigos_recuperacao_restantes: 8 })
    server.use(
      http.post('/api/perfil/2fa/desativar', () =>
        HttpResponse.json({ detail: 'Senha incorreta.' }, { status: 400 })
      )
    )
    const user = userEvent.setup()
    renderPagina()

    await user.click(await screen.findByRole('button', { name: 'Desativar' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Senha'), 'errada')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))

    expect(await within(dialogo).findByText('Senha incorreta.')).toBeInTheDocument()
  })

  it('Gestor com 2FA ativa não pode desativar: obrigatória', async () => {
    servirPerfil({
      ...tecnico,
      role: 'Gestor',
      totp_ativo: true,
      mfa_obrigatorio: true,
      codigos_recuperacao_restantes: 10,
    })
    renderPagina()

    expect(await screen.findByText('Obrigatória para Gestor.')).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Desativar' })).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Gerar novos códigos de recuperação' })).toBeInTheDocument()
  })

  it('servidor sem chave: aviso de indisponível e nenhum botão de ativar', async () => {
    servirPerfil({ ...tecnico, mfa_disponivel: false })
    renderPagina()

    expect(
      await screen.findByText(
        'A verificação em duas etapas está indisponível neste servidor. Avise o administrador.'
      )
    ).toBeInTheDocument()
    expect(
      screen.queryByRole('button', { name: 'Ativar verificação em duas etapas' })
    ).not.toBeInTheDocument()
  })

  it('gerar novos códigos com a senha mostra os 10 códigos devolvidos', async () => {
    servirPerfil({ ...tecnico, totp_ativo: true, codigos_recuperacao_restantes: 1 })
    server.use(
      http.post('/api/perfil/2fa/codigos', () => HttpResponse.json({ codigos_recuperacao: CODIGOS }))
    )
    const user = userEvent.setup()
    renderPagina()

    await user.click(await screen.findByRole('button', { name: 'Gerar novos códigos de recuperação' }))
    const dialogo = await screen.findByRole('dialog')
    await user.type(within(dialogo).getByLabelText('Senha'), 'minha-senha')
    await user.click(within(dialogo).getByRole('button', { name: 'Confirmar' }))

    for (const c of CODIGOS) expect(await screen.findByText(c)).toBeInTheDocument()
  })

  it('a navegação mostra "Meu perfil" para todos os papéis', () => {
    expect(getVisibleNavItems('Jovem Aprendiz').some((i) => i.to === '/perfil')).toBe(true)
    expect(getVisibleNavItems('Técnico').some((i) => i.to === '/perfil')).toBe(true)
    expect(getVisibleNavItems('Gestor').some((i) => i.to === '/perfil')).toBe(true)
  })
})
