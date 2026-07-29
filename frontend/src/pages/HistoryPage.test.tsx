import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import HistoryPage from '@/pages/HistoryPage'

// HistoryPage só mostra a coluna/botão de estorno para quem tem o papel
// 'Gestor' (ver hasRole('Gestor') em HistoryPage.tsx). Mocka-se o contexto de
// autenticação em vez de logar de verdade, porque o que este teste caracteriza
// é o fluxo de confirmação de senha do estorno, não o login em si (já coberto
// em LoginPage.test.tsx e client.test.ts).
vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor.teste', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

const entradaCadastro = {
  id: 10,
  item_id: 5,
  operador: 'gestor.teste',
  operation: 'Cadastro',
  data_operacao: '2026-01-10T09:00:00',
}

function mockHistoryList() {
  server.use(
    http.get('/api/history', () =>
      HttpResponse.json({ items: [entradaCadastro], total: 1 })
    )
  )
}

describe('HistoryPage — estorno', () => {
  it('o botão de estornar abre o painel de confirmação pedindo senha, sem chamar a API antes de confirmar', async () => {
    mockHistoryList()
    let chamadasReverse = 0
    server.use(
      http.post('/api/history/10/reverse', () => {
        chamadasReverse += 1
        return HttpResponse.json({ detail: 'ok' })
      })
    )

    const user = userEvent.setup()
    renderWithClient(<HistoryPage />)

    const botaoEstornar = await screen.findByRole('button', { name: /Estornar/i })
    await user.click(botaoEstornar)

    expect(screen.getByText('Confirmar Estorno — Operação #10')).toBeInTheDocument()
    // BUG DE ACESSIBILIDADE (encontrado, não corrigido — fora da posse deste módulo):
    // o <Label>Confirme sua senha</Label> em HistoryPage.tsx não tem `htmlFor`/`id`
    // associando-o ao <Input>, então getByLabelText não encontra a associação
    // programática. Consultamos pelo placeholder como contorno neste teste.
    expect(screen.getByPlaceholderText('Sua senha de acesso')).toBeInTheDocument()
    expect(chamadasReverse).toBe(0)

    // O botão "Confirmar" começa desabilitado sem senha digitada.
    expect(screen.getByRole('button', { name: 'Confirmar' })).toBeDisabled()
  })

  it('só chama a API de estorno depois de confirmar com senha preenchida', async () => {
    mockHistoryList()
    let chamadasReverse = 0
    let senhaRecebida: string | null = null
    server.use(
      http.post('/api/history/10/reverse', async ({ request }) => {
        chamadasReverse += 1
        const body = (await request.json()) as { password: string }
        senhaRecebida = body.password
        return HttpResponse.json({ detail: 'Operação estornada com sucesso.' })
      })
    )

    const user = userEvent.setup()
    renderWithClient(<HistoryPage />)

    await user.click(await screen.findByRole('button', { name: /Estornar/i }))
    await user.type(screen.getByPlaceholderText('Sua senha de acesso'), 'SenhaForte#123')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    await waitFor(() => expect(chamadasReverse).toBe(1))
    expect(senhaRecebida).toBe('SenhaForte#123')

    // Painel fecha depois do sucesso.
    await waitFor(() => expect(screen.queryByText('Confirmar Estorno — Operação #10')).not.toBeInTheDocument())
  })

  it('senha errada (403) mantém o painel aberto mostrando a mensagem de erro', async () => {
    mockHistoryList()
    server.use(
      http.post('/api/history/10/reverse', () =>
        HttpResponse.json({ detail: 'Senha incorreta.' }, { status: 403 })
      )
    )

    const user = userEvent.setup()
    renderWithClient(<HistoryPage />)

    await user.click(await screen.findByRole('button', { name: /Estornar/i }))
    await user.type(screen.getByPlaceholderText('Sua senha de acesso'), 'senha-errada')
    await user.click(screen.getByRole('button', { name: 'Confirmar' }))

    expect(await screen.findByText('Senha incorreta. Ação não autorizada.')).toBeInTheDocument()
    // O painel continua aberto para o usuário tentar de novo.
    expect(screen.getByText('Confirmar Estorno — Operação #10')).toBeInTheDocument()
    expect(screen.getByPlaceholderText('Sua senha de acesso')).toBeInTheDocument()
  })

  it('cancelar fecha o painel sem chamar a API', async () => {
    mockHistoryList()
    let chamadasReverse = 0
    server.use(
      http.post('/api/history/10/reverse', () => {
        chamadasReverse += 1
        return HttpResponse.json({ detail: 'ok' })
      })
    )

    const user = userEvent.setup()
    renderWithClient(<HistoryPage />)

    await user.click(await screen.findByRole('button', { name: /Estornar/i }))
    await user.click(screen.getByRole('button', { name: 'Cancelar' }))

    expect(screen.queryByText('Confirmar Estorno — Operação #10')).not.toBeInTheDocument()
    expect(chamadasReverse).toBe(0)
  })
})
