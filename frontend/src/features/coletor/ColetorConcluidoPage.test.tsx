import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, useLocation } from 'react-router-dom'
import ColetorConcluidoPage from '@/features/coletor/ColetorConcluidoPage'

const logout = vi.fn(() => Promise.resolve())

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({ user: null, isLoading: false, logout }),
}))

function Busca() {
  return <output data-testid="busca">{useLocation().search}</output>
}

const LEMBRETE = 'Entregue ao colaborador e imprima a etiqueta pelo seu computador (Empréstimo e Etiquetas).'

function renderPage(busca: string) {
  return render(
    <MemoryRouter initialEntries={[`/coletor/concluido${busca}`]}>
      <ColetorConcluidoPage />
      <Busca />
    </MemoryRouter>
  )
}

describe('ColetorConcluidoPage', () => {
  beforeEach(() => logout.mockClear())
  afterEach(() => vi.useRealTimers())

  it('saiu=1 confirma o cadastro e a saída da conta', () => {
    renderPage('?id=42&saiu=1')

    expect(screen.getByText('Máquina cadastrada (#42).')).toBeInTheDocument()
    expect(screen.getByText('Você saiu da sua conta neste computador.')).toBeInTheDocument()
    expect(screen.getByText(LEMBRETE)).toBeInTheDocument()
  })

  it('saiu=0 avisa para sair pelo menu', () => {
    renderPage('?id=42&saiu=0')

    expect(screen.getByText('Máquina cadastrada (#42).')).toBeInTheDocument()
    expect(
      screen.getByText('Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina.')
    ).toBeInTheDocument()
    expect(screen.queryByText('Você saiu da sua conta neste computador.')).not.toBeInTheDocument()
    expect(screen.getByText(LEMBRETE)).toBeInTheDocument()
  })

  it('sem saiu o logout ainda está em andamento: mostra o aviso neutro', () => {
    renderPage('?id=42')

    expect(screen.getByText('Máquina cadastrada (#42).')).toBeInTheDocument()
    expect(screen.getByText('Saindo da sua conta…')).toBeInTheDocument()
    expect(screen.queryByText('Você saiu da sua conta neste computador.')).not.toBeInTheDocument()
    expect(screen.queryByText(/Não foi possível sair/)).not.toBeInTheDocument()
  })

  it('sem id válido não inventa o número', () => {
    renderPage('?saiu=1')

    expect(screen.getByText('Máquina cadastrada.')).toBeInTheDocument()
  })

  it('saiu=0: "Tentar sair de novo" sai da conta e troca o endereço para saiu=1', async () => {
    const user = userEvent.setup()
    renderPage('?id=42&saiu=0')

    await user.click(screen.getByRole('button', { name: 'Tentar sair de novo' }))

    expect(logout).toHaveBeenCalledTimes(1)
    expect(await screen.findByText('Você saiu da sua conta neste computador.')).toBeInTheDocument()
    expect(screen.getByTestId('busca')).toHaveTextContent('?id=42&saiu=1')
    expect(screen.queryByRole('button', { name: 'Tentar sair de novo' })).not.toBeInTheDocument()
  })

  it('saiu=0: se o logout falhar de novo, mantém o aviso e diz que ainda não saiu', async () => {
    logout.mockRejectedValueOnce(new Error('500'))
    const user = userEvent.setup()
    renderPage('?id=42&saiu=0')

    await user.click(screen.getByRole('button', { name: 'Tentar sair de novo' }))

    expect(await screen.findByText('Ainda não foi possível sair.')).toBeInTheDocument()
    expect(screen.getByText(/Não foi possível sair da conta automaticamente/)).toBeInTheDocument()
    expect(screen.getByTestId('busca')).toHaveTextContent('?id=42&saiu=0')
    expect(screen.getByRole('button', { name: 'Tentar sair de novo' })).toBeEnabled()
  })

  it('saiu=1 não oferece nova tentativa', () => {
    renderPage('?id=42&saiu=1')

    expect(screen.queryByRole('button', { name: 'Tentar sair de novo' })).not.toBeInTheDocument()
  })

  it('sem saiu por mais de 5 s (página recarregada) oferece a nova tentativa', async () => {
    vi.useFakeTimers()
    renderPage('?id=42')

    expect(screen.queryByRole('button', { name: 'Tentar sair de novo' })).not.toBeInTheDocument()
    await act(async () => {
      await vi.advanceTimersByTimeAsync(4900)
    })
    expect(screen.queryByRole('button', { name: 'Tentar sair de novo' })).not.toBeInTheDocument()
    await act(async () => {
      await vi.advanceTimersByTimeAsync(200)
    })

    expect(screen.getByText('Saindo da sua conta…')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Tentar sair de novo' })).toBeInTheDocument()

    vi.useRealTimers()
    const user = userEvent.setup()
    await user.click(screen.getByRole('button', { name: 'Tentar sair de novo' }))
    expect(await screen.findByText('Você saiu da sua conta neste computador.')).toBeInTheDocument()
    expect(screen.getByTestId('busca')).toHaveTextContent('?id=42&saiu=1')
  })
})
