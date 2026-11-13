import { describe, it, expect, vi, afterEach } from 'vitest'
import { fireEvent, render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { ThemeProvider } from '@/lib/theme'
import AppLayout from './AppLayout'

let liberarLogout!: () => void
const pendente = new Promise<void>((r) => {
  liberarLogout = r
})
const logout = vi.fn(() => pendente)

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: { id: 1, username: 'gestor.teste', role: 'Gestor' },
    isLoading: false,
    login: vi.fn(),
    logout,
    hasRole: (...roles: string[]) => roles.includes('Gestor'),
  }),
}))

function renderLayout() {
  return render(
    <ThemeProvider>
      <MemoryRouter initialEntries={['/stock']}>
        <Routes>
          <Route element={<AppLayout />}>
            <Route path="/stock" element={<p>conteúdo</p>} />
          </Route>
          <Route path="/login" element={<p>login</p>} />
        </Routes>
      </MemoryRouter>
    </ThemeProvider>
  )
}

describe('AppLayout', () => {
  afterEach(() => {
    liberarLogout()
  })

  it('item ativo da navegação usa o destaque verde-limão', () => {
    renderLayout()
    expect(screen.getAllByRole('link', { name: 'Estoque' })[0]).toHaveClass('bg-lime', 'text-lime-foreground')
  })

  it('gaveta abre como diálogo e fecha pelo botão', () => {
    renderLayout()
    fireEvent.click(screen.getByRole('button', { name: 'Abrir menu de navegação' }))
    expect(screen.getByRole('dialog', { name: 'Menu de navegação' })).toBeInTheDocument()
    fireEvent.click(screen.getByRole('button', { name: 'Fechar menu' }))
    expect(screen.queryByRole('dialog', { name: 'Menu de navegação' })).not.toBeInTheDocument()
  })

  it('gaveta: o foco vai para "Fechar menu" ao abrir e volta ao botão do menu ao fechar (Escape e botão)', async () => {
    renderLayout()
    const user = userEvent.setup()
    const abrir = screen.getByRole('button', { name: 'Abrir menu de navegação' })
    await user.click(abrir)
    expect(screen.getByRole('button', { name: 'Fechar menu' })).toHaveFocus()
    await user.keyboard('{Escape}')
    expect(screen.queryByRole('dialog', { name: 'Menu de navegação' })).not.toBeInTheDocument()
    expect(abrir).toHaveFocus()

    await user.click(abrir)
    await user.click(screen.getByRole('button', { name: 'Fechar menu' }))
    expect(abrir).toHaveFocus()
  })

  it('botão de busca abre a paleta de comandos', () => {
    renderLayout()
    fireEvent.click(screen.getByRole('button', { name: 'Buscar tela (Ctrl+K)' }))
    expect(screen.getByRole('dialog', { name: 'Paleta de comandos' })).toBeInTheDocument()
  })

  it('sair com duplo clique chama logout uma vez', () => {
    renderLayout()
    const sair = screen.getByRole('button', { name: 'Sair do sistema' })
    fireEvent.click(sair)
    fireEvent.click(sair)
    expect(logout).toHaveBeenCalledTimes(1)
    expect(sair).toHaveAttribute('aria-busy', 'true')
  })
})
