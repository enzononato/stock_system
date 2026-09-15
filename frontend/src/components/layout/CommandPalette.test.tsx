import { useState } from 'react'
import { describe, it, expect, beforeEach, vi } from 'vitest'
import { fireEvent, render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { CommandPalette } from './CommandPalette'

// A paleta recebe `open`/`onOpenChange` de fora (controlada pelo AppLayout em
// produção) mas registra o próprio listener global de Ctrl+K/Escape — este
// Harness reproduz só o estado que o AppLayout manteria, para poder testar o
// componente isoladamente.
const usuarioMock = { id: 1, username: 'operador.teste', role: 'Gestor' }

vi.mock('@/contexts/AuthContext', () => ({
  useAuth: () => ({
    user: usuarioMock,
    isLoading: false,
    login: vi.fn(),
    logout: vi.fn(),
    hasRole: (...roles: string[]) => roles.includes(usuarioMock.role),
  }),
}))

function Harness({ initialOpen = false }: { initialOpen?: boolean }) {
  const [open, setOpen] = useState(initialOpen)
  return (
    <MemoryRouter>
      <CommandPalette open={open} onOpenChange={setOpen} />
    </MemoryRouter>
  )
}

describe('CommandPalette', () => {
  beforeEach(() => {
    usuarioMock.role = 'Gestor'
  })

  it('Ctrl+K abre a paleta', () => {
    render(<Harness />)
    expect(screen.queryByRole('dialog', { name: 'Paleta de comandos' })).not.toBeInTheDocument()

    fireEvent.keyDown(window, { key: 'k', ctrlKey: true })

    expect(screen.getByRole('dialog', { name: 'Paleta de comandos' })).toBeInTheDocument()
  })

  it('⌘K (metaKey, macOS) também abre a paleta', () => {
    render(<Harness />)

    fireEvent.keyDown(window, { key: 'k', metaKey: true })

    expect(screen.getByRole('dialog', { name: 'Paleta de comandos' })).toBeInTheDocument()
  })

  it('Escape fecha a paleta', () => {
    render(<Harness initialOpen />)
    expect(screen.getByRole('dialog', { name: 'Paleta de comandos' })).toBeInTheDocument()

    fireEvent.keyDown(window, { key: 'Escape' })

    expect(screen.queryByRole('dialog', { name: 'Paleta de comandos' })).not.toBeInTheDocument()
  })

  it('o campo de busca filtra a lista de destinos pelo rótulo', async () => {
    const user = userEvent.setup()
    render(<Harness initialOpen />)

    await user.type(screen.getByRole('textbox', { name: 'Buscar uma tela do sistema' }), 'Periféricos')

    expect(screen.getByRole('option', { name: /Periféricos/ })).toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Dashboard/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Estoque/ })).not.toBeInTheDocument()
  })

  it('mostra a mensagem de vazio quando a busca não encontra nenhum destino', async () => {
    const user = userEvent.setup()
    render(<Harness initialOpen />)

    await user.type(screen.getByRole('textbox', { name: 'Buscar uma tela do sistema' }), 'xyz-inexistente')

    expect(screen.getByText('Nenhum destino encontrado.')).toBeInTheDocument()
  })

  it('não oferece destino fora do papel do usuário logado', () => {
    // Técnico não tem acesso a /remove, /users nem /unidades (roles: ['Gestor']
    // em navigation.ts) — a paleta é alimentada pela mesma lista filtrada por
    // papel que a Sidebar usa, então esses destinos não podem aparecer aqui.
    usuarioMock.role = 'Técnico'
    render(<Harness initialOpen />)

    expect(screen.getByRole('option', { name: /Cadastrar Equipamento/ })).toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Remover \/ Estorno/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Gestão de Usuários/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Unidades de Revenda/ })).not.toBeInTheDocument()
  })

  it('Enter navega para o destino filtrado e fecha a paleta', async () => {
    const user = userEvent.setup()
    render(<Harness initialOpen />)

    await user.type(screen.getByRole('textbox', { name: 'Buscar uma tela do sistema' }), 'Estoque')
    await user.keyboard('{Enter}')

    expect(screen.queryByRole('dialog', { name: 'Paleta de comandos' })).not.toBeInTheDocument()
  })

  it('clicar fora (no overlay) fecha a paleta', () => {
    render(<Harness initialOpen />)
    const dialogo = screen.getByRole('dialog', { name: 'Paleta de comandos' })

    // O overlay é o irmão anterior do próprio diálogo, dentro do mesmo container fixo.
    const overlay = dialogo.previousElementSibling as HTMLElement
    fireEvent.click(overlay)

    expect(screen.queryByRole('dialog', { name: 'Paleta de comandos' })).not.toBeInTheDocument()
  })
})
