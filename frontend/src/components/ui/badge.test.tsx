import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { Badge, StatusBadge } from './badge'

describe('Badge', () => {
  it('renderiza o conteúdo', () => {
    render(<Badge>Olá</Badge>)
    expect(screen.getByText('Olá')).toBeInTheDocument()
  })

  it('mantém as variantes usadas pelas páginas', () => {
    for (const v of ['default', 'success', 'warning', 'danger', 'info', 'purple'] as const) {
      const { unmount } = render(<Badge variant={v}>{v}</Badge>)
      expect(screen.getByText(v)).toBeInTheDocument()
      unmount()
    }
  })

  it('não usa mais a paleta antiga do Tailwind', () => {
    const { container } = render(<Badge variant="success">ok</Badge>)
    expect(container.innerHTML).not.toMatch(/emerald-|rose-|amber-|sky-|purple-|indigo-|slate-/)
  })

  it('usa token de status semântico no variant success', () => {
    const { container } = render(<Badge variant="success">ok</Badge>)
    expect(container.innerHTML).toContain('badge-status-available')
  })
})

describe('StatusBadge', () => {
  it('mostra travessão quando não há status', () => {
    render(<StatusBadge />)
    expect(screen.getByText('-')).toBeInTheDocument()
  })

  it.each([
    ['Disponível', 'badge-status-available'],
    ['Indisponível', 'badge-status-loaned'],
    ['Pendente Devolução', 'badge-status-return'],
    ['Pendente', 'badge-status-pending'],
  ])('mapeia %s para %s', (status, classe) => {
    const { container } = render(<StatusBadge status={status} />)
    expect(container.innerHTML).toContain(classe)
  })

  it('status desconhecido cai no neutro em vez de sumir', () => {
    const { container } = render(<StatusBadge status="Qualquer Coisa" />)
    expect(screen.getByText('Qualquer Coisa')).toBeInTheDocument()
    expect(container.innerHTML).toContain('badge-status-neutral')
  })
})
