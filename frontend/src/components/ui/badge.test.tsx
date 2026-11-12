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
    // O selo virou monocromático: a cor só sobrevive no ponto (dot), não mais
    // no shell — por isso precisa de showDot para inspecionar a cor aqui.
    const { container } = render(<Badge variant="success" showDot>ok</Badge>)
    expect(container.innerHTML).toContain('bg-[var(--status-available)]')
  })
})

describe('StatusBadge', () => {
  // StatusBadge agora renderiza StatusPill: o selo é uma pílula colorida
  // (texto no tom, fundo no tom -bg) em vez do selo monocromático com ponto.
  it('mostra travessão quando não há status', () => {
    render(<StatusBadge />)
    expect(screen.getByText('—')).toBeInTheDocument()
  })

  it.each([
    ['Disponível', 'var(--status-available)'],
    ['Indisponível', 'var(--status-loaned)'],
    ['Pendente Devolução', 'var(--status-return)'],
    ['Pendente', 'var(--status-pending)'],
  ])('mapeia %s para a cor %s', (status, cor) => {
    render(<StatusBadge status={status} />)
    expect(screen.getByText(status).style.color).toBe(cor)
  })

  it('status desconhecido cai no neutro em vez de sumir', () => {
    render(<StatusBadge status="Qualquer Coisa" />)
    expect(screen.getByText('Qualquer Coisa').style.color).toBe('var(--status-neutral)')
  })
})
