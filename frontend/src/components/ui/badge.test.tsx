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
  it('mostra travessão quando não há status', () => {
    render(<StatusBadge />)
    expect(screen.getByText('-')).toBeInTheDocument()
  })

  // O selo virou monocromático: a cor da variante não aparece mais na classe
  // do shell, só na do ponto (dot) — por isso a tabela abaixo passou a
  // comparar com a classe de `dotClasses`, não mais com `badge-status-*`.
  it.each([
    ['Disponível', 'bg-[var(--status-available)]'],
    ['Indisponível', 'bg-[var(--status-loaned)]'],
    ['Pendente Devolução', 'bg-[var(--status-return)]'],
    ['Pendente', 'bg-[var(--status-pending)]'],
  ])('mapeia %s para %s', (status, classe) => {
    const { container } = render(<StatusBadge status={status} />)
    expect(container.innerHTML).toContain(classe)
  })

  it('status desconhecido cai no neutro em vez de sumir', () => {
    const { container } = render(<StatusBadge status="Qualquer Coisa" />)
    expect(screen.getByText('Qualquer Coisa')).toBeInTheDocument()
    // Shell é idêntico entre variantes agora, então não dá mais para provar o
    // branch "default" pela classe do selo. O fallback é o único caminho que
    // não passa showDot, então a ausência do ponto colorido comprova que caiu
    // nele (se "Qualquer Coisa" batesse por engano num branch mapeado, um
    // ponto apareceria e este assert quebraria).
    expect(container.querySelector('.rounded-full')).not.toBeInTheDocument()
  })
})
