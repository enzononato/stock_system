import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { StatusPill } from './StatusPill'

describe('StatusPill', () => {
  it('mostra o texto em negrito', () => {
    render(<StatusPill status="Disponível" />)
    const el = screen.getByText('Disponível')
    expect(el).toHaveClass('font-bold', 'rounded-full')
  })

  it('usa o tom derivado do status', () => {
    render(<StatusPill status="Em Manutenção" />)
    const el = screen.getByText('Em Manutenção')
    expect(el.style.color).toBe('var(--status-maintenance)')
    expect(el.style.backgroundColor).toBe('var(--status-maintenance-bg)')
  })

  it('tone explícito sobrescreve o derivado', () => {
    render(<StatusPill status="Disponível" tone="danger" />)
    expect(screen.getByText('Disponível').style.color).toBe('var(--status-danger)')
  })

  it('status vazio mostra travessão neutro', () => {
    render(<StatusPill status="" />)
    const el = screen.getByText('—')
    expect(el.style.color).toBe('var(--status-neutral)')
  })

  it('aceita className', () => {
    render(<StatusPill status="Ativo" className="ml-2" />)
    expect(screen.getByText('Ativo')).toHaveClass('ml-2')
  })
})
