import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { StatBlock } from './StatBlock'

describe('StatBlock', () => {
  it('renderiza label e value', () => {
    render(<StatBlock label="Total em Estoque" value={42} />)
    expect(screen.getByText('Total em Estoque')).toBeInTheDocument()
    expect(screen.getByText('42')).toBeInTheDocument()
  })

  it('renderiza o hint quando informado', () => {
    render(<StatBlock label="Empréstimos" value={10} hint="no mês" />)
    expect(screen.getByText('no mês')).toBeInTheDocument()
  })

  it('não renderiza hint quando não informado', () => {
    const { container } = render(<StatBlock label="Empréstimos" value={10} />)
    expect(container.textContent).toBe('Empréstimos10')
  })

  it('renderiza o símbolo mono quando informado', () => {
    render(<StatBlock label="Cadastros" value={7} symbol="#" />)
    expect(screen.getByText('#')).toBeInTheDocument()
  })

  it('não renderiza símbolo quando não informado', () => {
    const { container } = render(<StatBlock label="Cadastros" value={7} />)
    expect(container.querySelector('.font-mono.text-body-lg')).not.toBeInTheDocument()
  })
})
