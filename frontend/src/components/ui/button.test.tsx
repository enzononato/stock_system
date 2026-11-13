import { describe, it, expect, vi } from 'vitest'
import { render, screen, fireEvent } from '@testing-library/react'
import { Button } from './button'

describe('Button', () => {
  it('dispara onClick normalmente', () => {
    const onClick = vi.fn()
    render(<Button onClick={onClick}>Salvar</Button>)
    fireEvent.click(screen.getByRole('button', { name: 'Salvar' }))
    expect(onClick).toHaveBeenCalledTimes(1)
    expect(screen.getByRole('button')).not.toHaveAttribute('aria-busy')
  })

  it('loading: aria-busy, desabilitado e sem clique', () => {
    const onClick = vi.fn()
    render(<Button loading onClick={onClick}>Salvar</Button>)
    const btn = screen.getByRole('button', { name: 'Salvar' })
    expect(btn).toHaveAttribute('aria-busy', 'true')
    expect(btn).toBeDisabled()
    fireEvent.click(btn)
    expect(onClick).not.toHaveBeenCalled()
  })

  it('loading mostra spinner que respeita motion-reduce', () => {
    const { container } = render(<Button loading>Salvar</Button>)
    const svg = container.querySelector('svg')
    expect(svg).toBeInTheDocument()
    expect(svg).toHaveClass('animate-spin', 'motion-reduce:animate-none')
    expect(svg).toHaveAttribute('aria-hidden', 'true')
  })

  it('sem loading não mostra spinner', () => {
    const { container } = render(<Button>Salvar</Button>)
    expect(container.querySelector('svg')).not.toBeInTheDocument()
  })

  it('funciona com asChild sem quebrar', () => {
    render(<Button asChild loading><a href="/x">Ir</a></Button>)
    expect(screen.getByRole('link', { name: 'Ir' })).toBeInTheDocument()
  })

  it('aria-busy do consumidor não sobrescreve o loading', () => {
    render(<Button loading aria-busy={false}>Salvar</Button>)
    expect(screen.getByRole('button', { name: 'Salvar' })).toHaveAttribute('aria-busy', 'true')
  })

  it('asChild + loading: o link não chama onClick e a navegação é cancelada', () => {
    const onClick = vi.fn()
    render(<Button asChild loading onClick={onClick}><a href="/x">Ir</a></Button>)
    const link = screen.getByRole('link', { name: 'Ir' })
    expect(fireEvent.click(link)).toBe(false)
    expect(onClick).not.toHaveBeenCalled()
    expect(link).toHaveAttribute('aria-busy', 'true')
  })
})
