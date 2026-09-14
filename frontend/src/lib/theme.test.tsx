import { describe, it, expect, beforeEach } from 'vitest'
import { render, screen, act } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ThemeProvider, useTheme } from './theme'

function Sonda() {
  const { theme, toggle } = useTheme()
  return (
    <div>
      <span data-testid="tema">{theme}</span>
      <button onClick={toggle}>alternar</button>
    </div>
  )
}

describe('ThemeProvider', () => {
  beforeEach(() => {
    localStorage.clear()
    document.documentElement.classList.remove('dark')
  })

  it('começa no modo claro quando não há preferência salva', () => {
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('light')
    expect(document.documentElement).not.toHaveClass('dark')
  })

  it('alterna para escuro e aplica a classe no html', async () => {
    const user = userEvent.setup()
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    await user.click(screen.getByRole('button', { name: 'alternar' }))
    expect(screen.getByTestId('tema')).toHaveTextContent('dark')
    expect(document.documentElement).toHaveClass('dark')
  })

  it('persiste a escolha no localStorage', async () => {
    const user = userEvent.setup()
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    await user.click(screen.getByRole('button', { name: 'alternar' }))
    expect(localStorage.getItem('tema')).toBe('dark')
  })

  it('restaura a preferência salva ao montar', () => {
    localStorage.setItem('tema', 'dark')
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('dark')
    expect(document.documentElement).toHaveClass('dark')
  })

  it('ignora valor inválido no localStorage e cai no claro', () => {
    localStorage.setItem('tema', 'roxo')
    render(<ThemeProvider><Sonda /></ThemeProvider>)
    expect(screen.getByTestId('tema')).toHaveTextContent('light')
  })

  it('useTheme fora do provider lança erro explicativo', () => {
    expect(() => render(<Sonda />)).toThrow(/ThemeProvider/)
  })
})
