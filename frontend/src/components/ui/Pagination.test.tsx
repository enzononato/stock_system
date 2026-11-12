import { describe, it, expect, vi } from 'vitest'
import { render, screen, fireEvent } from '@testing-library/react'
import { Pagination } from './Pagination'

function setup(props: Partial<React.ComponentProps<typeof Pagination>> = {}) {
  const onPageChange = vi.fn()
  const onPageSizeChange = vi.fn()
  render(
    <Pagination
      page={2}
      pageSize={25}
      total={120}
      onPageChange={onPageChange}
      onPageSizeChange={onPageSizeChange}
      {...props}
    />
  )
  return { onPageChange, onPageSizeChange }
}

describe('Pagination', () => {
  it('mostra o intervalo', () => {
    setup()
    expect(screen.getByText('Mostrando 26–50 de 120')).toBeInTheDocument()
  })

  it('última página limita o fim ao total', () => {
    setup({ page: 5 })
    expect(screen.getByText('Mostrando 101–120 de 120')).toBeInTheDocument()
  })

  it('total zero mostra 0–0', () => {
    setup({ page: 1, total: 0 })
    expect(screen.getByText('Mostrando 0–0 de 0')).toBeInTheDocument()
  })

  it('Próxima fica desabilitada na última página', () => {
    setup({ page: 5 })
    expect(screen.getByRole('button', { name: 'Próxima' })).toBeDisabled()
    expect(screen.getByRole('button', { name: 'Anterior' })).toBeEnabled()
  })

  it('Anterior fica desabilitada na primeira página', () => {
    setup({ page: 1 })
    expect(screen.getByRole('button', { name: 'Anterior' })).toBeDisabled()
  })

  it('Anterior e Próxima navegam', () => {
    const { onPageChange } = setup()
    fireEvent.click(screen.getByRole('button', { name: 'Próxima' }))
    expect(onPageChange).toHaveBeenLastCalledWith(3)
    fireEvent.click(screen.getByRole('button', { name: 'Anterior' }))
    expect(onPageChange).toHaveBeenLastCalledWith(1)
  })

  it('Ir para 999 vai para a última página (Enter)', () => {
    const { onPageChange } = setup()
    const input = screen.getByLabelText('Ir para a página')
    fireEvent.change(input, { target: { value: '999' } })
    fireEvent.keyDown(input, { key: 'Enter' })
    expect(onPageChange).toHaveBeenCalledWith(5)
  })

  it('Ir para 0 vai para a primeira página (botão Ir)', () => {
    const { onPageChange } = setup()
    fireEvent.change(screen.getByLabelText('Ir para a página'), { target: { value: '0' } })
    fireEvent.click(screen.getByRole('button', { name: 'Ir' }))
    expect(onPageChange).toHaveBeenCalledWith(1)
  })

  it('texto inválido não chama onPageChange', () => {
    const { onPageChange } = setup()
    const input = screen.getByLabelText('Ir para a página')
    fireEvent.change(input, { target: { value: 'abc' } })
    fireEvent.keyDown(input, { key: 'Enter' })
    fireEvent.click(screen.getByRole('button', { name: 'Ir' }))
    expect(onPageChange).not.toHaveBeenCalled()
  })

  it('troca de tamanho chama onPageSizeChange(50)', () => {
    const { onPageSizeChange } = setup()
    fireEvent.change(screen.getByLabelText('Itens por página'), { target: { value: '50' } })
    expect(onPageSizeChange).toHaveBeenCalledWith(50)
  })

  it('sem onPageSizeChange não mostra o seletor', () => {
    setup({ onPageSizeChange: undefined })
    expect(screen.queryByLabelText('Itens por página')).not.toBeInTheDocument()
  })

  it('isFetching marca a região como ocupada', () => {
    setup({ isFetching: true })
    expect(screen.getByRole('navigation', { name: 'Paginação' })).toHaveAttribute('aria-busy', 'true')
  })
})
