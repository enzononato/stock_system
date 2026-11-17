import { beforeEach, describe, expect, it, vi } from 'vitest'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { BotoesExportar } from './BotoesExportar'

const toastMock = vi.hoisted(() => vi.fn())
vi.mock('@/components/ui/toast', () => ({ toast: toastMock }))

beforeEach(() => {
  toastMock.mockReset()
})

describe('BotoesExportar', () => {
  it('expõe um grupo com os botões Excel e PDF', () => {
    render(<BotoesExportar onExportar={vi.fn()} />)
    expect(screen.getByRole('group', { name: 'Exportar' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Exportar em Excel' })).toHaveTextContent('Excel')
    expect(screen.getByRole('button', { name: 'Exportar em PDF' })).toHaveTextContent('PDF')
  })

  it('clicar em PDF chama onExportar("pdf") e mostra aria-busy até resolver', async () => {
    let resolver!: () => void
    const onExportar = vi.fn(
      () =>
        new Promise<void>((r) => {
          resolver = r
        })
    )
    render(<BotoesExportar onExportar={onExportar} />)

    await userEvent.click(screen.getByRole('button', { name: 'Exportar em PDF' }))

    expect(onExportar).toHaveBeenCalledWith('pdf')
    expect(screen.getByRole('button', { name: 'Exportar em PDF' })).toHaveAttribute('aria-busy', 'true')
    expect(screen.getByRole('button', { name: 'Exportar em Excel' })).toBeDisabled()
    expect(screen.getByRole('button', { name: 'Exportar em Excel' })).not.toHaveAttribute('aria-busy', 'true')

    resolver()
    await waitFor(() =>
      expect(screen.getByRole('button', { name: 'Exportar em PDF' })).not.toHaveAttribute('aria-busy', 'true')
    )
    expect(screen.getByRole('button', { name: 'Exportar em Excel' })).toBeEnabled()
  })

  it('erro mostra a mensagem em um toast de erro', async () => {
    render(<BotoesExportar onExportar={vi.fn().mockRejectedValue(new Error('Falhou'))} />)
    await userEvent.click(screen.getByRole('button', { name: 'Exportar em Excel' }))
    await waitFor(() => expect(toastMock).toHaveBeenCalledWith('Falhou', 'error'))
  })

  it('disabled desabilita os dois botões e o rótulo muda os nomes', () => {
    render(<BotoesExportar onExportar={vi.fn()} disabled rotulo="Exportar termos pendentes" />)
    expect(screen.getByRole('group', { name: 'Exportar termos pendentes' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Exportar termos pendentes em Excel' })).toBeDisabled()
    expect(screen.getByRole('button', { name: 'Exportar termos pendentes em PDF' })).toBeDisabled()
  })
})
