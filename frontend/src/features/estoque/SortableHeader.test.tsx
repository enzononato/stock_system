import { describe, expect, it, vi } from 'vitest'
import { fireEvent, render, screen } from '@testing-library/react'
import { SortableHeader } from './SortableHeader'

function renderHeader(props: Partial<React.ComponentProps<typeof SortableHeader>> = {}) {
  const onSort = vi.fn()
  render(
    <table>
      <thead>
        <tr>
          <SortableHeader label="Tipo" coluna="tipo" sort="id" dir="desc" onSort={onSort} {...props} />
        </tr>
      </thead>
    </table>
  )
  return { onSort }
}

describe('SortableHeader', () => {
  it('coluna ativa em asc expõe aria-sort="ascending"', () => {
    renderHeader({ sort: 'tipo', dir: 'asc' })
    expect(screen.getByRole('columnheader', { name: /Tipo/ })).toHaveAttribute('aria-sort', 'ascending')
  })

  it('coluna ativa em desc expõe aria-sort="descending"', () => {
    renderHeader({ sort: 'tipo', dir: 'desc' })
    expect(screen.getByRole('columnheader', { name: /Tipo/ })).toHaveAttribute('aria-sort', 'descending')
  })

  it('coluna inativa expõe aria-sort="none"', () => {
    renderHeader()
    expect(screen.getByRole('columnheader', { name: /Tipo/ })).toHaveAttribute('aria-sort', 'none')
  })

  it('clicar o botão chama onSort com a coluna', () => {
    const { onSort } = renderHeader()
    fireEvent.click(screen.getByRole('button', { name: 'Tipo' }))
    expect(onSort).toHaveBeenCalledWith('tipo')
  })
})
