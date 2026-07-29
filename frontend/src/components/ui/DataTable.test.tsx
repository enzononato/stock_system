import { describe, expect, it, vi } from 'vitest'
import { render, screen, within } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import type { ColumnDef } from '@tanstack/react-table'
import { DataTable } from '@/components/ui/DataTable'

interface Pessoa {
  id: number
  nome: string
}

const columns: ColumnDef<Pessoa, unknown>[] = [
  { accessorKey: 'id', header: 'ID' },
  { accessorKey: 'nome', header: 'Nome' },
]

const pessoas: Pessoa[] = [
  { id: 1, nome: 'Ana' },
  { id: 2, nome: 'Bruno' },
  { id: 3, nome: 'Carla' },
]

describe('DataTable — sem a prop pagination (retrocompatibilidade)', () => {
  it('mantém busca e contagem no cliente com base em data.length', () => {
    render(<DataTable data={pessoas} columns={columns} />)

    expect(screen.getByText('3 de 3 registros')).toBeInTheDocument()
    expect(screen.getAllByRole('row')).toHaveLength(1 + pessoas.length) // header + linhas

    // Sem controles de paginação server-side.
    expect(screen.queryByText(/Página \d+ de \d+/)).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Anterior/i })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Próxima/i })).not.toBeInTheDocument()
  })

  it('filtra as linhas no cliente ao digitar na busca, e atualiza a contagem', async () => {
    const user = userEvent.setup()
    render(<DataTable data={pessoas} columns={columns} />)

    const busca = screen.getByPlaceholderText('Buscar...')
    await user.type(busca, 'bru')

    expect(screen.getByText('Bruno')).toBeInTheDocument()
    expect(screen.queryByText('Ana')).not.toBeInTheDocument()
    expect(screen.queryByText('Carla')).not.toBeInTheDocument()
    expect(screen.getByText('1 de 3 registros')).toBeInTheDocument()
  })
})

describe('DataTable — com a prop pagination (server-side)', () => {
  it('usa o total do servidor (não data.length) para a contagem e a paginação', () => {
    const onPageChange = vi.fn()
    render(
      <DataTable
        data={[pessoas[0], pessoas[1]]}
        columns={columns}
        pagination={{
          total: 5,
          pageIndex: 0,
          pageSize: 2,
          onPageChange,
        }}
      />
    )

    expect(screen.getByText('1–2 de 5 registros')).toBeInTheDocument()
    expect(screen.getByText('Página 1 de 3')).toBeInTheDocument()

    // "Anterior" desabilitado na primeira página, "Próxima" habilitado.
    expect(screen.getByRole('button', { name: /Anterior/i })).toBeDisabled()
    expect(screen.getByRole('button', { name: /Próxima/i })).not.toBeDisabled()
  })

  it('navega entre páginas chamando onPageChange com o índice certo', async () => {
    const user = userEvent.setup()
    const onPageChange = vi.fn()
    render(
      <DataTable
        data={[pessoas[0], pessoas[1]]}
        columns={columns}
        pagination={{ total: 5, pageIndex: 0, pageSize: 2, onPageChange }}
      />
    )

    await user.click(screen.getByRole('button', { name: /Próxima/i }))
    expect(onPageChange).toHaveBeenCalledWith(1)
  })

  it('desabilita "Próxima" na última página', () => {
    const onPageChange = vi.fn()
    render(
      <DataTable
        data={[pessoas[2]]}
        columns={columns}
        pagination={{ total: 5, pageIndex: 2, pageSize: 2, onPageChange }}
      />
    )

    expect(screen.getByText('Página 3 de 3')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Próxima/i })).toBeDisabled()
    expect(screen.getByRole('button', { name: /Anterior/i })).not.toBeDisabled()
  })

  it('sem onSearchChange, o campo de busca fica oculto (não oferece uma busca que mentiria sobre o alcance)', () => {
    render(
      <DataTable
        data={pessoas}
        columns={columns}
        pagination={{ total: 3, pageIndex: 0, pageSize: 10, onPageChange: vi.fn() }}
      />
    )

    expect(screen.queryByPlaceholderText('Buscar...')).not.toBeInTheDocument()
  })

  it('com onSearchChange, o campo de busca aparece e repassa o termo digitado para o servidor', async () => {
    const user = userEvent.setup()
    const onSearchChange = vi.fn()
    const { rerender } = render(
      <DataTable
        data={pessoas}
        columns={columns}
        searchPlaceholder="Buscar pessoas..."
        pagination={{ total: 3, pageIndex: 0, pageSize: 10, onPageChange: vi.fn(), search: '', onSearchChange }}
      />
    )

    const busca = screen.getByPlaceholderText('Buscar pessoas...')
    await user.type(busca, 'x')
    expect(onSearchChange).toHaveBeenCalledWith('x')

    // O valor exibido é controlado por quem chama (prop `search`), não estado interno.
    rerender(
      <DataTable
        data={pessoas}
        columns={columns}
        searchPlaceholder="Buscar pessoas..."
        pagination={{ total: 3, pageIndex: 0, pageSize: 10, onPageChange: vi.fn(), search: 'controlado', onSearchChange }}
      />
    )
    expect(screen.getByPlaceholderText('Buscar pessoas...')).toHaveValue('controlado')
  })

  it('mostra "0 registros" quando o total do servidor é zero', () => {
    render(
      <DataTable
        data={[]}
        columns={columns}
        pagination={{ total: 0, pageIndex: 0, pageSize: 10, onPageChange: vi.fn() }}
      />
    )

    expect(screen.getByText('0 registros')).toBeInTheDocument()
    expect(within(screen.getByRole('table')).getByText('Nenhum resultado encontrado.')).toBeInTheDocument()
  })
})
