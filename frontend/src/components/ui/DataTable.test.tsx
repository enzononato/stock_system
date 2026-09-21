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

/** Gera N pessoas nomeadas "Pessoa 1", "Pessoa 2"... para testar paginação client-side. */
function gerarPessoas(quantidade: number): Pessoa[] {
  return Array.from({ length: quantidade }, (_, i) => ({ id: i + 1, nome: `Pessoa ${i + 1}` }))
}

describe('DataTable — sem a prop pagination (retrocompatibilidade)', () => {
  it('mantém busca e contagem no cliente com base em data.length', () => {
    render(<DataTable data={pessoas} columns={columns} />)

    // O restyle passou os dígitos para <span className="num">, então o texto some do textContent direto do nó — usa matcher de função comparando o textContent completo do <p>.
    expect(
      screen.getByText((_, el) => el instanceof HTMLParagraphElement && el.textContent === '3 de 3 registros')
    ).toBeInTheDocument()
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
    // O restyle passou os dígitos para <span className="num">, então o texto some do textContent direto do nó — usa matcher de função comparando o textContent completo do <p>.
    expect(
      screen.getByText((_, el) => el instanceof HTMLParagraphElement && el.textContent === '1 de 3 registros')
    ).toBeInTheDocument()
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

    // O restyle passou os dígitos para <span className="num">, então o texto some do textContent direto do nó — usa matcher de função comparando o textContent completo do <p>.
    expect(
      screen.getByText((_, el) => el instanceof HTMLParagraphElement && el.textContent === '1–2 de 5 registros')
    ).toBeInTheDocument()
    // Idem para o indicador de página, que também ganhou <span className="num"> nos dígitos.
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 1 de 3')
    ).toBeInTheDocument()

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

    // O restyle passou os dígitos para <span className="num">, então o texto some do textContent direto do nó — usa matcher de função comparando o textContent completo do <span>.
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 3 de 3')
    ).toBeInTheDocument()
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

    // O restyle passou o dígito para <span className="num">, então o texto some do textContent direto do nó — usa matcher de função comparando o textContent completo do <p>.
    expect(
      screen.getByText((_, el) => el instanceof HTMLParagraphElement && el.textContent === '0 registros')
    ).toBeInTheDocument()
    expect(within(screen.getByRole('table')).getByText('Nenhum resultado encontrado.')).toBeInTheDocument()
  })
})

describe('DataTable — paginação client-side (sem a prop pagination)', () => {
  it('não mostra rodapé de navegação quando os dados cabem em uma única página', () => {
    render(<DataTable data={pessoas} columns={columns} />)

    expect(screen.queryByText(/Página \d+ de \d+/)).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Anterior/i })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Próxima/i })).not.toBeInTheDocument()
  })

  it('mostra o rodapé de navegação quando há mais de uma página (padrão de 10 por página)', () => {
    render(<DataTable data={gerarPessoas(15)} columns={columns} />)

    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 1 de 2')
    ).toBeInTheDocument()
    // Só as 10 primeiras linhas da página 1 aparecem.
    expect(screen.getByText('Pessoa 1')).toBeInTheDocument()
    expect(screen.getByText('Pessoa 10')).toBeInTheDocument()
    expect(screen.queryByText('Pessoa 11')).not.toBeInTheDocument()
  })

  it('navegar para a próxima página troca as linhas exibidas', async () => {
    const user = userEvent.setup()
    render(<DataTable data={gerarPessoas(15)} columns={columns} />)

    await user.click(screen.getByRole('button', { name: /Próxima/i }))

    expect(screen.queryByText('Pessoa 1')).not.toBeInTheDocument()
    expect(screen.getByText('Pessoa 11')).toBeInTheDocument()
    expect(screen.getByText('Pessoa 15')).toBeInTheDocument()
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 2 de 2')
    ).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Próxima/i })).toBeDisabled()
  })

  it('buscar reseta a paginação para a primeira página', async () => {
    const user = userEvent.setup()
    render(<DataTable data={gerarPessoas(15)} columns={columns} />)

    await user.click(screen.getByRole('button', { name: /Próxima/i }))
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 2 de 2')
    ).toBeInTheDocument()
    expect(screen.getByText('Pessoa 11')).toBeInTheDocument()

    // Busca por um termo que continua batendo em todos os 15 registros (2
    // páginas) — o que importa aqui é que a digitação volta para a página 1,
    // não o resultado do filtro em si.
    const busca = screen.getByPlaceholderText('Buscar...')
    await user.type(busca, 'Pessoa')

    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 1 de 2')
    ).toBeInTheDocument()
    expect(screen.getByText('Pessoa 1')).toBeInTheDocument()
    expect(screen.queryByText('Pessoa 11')).not.toBeInTheDocument()
  })

  it('não deixa a tabela vazia quando os dados encolhem e a página atual deixa de existir', async () => {
    const user = userEvent.setup()
    const { rerender } = render(<DataTable data={gerarPessoas(25)} columns={columns} />)

    // 25 registros / 10 por página = 3 páginas. Navega até a última página,
    // simulando o usuário parado nela antes dos dados encolherem.
    await user.click(screen.getByRole('button', { name: /Próxima/i }))
    await user.click(screen.getByRole('button', { name: /Próxima/i }))
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 3 de 3')
    ).toBeInTheDocument()

    // Os dados encolhem por fora (ex.: filtro externo, recarregamento) para
    // caber em só 2 páginas — a página 3 deixou de existir.
    rerender(<DataTable data={gerarPessoas(12)} columns={columns} />)

    // A tabela não pode ficar vazia: volta para a página 1 automaticamente.
    expect(
      screen.getByText((_, el) => el instanceof HTMLSpanElement && el.textContent === 'Página 1 de 2')
    ).toBeInTheDocument()
    expect(screen.getByText('Pessoa 1')).toBeInTheDocument()
    expect(screen.queryByText('Nenhum resultado encontrado.')).not.toBeInTheDocument()
  })
})
