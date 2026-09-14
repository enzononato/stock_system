import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { PageHeader, PanelHeader } from './PageHeader'

describe('PageHeader', () => {
  it('renderiza o título como <h1>', () => {
    render(<PageHeader title="Histórico de Operações" />)
    const titulo = screen.getByRole('heading', { level: 1, name: 'Histórico de Operações' })
    expect(titulo).toBeInTheDocument()
  })

  it('renderiza o eyebrow quando informado', () => {
    render(<PageHeader title="Título" eyebrow="Auditoria e Rastreabilidade" />)
    expect(screen.getByText('Auditoria e Rastreabilidade')).toBeInTheDocument()
  })

  it('não renderiza a linha de eyebrow quando não informado', () => {
    const { container } = render(<PageHeader title="Título" />)
    expect(container.querySelector('.text-caption')).not.toBeInTheDocument()
  })

  it('renderiza o separador "•" e o eyebrowDetail somente quando eyebrowDetail é informado', () => {
    const { rerender } = render(<PageHeader title="Título" eyebrow="Inventário" />)
    expect(screen.queryByText('•')).not.toBeInTheDocument()

    rerender(<PageHeader title="Título" eyebrow="Inventário" eyebrowDetail="Data Grid Operacional" />)
    expect(screen.getByText('•')).toBeInTheDocument()
    expect(screen.getByText('Data Grid Operacional')).toBeInTheDocument()
  })

  it('renderiza a descrição quando informada', () => {
    render(<PageHeader title="Título" description="Descrição de teste da tela." />)
    expect(screen.getByText('Descrição de teste da tela.')).toBeInTheDocument()
  })

  it('não renderiza descrição quando não informada', () => {
    render(<PageHeader title="Título" />)
    expect(screen.queryByText(/Descrição/)).not.toBeInTheDocument()
  })

  it('renderiza o conteúdo do slot de ações', () => {
    render(<PageHeader title="Título" actions={<button>Nova Ação</button>} />)
    expect(screen.getByRole('button', { name: 'Nova Ação' })).toBeInTheDocument()
  })
})

describe('PanelHeader', () => {
  it('renderiza o título da faixa de legenda do painel', () => {
    render(<PanelHeader title="Filtros de Consulta" />)
    expect(screen.getByText('Filtros de Consulta')).toBeInTheDocument()
  })

  it('renderiza a descrição quando informada', () => {
    render(<PanelHeader title="Filtros de Consulta" description="Refine a lista pelos campos abaixo." />)
    expect(screen.getByText('Refine a lista pelos campos abaixo.')).toBeInTheDocument()
  })

  it('renderiza o conteúdo do slot de ações', () => {
    render(<PanelHeader title="Filtros de Consulta" actions={<button>Limpar</button>} />)
    expect(screen.getByRole('button', { name: 'Limpar' })).toBeInTheDocument()
  })
})
