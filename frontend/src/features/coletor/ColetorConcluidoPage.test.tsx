import { describe, expect, it } from 'vitest'
import { render, screen } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import ColetorConcluidoPage from '@/features/coletor/ColetorConcluidoPage'

const LEMBRETE = 'Entregue ao colaborador e imprima a etiqueta pelo seu computador (Empréstimo e Etiquetas).'

function renderPage(busca: string) {
  return render(
    <MemoryRouter initialEntries={[`/coletor/concluido${busca}`]}>
      <ColetorConcluidoPage />
    </MemoryRouter>
  )
}

describe('ColetorConcluidoPage', () => {
  it('saiu=1 confirma o cadastro e a saída da conta', () => {
    renderPage('?id=42&saiu=1')

    expect(screen.getByText('Máquina cadastrada (#42).')).toBeInTheDocument()
    expect(screen.getByText('Você saiu da sua conta neste computador.')).toBeInTheDocument()
    expect(screen.getByText(LEMBRETE)).toBeInTheDocument()
  })

  it('saiu=0 avisa para sair pelo menu', () => {
    renderPage('?id=42&saiu=0')

    expect(screen.getByText('Máquina cadastrada (#42).')).toBeInTheDocument()
    expect(
      screen.getByText('Não foi possível sair da conta automaticamente. Saia pelo menu antes de entregar a máquina.')
    ).toBeInTheDocument()
    expect(screen.queryByText('Você saiu da sua conta neste computador.')).not.toBeInTheDocument()
    expect(screen.getByText(LEMBRETE)).toBeInTheDocument()
  })

  it('sem id válido não inventa o número', () => {
    renderPage('?saiu=1')

    expect(screen.getByText('Máquina cadastrada.')).toBeInTheDocument()
  })
})
