import { beforeEach, describe, expect, it } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { MemoryRouter } from 'react-router-dom'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import { getVisibleNavItems } from '@/components/layout/navigation'
import PacotesPage from './PacotesPage'

const ad = { id: 1, nome: 'AD', descricao: null, ativo: true }
const email = { id: 2, nome: 'E-mail', descricao: null, ativo: true }
const vpn = { id: 3, nome: 'VPN', descricao: null, ativo: true }
const promax = { id: 9, nome: 'Promax', descricao: null, ativo: false }

function renderPagina() {
  return renderWithClient(
    <MemoryRouter>
      <PacotesPage />
      <ToastContainer />
    </MemoryRouter>
  )
}

beforeEach(() => {
  server.use(
    http.get('/api/colaboradores/cargos', () => HttpResponse.json(['Supervisor', 'Vendedor'])),
    http.get('/api/acessos/pacotes', () =>
      HttpResponse.json([
        { cargo: 'Vendedor', sistemas: [ad, email] },
        { cargo: 'Motorista', sistemas: [promax] },
      ])
    ),
    http.get('/api/acessos/sistemas', () => HttpResponse.json([ad, email, vpn]))
  )
})

describe('PacotesPage', () => {
  it('lista a união de cargos em ordem alfabética, com o resumo de cada pacote', async () => {
    renderPagina()
    await screen.findByRole('button', { name: /Vendedor/ })
    const botoes = screen.getAllByRole('button', { pressed: false })
    expect(botoes.map((b) => b.textContent)).toEqual([
      expect.stringContaining('Motorista'),
      expect.stringContaining('Supervisor'),
      expect.stringContaining('Vendedor'),
    ])
    expect(botoes[0].textContent).toContain('sem colaboradores com este cargo')
    expect(botoes[0].textContent).toContain('1 sistema')
    expect(botoes[1].textContent).toContain('Sem pacote')
    expect(botoes[2].textContent).toContain('2 sistemas')
  })

  it('sem cargo escolhido pede para selecionar um', async () => {
    renderPagina()
    expect(
      await screen.findByText('Selecione um cargo para definir os sistemas padrão.')
    ).toBeInTheDocument()
  })

  it('mostra o pacote do cargo e salva com o sistema marcado', async () => {
    let corpo: unknown = null
    server.use(
      http.put('/api/acessos/pacotes', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ cargo: 'Vendedor', sistemas: [ad, email, vpn] })
      })
    )
    const user = userEvent.setup()
    renderPagina()
    await user.click(await screen.findByRole('button', { name: /Vendedor/ }))
    expect(screen.getByRole('heading', { level: 2, name: 'Vendedor' })).toBeInTheDocument()
    expect(screen.getByRole('checkbox', { name: 'AD' })).toBeChecked()
    expect(screen.getByRole('checkbox', { name: 'E-mail' })).toBeChecked()
    expect(screen.getByRole('checkbox', { name: 'VPN' })).not.toBeChecked()

    await user.click(screen.getByRole('checkbox', { name: 'VPN' }))
    await user.click(screen.getByRole('button', { name: 'Salvar pacote' }))
    await waitFor(() => expect(corpo).toEqual({ cargo: 'Vendedor', sistema_ids: [1, 2, 3] }))
    expect(await screen.findByText('Pacote salvo.')).toBeInTheDocument()
  })

  it('avisa dos sistemas desativados que ainda estão no pacote', async () => {
    const user = userEvent.setup()
    renderPagina()
    await user.click(await screen.findByRole('button', { name: /Motorista/ }))
    expect(
      screen.getByText('Sistemas desativados neste pacote (não geram tarefas): Promax')
    ).toBeInTheDocument()
  })

  it('a busca filtra os cargos', async () => {
    const user = userEvent.setup()
    renderPagina()
    await screen.findByRole('button', { name: /Vendedor/ })
    await user.type(screen.getByLabelText('Buscar cargo'), 'sup')
    expect(screen.getByRole('button', { name: /Supervisor/ })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Vendedor/ })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: /Motorista/ })).not.toBeInTheDocument()
  })

  it('o menu mostra Pacotes por função só para o Gestor', () => {
    expect(getVisibleNavItems('Técnico').map((i) => i.to)).not.toContain('/acessos/pacotes')
    expect(getVisibleNavItems('Gestor').map((i) => i.to)).toContain('/acessos/pacotes')
  })
})
