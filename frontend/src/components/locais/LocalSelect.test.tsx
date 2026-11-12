import { describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { LocalSelect } from '@/components/locais/LocalSelect'

const locais = [
  { id: 10, unidade_id: 1, unidade_nome: 'Juazeiro', nome: 'Almoxarifado', descricao: null, ativo: true },
  { id: 11, unidade_id: 1, unidade_nome: 'Juazeiro', nome: 'Sala TI', descricao: null, ativo: false },
]

describe('LocalSelect', () => {
  it('fica desabilitado sem unidade e não consulta a API', () => {
    // Sem handler: uma requisição não tratada derrubaria o teste.
    renderWithClient(<LocalSelect unidadeId={null} value={null} onChange={() => {}} />)
    expect(screen.getByLabelText('Local')).toBeDisabled()
    expect(screen.getByRole('option', { name: 'Sem local' })).toBeInTheDocument()
  })

  it('oferece "Sem local" e os locais ativos da unidade', async () => {
    server.use(http.get('/api/locais', () => HttpResponse.json(locais)))
    renderWithClient(<LocalSelect unidadeId={1} value={null} onChange={() => {}} />)

    expect(await screen.findByRole('option', { name: 'Almoxarifado' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'Sem local' })).toBeInTheDocument()
    expect(screen.queryByRole('option', { name: /Sala TI/ })).not.toBeInTheDocument()
    expect(screen.getByLabelText('Local')).toBeEnabled()
  })

  it('mantém visível o local inativo que já está selecionado', async () => {
    server.use(http.get('/api/locais', () => HttpResponse.json(locais)))
    renderWithClient(<LocalSelect unidadeId={1} value={11} onChange={() => {}} />)
    expect(await screen.findByRole('option', { name: 'Sala TI (inativo)' })).toBeInTheDocument()
  })

  it('chama onChange com o id escolhido e com null em "Sem local"', async () => {
    server.use(http.get('/api/locais', () => HttpResponse.json(locais)))
    const onChange = vi.fn()
    const user = userEvent.setup()
    renderWithClient(<LocalSelect unidadeId={1} value={null} onChange={onChange} />)
    await screen.findByRole('option', { name: 'Almoxarifado' })

    await user.selectOptions(screen.getByLabelText('Local'), 'Almoxarifado')
    expect(onChange).toHaveBeenLastCalledWith(10)
  })

  it('devolve null ao escolher "Sem local"', async () => {
    server.use(http.get('/api/locais', () => HttpResponse.json(locais)))
    const onChange = vi.fn()
    const user = userEvent.setup()
    renderWithClient(<LocalSelect unidadeId={1} value={10} onChange={onChange} />)
    await screen.findByRole('option', { name: 'Almoxarifado' })
    await waitFor(() => expect(screen.getByLabelText('Local')).toHaveValue('10'))

    await user.selectOptions(screen.getByLabelText('Local'), 'Sem local')
    expect(onChange).toHaveBeenLastCalledWith(null)
  })

  it('respeita disabled', async () => {
    server.use(http.get('/api/locais', () => HttpResponse.json(locais)))
    renderWithClient(<LocalSelect unidadeId={1} value={null} onChange={() => {}} disabled />)
    expect(screen.getByLabelText('Local')).toBeDisabled()
  })
})
