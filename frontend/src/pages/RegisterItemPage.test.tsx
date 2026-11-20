import { beforeAll, describe, expect, it, vi } from 'vitest'
import { http, HttpResponse } from 'msw'
import { screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import userEvent from '@testing-library/user-event'
import { server } from '@/test/server'
import { renderWithClient } from '@/test/render'
import { ToastContainer } from '@/components/ui/toast'
import RegisterItemPage from '@/pages/RegisterItemPage'

// Os testes de cadastro digitam vários campos e abrem selects do Radix: sob carga (suíte inteira em paralelo) passam de 5 s.
vi.setConfig({ testTimeout: 20000 })

// Radix Select usa a Pointer Events API, que o jsdom não implementa.
beforeAll(() => {
  Element.prototype.hasPointerCapture = Element.prototype.hasPointerCapture ?? (() => false)
  Element.prototype.setPointerCapture = Element.prototype.setPointerCapture ?? (() => {})
  Element.prototype.releasePointerCapture = Element.prototype.releasePointerCapture ?? (() => {})
  Element.prototype.scrollIntoView = Element.prototype.scrollIntoView ?? (() => {})
})

const unidades = [
  { id: 1, nome: 'Revalle Juazeiro', razao_social: 'R1', cnpj: '1', is_active: true },
  { id: 2, nome: 'Revalle Petrolina', razao_social: 'R2', cnpj: '2', is_active: true },
]

const itemExistente = {
  id: 9,
  tipo: 'Notebook',
  brand: 'Dell',
  model: 'Latitude',
  revenda: 'Revalle Juazeiro',
  status: 'Disponível',
  date_registered: '2026-01-05T00:00:00',
  valor_compra: '4599.90',
  data_compra: '2026-01-10',
  garantia_ate: '2027-01-10',
  nf_anexo_key: null,
  local_id: 5,
}

function mockBase() {
  server.use(
    http.get('/api/constants', () =>
      HttpResponse.json({
        center_costs: [],
        revendas: ['Revalle Juazeiro', 'Revalle Petrolina'],
        setores: [],
        equipment_types: ['Notebook'],
        peripheral_types: [],
        removal_reasons: [],
        removal_reasons_attachment: {},
      })
    ),
    http.get('/api/unidades', () => HttpResponse.json(unidades)),
    http.get('/api/locais', ({ request }) => {
      const unidadeId = new URL(request.url).searchParams.get('unidade_id')
      return HttpResponse.json(
        unidadeId === '1' ? [{ id: 5, unidade_id: 1, nome: 'Sala TI', ativo: true }] : []
      )
    })
  )
}

function renderPage(inicial: string) {
  return renderWithClient(
    <MemoryRouter initialEntries={[inicial]}>
      <Routes>
        <Route path="/register" element={<RegisterItemPage mode="create" />} />
        <Route path="/edit/:id" element={<RegisterItemPage mode="edit" />} />
        <Route path="/stock" element={<p>Estoque</p>} />
      </Routes>
      <ToastContainer />
    </MemoryRouter>
  )
}

function arquivoPdf(tamanho?: number) {
  const file = new File(['%PDF-1.4 nf'], 'nf.pdf', { type: 'application/pdf' })
  if (tamanho !== undefined) Object.defineProperty(file, 'size', { value: tamanho })
  return file
}

async function preencherCadastro(user: ReturnType<typeof userEvent.setup>) {
  await user.click(await screen.findByRole('combobox', { name: /Tipo/ }))
  await user.click(await screen.findByRole('option', { name: 'Notebook' }))
  await user.type(screen.getByLabelText(/Marca/), 'Dell')
  await user.type(screen.getByLabelText(/Modelo/), 'Latitude')
  await user.click(screen.getByRole('combobox', { name: /Revenda/ }))
  await user.click(await screen.findByRole('option', { name: 'REVALLE JUAZEIRO' }))
}

function enviarNf(user: ReturnType<typeof userEvent.setup>, file: File) {
  const input = document.querySelector('input[type="file"]') as HTMLInputElement
  return user.upload(input, file)
}

describe('RegisterItemPage — cadastro', () => {
  it('cadastra com dados de compra e local e depois anexa a NF', async () => {
    mockBase()
    let corpo: Record<string, unknown> = {}
    let postsNf = 0
    server.use(
      http.post('/api/items', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'Item cadastrado com sucesso.', id: 42 })
      }),
      http.post('/api/items/42/nota-fiscal', () => {
        postsNf += 1
        return HttpResponse.json({ detail: 'Nota fiscal anexada.', nf_anexo_key: 'notas_fiscais/x.pdf' })
      })
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    const local = screen.getByLabelText('Local')
    await waitFor(() => expect(within(local).getByRole('option', { name: 'Sala TI' })).toBeInTheDocument())
    await user.selectOptions(local, 'Sala TI')
    await user.type(screen.getByLabelText('Valor de compra (R$)'), '4.599,90')
    await user.type(screen.getByLabelText('Data da compra'), '10012026')
    await user.type(screen.getByLabelText('Fim da garantia'), '10012027')
    await enviarNf(user, arquivoPdf())

    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))

    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(corpo).toMatchObject({
      valor_compra: 4599.9,
      data_compra: '10/01/2026',
      garantia_ate: '10/01/2027',
      local_id: 5,
      revenda: 'Revalle Juazeiro',
    })
    expect(postsNf).toBe(1)
  })

  it('cadastra sem os dados opcionais de compra: o corpo não leva as chaves', async () => {
    mockBase()
    let corpo: Record<string, unknown> = {}
    server.use(
      http.post('/api/items', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'ok', id: 43 })
      })
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))

    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(corpo).not.toHaveProperty('valor_compra')
    expect(corpo).not.toHaveProperty('data_compra')
    expect(corpo).not.toHaveProperty('garantia_ate')
    expect(corpo).not.toHaveProperty('local_id')
  })

  it('recusa NF acima de 10 MB sem enviar', async () => {
    mockBase()
    let postsNf = 0
    server.use(
      http.post('/api/items', () => HttpResponse.json({ detail: 'ok', id: 44 })),
      http.post('/api/items/44/nota-fiscal', () => {
        postsNf += 1
        return HttpResponse.json({ detail: 'x', nf_anexo_key: 'k' })
      })
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    await enviarNf(user, arquivoPdf(10 * 1024 * 1024 + 1))

    expect(await screen.findByText('A nota fiscal deve ter no máximo 10 MB.')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))
    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(postsNf).toBe(0)
  })

  it('valor de compra inválido não envia', async () => {
    mockBase()
    let posts = 0
    server.use(
      http.post('/api/items', () => {
        posts += 1
        return HttpResponse.json({ detail: 'ok', id: 45 })
      })
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    await user.type(screen.getByLabelText('Valor de compra (R$)'), '1,234')
    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))

    expect(await screen.findByText('Valor de compra inválido.')).toBeInTheDocument()
    expect(posts).toBe(0)
  })

  it('data da compra inválida não envia', async () => {
    mockBase()
    let posts = 0
    server.use(
      http.post('/api/items', () => {
        posts += 1
        return HttpResponse.json({ detail: 'ok', id: 46 })
      })
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    await user.type(screen.getByLabelText('Data da compra'), '31022026')
    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))

    expect(await screen.findByText('Data da compra inválida. Use dd/mm/aaaa.')).toBeInTheDocument()
    expect(posts).toBe(0)
  })

  it('falha no anexo leva para a edição do item', async () => {
    mockBase()
    let postsItem = 0
    server.use(
      http.post('/api/items', () => {
        postsItem += 1
        return HttpResponse.json({ detail: 'ok', id: 42 })
      }),
      http.post('/api/items/42/nota-fiscal', () =>
        HttpResponse.json({ detail: 'Arquivo excede o tamanho máximo de 10MB.' }, { status: 413 })
      ),
      http.get('/api/items/42', () => HttpResponse.json({ ...itemExistente, id: 42 }))
    )
    const user = userEvent.setup()
    renderPage('/register')

    await preencherCadastro(user)
    await enviarNf(user, arquivoPdf())
    await user.click(screen.getByRole('button', { name: /Cadastrar Equipamento/ }))

    expect(
      await screen.findByText(
        'Equipamento salvo, mas a nota fiscal não foi anexada: Arquivo excede o tamanho máximo de 10MB.'
      )
    ).toBeInTheDocument()
    expect(await screen.findByRole('heading', { name: 'Editar Equipamento' })).toBeInTheDocument()
    expect(postsItem).toBe(1)
  })

  it('Notebook e Desktop exibem o campo Nº de Série logo após o Host', async () => {
    mockBase()
    server.use(
      http.get('/api/constants', () =>
        HttpResponse.json({
          center_costs: [],
          revendas: ['Revalle Juazeiro', 'Revalle Petrolina'],
          setores: [],
          equipment_types: ['Notebook', 'Desktop'],
          peripheral_types: [],
          removal_reasons: [],
          removal_reasons_attachment: {},
        })
      )
    )
    const user = userEvent.setup()
    renderPage('/register')

    for (const tipo of ['Notebook', 'Desktop']) {
      await user.click(await screen.findByRole('combobox', { name: /Tipo/ }))
      await user.click(await screen.findByRole('option', { name: tipo }))
      const host = screen.getByLabelText('Host')
      const serie = screen.getByLabelText('Nº de Série')
      expect(host.compareDocumentPosition(serie) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy()
    }
  })
})

describe('RegisterItemPage — edição', () => {
  it('hidrata valor e datas sem deslocar o dia', async () => {
    mockBase()
    server.use(http.get('/api/items/9', () => HttpResponse.json(itemExistente)))
    renderPage('/edit/9')

    expect(await screen.findByLabelText('Valor de compra (R$)')).toHaveValue('4.599,90')
    expect(screen.getByLabelText('Data da compra')).toHaveValue('10/01/2026')
    expect(screen.getByLabelText('Fim da garantia')).toHaveValue('10/01/2027')
  })

  it('bloqueia a unidade e transfere pelo diálogo', async () => {
    mockBase()
    let corpo: unknown
    server.use(
      http.get('/api/items/9', () => HttpResponse.json(itemExistente)),
      http.post('/api/items/9/transferir', async ({ request }) => {
        corpo = await request.json()
        return HttpResponse.json({ detail: 'Equipamento transferido de Revalle Juazeiro para Revalle Petrolina.', id: 1 })
      })
    )
    const user = userEvent.setup()
    renderPage('/edit/9')

    expect(await screen.findByRole('combobox', { name: /Revenda/ })).toBeDisabled()
    expect(
      screen.getByText('Para mudar a unidade, use Transferir: a mudança fica registrada na linha do tempo.')
    ).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Transferir' }))
    expect(await screen.findByRole('group', { name: 'Local de destino' })).toBeInTheDocument()
    await user.selectOptions(await screen.findByLabelText('Unidade de destino'), 'REVALLE PETROLINA')
    await user.type(screen.getByLabelText('Motivo da transferência'), 'Mudança de setor')
    await user.click(screen.getByRole('button', { name: 'Confirmar transferência' }))

    await waitFor(() =>
      expect(corpo).toEqual({ para_unidade_id: 2, para_local_id: null, motivo: 'Mudança de setor' })
    )
    expect(
      await screen.findByText('Equipamento transferido de Revalle Juazeiro para Revalle Petrolina.')
    ).toBeInTheDocument()
  })

  it('envia null ao limpar o valor de compra', async () => {
    mockBase()
    let corpo: Record<string, unknown> = {}
    server.use(
      http.get('/api/items/9', () => HttpResponse.json(itemExistente)),
      http.put('/api/items/9', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'Item 9 atualizado.' })
      })
    )
    const user = userEvent.setup()
    renderPage('/edit/9')

    await user.clear(await screen.findByLabelText('Valor de compra (R$)'))
    await user.click(screen.getByRole('button', { name: /Salvar Alterações/ }))

    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(corpo.valor_compra).toBeNull()
    expect(corpo.data_compra).toBe('10/01/2026')
    expect(corpo.garantia_ate).toBe('10/01/2027')
    expect(corpo.local_id).toBe(5)
    // A unidade está travada na edição: nunca vai no corpo.
    expect(corpo).not.toHaveProperty('revenda')
  })

  it('item legado sem unidade pode ser salvo na edição, sem enviar revenda', async () => {
    mockBase()
    let corpo: Record<string, unknown> = {}
    server.use(
      http.get('/api/items/9', () => HttpResponse.json({ ...itemExistente, revenda: null, local_id: null })),
      http.put('/api/items/9', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'Item 9 atualizado.' })
      })
    )
    const user = userEvent.setup()
    renderPage('/edit/9')

    await screen.findByLabelText('Valor de compra (R$)')
    await user.click(screen.getByRole('button', { name: /Salvar Alterações/ }))

    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(corpo).not.toHaveProperty('revenda')
  })

  it('unidade grafada em outra caixa não apaga o local', async () => {
    mockBase()
    let corpo: Record<string, unknown> = {}
    server.use(
      http.get('/api/items/9', () => HttpResponse.json({ ...itemExistente, revenda: 'REVALLE JUAZEIRO' })),
      http.put('/api/items/9', async ({ request }) => {
        corpo = (await request.json()) as Record<string, unknown>
        return HttpResponse.json({ detail: 'Item 9 atualizado.' })
      })
    )
    const user = userEvent.setup()
    renderPage('/edit/9')

    const local = await screen.findByLabelText('Local')
    await waitFor(() => expect(within(local).getByRole('option', { name: 'Sala TI' })).toBeInTheDocument(), {
      timeout: 4000,
    })
    expect(local).toHaveValue('5')
    await user.click(screen.getByRole('button', { name: /Salvar Alterações/ }))

    expect(await screen.findByText('Estoque')).toBeInTheDocument()
    expect(corpo.local_id).toBe(5)
  })

  it('mostra o status atual e o botão de baixar a NF quando há anexo', async () => {
    mockBase()
    server.use(
      http.get('/api/items/9', () =>
        HttpResponse.json({ ...itemExistente, nf_anexo_key: 'notas_fiscais/nf_9_x.pdf' })
      )
    )
    renderPage('/edit/9')

    expect(await screen.findByText('Disponível')).toBeInTheDocument()
    expect(screen.getByText('Nota fiscal anexada')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Baixar nota fiscal' })).toBeInTheDocument()
  })
})
