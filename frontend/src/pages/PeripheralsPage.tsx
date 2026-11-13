import { useEffect, useState, type FormEvent } from 'react'
import { keepPreviousData, useMutation, useQueries, useQuery, useQueryClient } from '@tanstack/react-query'
import { toast } from '@/components/ui/toast'
import {
  Download,
  Link2,
  Trash2,
  Unlink,
  Plus,
  Search,
  ChevronRight,
  Loader2,
} from 'lucide-react'

import {
  createPeripheral,
  deletePeripheral,
  listPeripherals,
  linkPeripheral,
  unlinkPeripheral,
  type Peripheral,
} from '@/api/peripherals'
import { listItemsPaginated, type Item } from '@/api/items'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { exportToCsv, formatDate } from '@/lib/utils'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { PageHeader } from '@/components/layout/PageHeader'
import { Pagination } from '@/components/ui/Pagination'
import { StatusPill } from '@/components/ui/StatusPill'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from '@/components/ui/alert-dialog'

// Tipos de equipamento que aceitam periféricos (mesma regra de LinkPeripheralPage).
const LINK_ALLOWED_TYPES = ['Desktop', 'Notebook', 'Switch', 'Impressora']

// Busca de equipamento na ficha: uma busca pequena por tipo, já filtrada e
// limitada no servidor (com debounce), em vez de carregar centenas de itens e
// filtrar no cliente — mesmo padrão de LinkPeripheralPage.
const ITEM_SEARCH_LIMIT_PER_TYPE = 10

function itemLabel(i: Item) {
  return `#${i.id} — ${i.tipo ?? ''} ${i.brand ?? ''} ${i.model ?? ''}`
}

export default function PeripheralsPage() {
  const queryClient = useQueryClient()
  const { hasRole } = useAuth()
  const { peripheralTypes, isLoading: constantsLoading } = useConstants()
  const canManage = hasRole('Gestor', 'Técnico')

  // Formulário: novo periférico (recolhido por padrão)
  const [tipo, setTipo] = useState('')
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [identificador, setIdentificador] = useState('')
  const [showCreateForm, setShowCreateForm] = useState(false)

  // Filtros de Tipo e Status: `GET /api/peripherals` aceita `status` e `tipo`,
  // então vão para o servidor. O endpoint não aceita `search` nem
  // `limit`/`offset` — busca e paginação são resolvidas no cliente sobre a
  // lista já filtrada pelo servidor.
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [searchQuery, setSearchQuery] = useState('')
  const [tablePage, setTablePage] = useState(1)
  const [pageSize, setPageSize] = usePageSize('perifericos')

  // Detalhe / relação. Guarda um snapshot do periférico escolhido para a
  // ficha continuar aberta mesmo que um filtro do servidor o tire da lista.
  const [selectedPeripheralId, setSelectedPeripheralId] = useState<number | null>(null)
  const [selectedSnapshot, setSelectedSnapshot] = useState<Peripheral | null>(null)
  const [selectedItemId, setSelectedItemId] = useState('')
  const [selectedItemSnapshot, setSelectedItemSnapshot] = useState<Item | null>(null)

  const { data: peripherals = [], isLoading } = useQuery({
    queryKey: ['peripherals', filterTipo, filterStatus],
    queryFn: () =>
      listPeripherals({
        status: filterStatus !== 'all' ? filterStatus : undefined,
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
      }),
  })

  // Busca no cliente: id, tipo, marca, modelo e identificador (não dispara
  // requisição, então não precisa de debounce).
  const filteredPeripherals = peripherals.filter((p) => {
    if (!searchQuery.trim()) return true
    const q = searchQuery.toLowerCase().trim()
    return (
      String(p.id).includes(q) ||
      (p.tipo ?? '').toLowerCase().includes(q) ||
      (p.brand ?? '').toLowerCase().includes(q) ||
      (p.model ?? '').toLowerCase().includes(q) ||
      (p.identificador ?? '').toLowerCase().includes(q)
    )
  })

  const totalTablePages = Math.max(1, Math.ceil(filteredPeripherals.length / pageSize))
  const paginaAtual = Math.min(tablePage, totalTablePages)
  const paginatedPeripherals = filteredPeripherals.slice((paginaAtual - 1) * pageSize, paginaAtual * pageSize)

  // Se um filtro/busca reduzir o total a ponto da página atual deixar de
  // existir, volta para a última página válida.
  useEffect(() => {
    if (tablePage > totalTablePages) setTablePage(totalTablePages)
  }, [tablePage, totalTablePages])

  const selectedPeripheral =
    selectedPeripheralId === null
      ? null
      : (peripherals.find((p) => p.id === selectedPeripheralId) ?? selectedSnapshot)

  function handleSelectPeripheral(p: Peripheral) {
    if (selectedPeripheralId === p.id) {
      setSelectedPeripheralId(null)
      setSelectedSnapshot(null)
    } else {
      setSelectedPeripheralId(p.id)
      setSelectedSnapshot(p)
    }
    setSelectedItemId('')
    setSelectedItemSnapshot(null)
  }

  // Busca do equipamento a vincular (ficha): `itemSearch` é o que o operador
  // digita, `itemSearchAplicado` vai para o servidor com atraso de 400ms.
  const [itemSearch, setItemSearch] = useState('')
  const [itemSearchAplicado, setItemSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => setItemSearchAplicado(itemSearch.trim()), 400)
    return () => clearTimeout(timer)
  }, [itemSearch])

  // Só busca equipamentos quando a seção "Vincular a um ativo" está visível.
  const linkSectionVisible = canManage && !!selectedPeripheral && !selectedPeripheral.link_id
  const linkableQueries = useQueries({
    queries: LINK_ALLOWED_TYPES.map((t) => ({
      queryKey: ['items', 'select', 'linkable', t, itemSearchAplicado],
      queryFn: () =>
        listItemsPaginated({ tipo: t, search: itemSearchAplicado || undefined, limit: ITEM_SEARCH_LIMIT_PER_TYPE }),
      placeholderData: keepPreviousData,
      enabled: linkSectionVisible,
    })),
  })
  const linkableItems = linkableQueries.flatMap((q) => q.data?.items ?? [])

  // Mantém o equipamento escolhido entre as opções mesmo depois que a busca
  // mudar, para o rótulo do seletor não sumir.
  const linkableOptions = [
    ...(selectedItemSnapshot && !linkableItems.some((i) => i.id === selectedItemSnapshot.id)
      ? [selectedItemSnapshot]
      : []),
    ...linkableItems,
  ].map((i) => ({
    value: String(i.id),
    label: itemLabel(i),
    subtitle: upper(i.revenda) || undefined,
  }))

  function handleSelectItem(val: string) {
    setSelectedItemId(val)
    const found = linkableItems.find((i) => String(i.id) === val)
    if (found) setSelectedItemSnapshot(found)
  }

  const createMutation = useMutation({
    mutationFn: createPeripheral,
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setTipo('')
      setBrand('')
      setModel('')
      setIdentificador('')
      setShowCreateForm(false)
      toast.success('Periférico cadastrado com sucesso!')
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, 'Erro ao cadastrar periférico.')),
  })

  const linkMutation = useMutation({
    mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals'] })
      setSelectedItemId('')
      setSelectedItemSnapshot(null)
      toast.success('Periférico vinculado com sucesso!')
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, 'Erro ao vincular.')),
  })

  const unlinkMutation = useMutation({
    mutationFn: (id: number) => unlinkPeripheral(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals'] })
      toast.success('Periférico desvinculado.')
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, 'Erro ao desvincular.')),
  })

  const deleteMutation = useMutation({
    mutationFn: (id: number) => deletePeripheral(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setSelectedPeripheralId(null)
      setSelectedSnapshot(null)
      toast.success('Periférico inativado.')
    },
    onError: (e: unknown) => toast.error(getErrorMessage(e, 'Erro ao inativar.')),
  })

  function handleSubmit(e: FormEvent) {
    e.preventDefault()
    createMutation.mutate({ tipo, brand, model, identificador })
  }

  return (
    <div className="space-y-6 max-w-7xl mx-auto">
      {/* Cabeçalho */}
      <PageHeader
        eyebrow="Inventário Relacional"
        eyebrowDetail="Master / Detail"
        title="Periféricos e Vínculos"
        description="Gestão de periféricos — mouse, teclado, monitor, headset — com rastreabilidade relacional ao ativo principal."
        actions={
          <>
          <Button
            variant="outline"
            size="sm"
            disabled={!filteredPeripherals.length}
            onClick={() =>
              exportToCsv(
                'perifericos',
                filteredPeripherals as unknown as Record<string, unknown>[],
                [
                  { key: 'id', label: 'ID' },
                  { key: 'tipo', label: 'Tipo' },
                  { key: 'brand', label: 'Marca' },
                  { key: 'model', label: 'Modelo' },
                  { key: 'identificador', label: 'S/N' },
                  { key: 'status', label: 'Status' },
                  { key: 'date_registered', label: 'Cadastro' },
                ]
              )
            }
          >
            <Download />
            Exportar CSV
          </Button>
          {canManage && (
            <Button size="sm" onClick={() => setShowCreateForm((p) => !p)}>
              <Plus />
              Novo Periférico
            </Button>
          )}
          </>
        }
      />

      {/* Painel de Cadastro Rápido (colapsável) */}
      {showCreateForm && canManage && (
        <div className="rounded-md border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <h2 className="text-body font-semibold text-brand-text">Cadastrar Novo Periférico</h2>
            <p className="text-caption text-muted-foreground">
              Registre um periférico individual para posterior vinculação a um ativo.
            </p>
          </div>
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Tipo *</Label>
                <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                  <SelectTrigger className="h-9">
                    <SelectValue placeholder="Selecione o tipo" />
                  </SelectTrigger>
                  <SelectContent>
                    {peripheralTypes.map((t) => (
                      <SelectItem key={t} value={t}>
                        {t}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">S/N (Identificador) *</Label>
                <Input
                  value={identificador}
                  onChange={(e) => setIdentificador(e.target.value)}
                  required
                  placeholder="Número de série único"
                  className="font-mono h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Marca</Label>
                <Input
                  value={brand}
                  onChange={(e) => setBrand(e.target.value)}
                  placeholder="Ex: Logitech"
                  className="h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Modelo</Label>
                <Input
                  value={model}
                  onChange={(e) => setModel(e.target.value)}
                  placeholder="Ex: MX Keys"
                  className="h-9"
                />
              </div>
            </div>
            <div className="flex justify-end gap-2">
              <Button type="button" variant="outline" size="sm" onClick={() => setShowCreateForm(false)}>
                Cancelar
              </Button>
              <Button type="submit" size="sm" loading={createMutation.isPending}>
                Cadastrar Periférico
              </Button>
            </div>
          </form>
        </div>
      )}

      {/* Grade Mestre / Detalhe */}
      <div className="grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_360px] gap-6 items-start">
        {/* MESTRE — Lista de Periféricos */}
        <div className="min-w-0 rounded-md border border-border bg-surface p-5 space-y-4">
          {/* Barra de filtros */}
          <div className="flex flex-col sm:flex-row gap-3 border-b border-border pb-4">
            <div className="relative flex-1">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
              <Input
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value)
                  setTablePage(1)
                }}
                aria-label="Buscar periféricos"
                placeholder="Buscar por ID, tipo, marca, S/N…"
                className="pl-8 h-8 text-body-sm"
              />
            </div>
            <Select
              value={filterTipo}
              onValueChange={(val) => {
                setFilterTipo(val)
                setTablePage(1)
              }}
              disabled={constantsLoading}
            >
              <SelectTrigger aria-label="Tipo" className="w-full sm:w-36 h-8 text-body-sm">
                <SelectValue placeholder="Tipo" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos os tipos</SelectItem>
                {peripheralTypes.map((t) => (
                  <SelectItem key={t} value={t}>
                    {t}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Select
              value={filterStatus}
              onValueChange={(val) => {
                setFilterStatus(val)
                setTablePage(1)
              }}
            >
              <SelectTrigger aria-label="Status" className="w-full sm:w-36 h-8 text-body-sm">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">Todos</SelectItem>
                <SelectItem value="Disponível">Disponível</SelectItem>
                <SelectItem value="Em Uso">Em Uso</SelectItem>
                <SelectItem value="Substituido">Substituído</SelectItem>
              </SelectContent>
            </Select>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full min-w-[640px] text-left text-body-sm">
              <thead className="bg-brand-soft text-brand-soft-foreground text-caption uppercase tracking-wider">
                <tr>
                  <th className="py-2 px-3 font-semibold">ID</th>
                  <th className="py-2 px-3 font-semibold">Tipo</th>
                  <th className="py-2 px-3 font-semibold">Marca / Modelo</th>
                  <th className="py-2 px-3 font-semibold">S/N</th>
                  <th className="py-2 px-3 font-semibold">Status</th>
                  <th className="py-2 px-3 font-semibold">Cadastro</th>
                  <th className="py-2 px-3 text-right font-semibold">Detalhe</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {isLoading ? (
                  <tr>
                    <td colSpan={7} className="py-8 text-center text-caption text-muted-foreground">
                      <span className="inline-flex items-center gap-2">
                        <Loader2 className="size-3.5 animate-spin" />
                        Carregando…
                      </span>
                    </td>
                  </tr>
                ) : filteredPeripherals.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="py-8 text-center text-caption text-muted-foreground">
                      Nenhum periférico encontrado.
                    </td>
                  </tr>
                ) : (
                  paginatedPeripherals.map((p) => {
                    const isSelected = selectedPeripheralId === p.id
                    return (
                      <tr
                        key={p.id}
                        onClick={() => handleSelectPeripheral(p)}
                        className={`cursor-pointer transition-colors ${isSelected ? 'bg-brand-soft text-brand-soft-foreground' : 'hover:bg-surface-alt'}`}
                      >
                        <td className="py-2 px-3 font-mono font-semibold">#{p.id}</td>
                        <td className="py-2 px-3 font-medium">{p.tipo}</td>
                        <td className="py-2 px-3 text-muted-foreground">
                          {p.brand} {p.model}
                        </td>
                        <td className="py-2 px-3 font-mono text-foreground">
                          {p.identificador || '—'}
                        </td>
                        <td className="py-2 px-3">
                          <StatusPill status={p.status} />
                        </td>
                        <td className="py-2 px-3 font-mono text-muted-foreground">
                          {formatDate(p.date_registered)}
                        </td>
                        <td className="py-2 px-3 text-right">
                          <ChevronRight
                            className={`size-4 ml-auto transition-transform ${isSelected ? 'rotate-90 text-foreground' : 'text-muted-foreground'}`}
                          />
                        </td>
                      </tr>
                    )
                  })
                )}
              </tbody>
            </table>
          </div>

          <div className="border-t border-border pt-3">
            <Pagination
              page={tablePage}
              pageSize={pageSize}
              total={filteredPeripherals.length}
              onPageChange={setTablePage}
              onPageSizeChange={(n) => {
                setPageSize(n)
                setTablePage(1)
              }}
            />
          </div>
        </div>

        {/* DETALHE — Painel de Relações */}
        <div className="rounded-md border border-border bg-surface p-5 space-y-4 min-h-[400px]">
          {selectedPeripheral ? (
            <>
              <div className="border-b border-border pb-3">
                <span className="text-caption font-semibold uppercase tracking-wider text-brand-text block">
                  Ficha do Periférico
                </span>
                <p className="text-body-lg font-semibold text-foreground mt-0.5">
                  #{selectedPeripheral.id} — {selectedPeripheral.tipo}
                </p>
                <p className="text-caption text-muted-foreground">
                  {selectedPeripheral.brand} {selectedPeripheral.model} • S/N:{' '}
                  <span className="font-mono text-foreground">
                    {selectedPeripheral.identificador || '—'}
                  </span>
                </p>
              </div>

              <div className="space-y-2 text-caption">
                <div className="flex justify-between py-1 border-b border-border">
                  <span className="text-muted-foreground">Status</span>
                  <StatusPill status={selectedPeripheral.status} />
                </div>
                <div className="flex justify-between py-1 border-b border-border">
                  <span className="text-muted-foreground">Vinculado</span>
                  <span className="font-semibold text-foreground">
                    {selectedPeripheral.link_id
                      ? `✓ Sim (link #${selectedPeripheral.link_id})`
                      : '○ Livre'}
                  </span>
                </div>
                <div className="flex justify-between py-1 border-b border-border">
                  <span className="text-muted-foreground">Cadastro</span>
                  <span className="font-mono text-foreground">
                    {formatDate(selectedPeripheral.date_registered)}
                  </span>
                </div>
              </div>

              {/* Ações de Gestão */}
              {canManage && (
                <div className="space-y-3 pt-3 border-t border-border">
                  <span className="text-caption font-semibold uppercase tracking-wider text-brand-text block">
                    Ações de Gestão
                  </span>

                  {selectedPeripheral.link_id ? (
                    <Button
                      size="sm"
                      variant="outline"
                      loading={unlinkMutation.isPending}
                      onClick={() => unlinkMutation.mutate(selectedPeripheral.link_id!)}
                      className="w-full"
                    >
                      <Unlink />
                      Desvincular do Ativo
                    </Button>
                  ) : (
                    <div className="space-y-2">
                      <Label className="text-caption font-medium">Vincular a um ativo:</Label>
                      {/* Busca no servidor (debounce de 400ms) controlando o
                          campo de busca do próprio dropdown. */}
                      <SearchableSelect
                        options={linkableOptions}
                        value={selectedItemId}
                        onValueChange={handleSelectItem}
                        placeholder="Selecione o equipamento…"
                        searchPlaceholder="Buscar ativo…"
                        search={itemSearch}
                        onSearchChange={setItemSearch}
                      />
                      {selectedItemId && (
                        <Button
                          size="sm"
                          loading={linkMutation.isPending}
                          onClick={() => linkMutation.mutate(selectedPeripheral.id)}
                          className="w-full"
                        >
                          <Link2 />
                          Confirmar Vínculo
                        </Button>
                      )}
                    </div>
                  )}

                  <AlertDialog>
                    <AlertDialogTrigger asChild>
                      <Button
                        size="sm"
                        variant="outline"
                        loading={deleteMutation.isPending}
                        disabled={selectedPeripheral.status === 'Em Uso'}
                        className="w-full text-destructive hover:bg-surface-alt"
                      >
                        <Trash2 />
                        Inativar Periférico
                      </Button>
                    </AlertDialogTrigger>
                    <AlertDialogContent>
                      <AlertDialogHeader>
                        <AlertDialogTitle>Inativar periférico #{selectedPeripheral.id}?</AlertDialogTitle>
                        <AlertDialogDescription>
                          {selectedPeripheral.tipo} — {selectedPeripheral.brand || '-'} {selectedPeripheral.model || ''} (S/N: {selectedPeripheral.identificador || '-'}).
                          {' '}Periféricos em uso não podem ser inativados. Esta ação não pode ser desfeita.
                        </AlertDialogDescription>
                      </AlertDialogHeader>
                      <AlertDialogFooter>
                        <AlertDialogCancel>Cancelar</AlertDialogCancel>
                        <AlertDialogAction
                          onClick={() => deleteMutation.mutate(selectedPeripheral.id)}
                        >
                          Confirmar Inativação
                        </AlertDialogAction>
                      </AlertDialogFooter>
                    </AlertDialogContent>
                  </AlertDialog>
                </div>
              )}
            </>
          ) : (
            <div className="py-16 text-center text-caption text-muted-foreground">
              Selecione um periférico na tabela para visualizar a ficha técnica e opções de vinculação.
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
