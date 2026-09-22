import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import { useNavigate } from 'react-router-dom'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listPeripherals, createPeripheral, deletePeripheral, type Peripheral } from '@/api/peripherals'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Badge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
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
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import type { ColumnDef } from '@tanstack/react-table'
import { formatDate } from '@/lib/utils'
import { Trash2, Loader2, Plus, X, Link2, ArrowRight } from 'lucide-react'

const PAGE_SIZE = 10
const STATUS_OPTIONS = ['Disponível', 'Em Uso', 'Substituido']

function PeripheralStatusBadge({ status }: { status?: string }) {
  if (status === 'Disponível') return <Badge variant="success" showDot>{status}</Badge>
  if (status === 'Em Uso') return <Badge variant="warning" showDot>{status}</Badge>
  if (status === 'Substituido') return <Badge variant="danger" showDot>{status}</Badge>
  return <Badge>{status ?? '-'}</Badge>
}

/**
 * Painel de detalhe (ficha do periférico), aberto ao clicar numa linha da
 * tabela. Não duplica o que `/link` (LinkPeripheralPage) já faz — vincular,
 * desvincular e substituir continuam morando só lá, com a busca de
 * equipamento, a lista de periféricos vinculados etc. Aqui só levamos o
 * operador até lá, com um botão que navega para `/link`.
 */
function PeripheralDetailModal({
  peripheral,
  onClose,
  canManage,
  onDeactivate,
  isDeactivating,
}: {
  peripheral: Peripheral | null
  onClose: () => void
  canManage: boolean
  onDeactivate: (id: number) => void
  isDeactivating: boolean
}) {
  const navigate = useNavigate()
  if (!peripheral) return null

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-black/50 animate-fade-in select-none">
      <div className="relative w-full max-w-md surface-panel shadow-overlay overflow-hidden">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border">
          <div>
            <span className="text-caption text-muted-foreground">Ficha do Periférico</span>
            <h2 className="text-heading-sm text-foreground mt-0.5">
              #{peripheral.id} — {peripheral.tipo}
            </h2>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose} className="text-muted-foreground hover:text-foreground">
            <X size={18} />
          </Button>
        </div>

        <div className="p-6 space-y-4">
          <div className="grid grid-cols-2 gap-x-4 gap-y-3">
            <div>
              <p className="text-caption text-muted-foreground">Marca</p>
              <p className="text-body-sm text-foreground">{peripheral.brand || '-'}</p>
            </div>
            <div>
              <p className="text-caption text-muted-foreground">Modelo</p>
              <p className="text-body-sm text-foreground">{peripheral.model || '-'}</p>
            </div>
            <div>
              <p className="text-caption text-muted-foreground">Identificador (S/N)</p>
              <p className="text-body-sm num text-foreground">{peripheral.identificador || '-'}</p>
            </div>
            <div>
              <p className="text-caption text-muted-foreground">Cadastro</p>
              <p className="text-body-sm num text-foreground">{formatDate(peripheral.date_registered)}</p>
            </div>
            <div className="col-span-2">
              <p className="text-caption text-muted-foreground">Status</p>
              <div className="mt-0.5"><PeripheralStatusBadge status={peripheral.status} /></div>
            </div>
          </div>

          {canManage && (
            <div className="space-y-2 pt-4 border-t border-border">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                Ações
              </span>

              {/* Vincular/desvincular/substituir são todos feitos em /link — este
                  botão só leva até lá, sem duplicar a tela. */}
              <Button
                variant="outline"
                size="sm"
                className="w-full justify-between"
                onClick={() => { onClose(); navigate('/link') }}
              >
                <span className="flex items-center gap-1.5">
                  <Link2 size={14} />
                  {peripheral.status === 'Em Uso' ? 'Gerenciar vínculo em Vincular Periféricos' : 'Vincular a um equipamento'}
                </span>
                <ArrowRight size={14} />
              </Button>

              <AlertDialog>
                <AlertDialogTrigger asChild>
                  <Button
                    size="sm"
                    variant="outline"
                    className="w-full text-muted-foreground hover:text-destructive"
                    disabled={isDeactivating}
                  >
                    <Trash2 size={14} />
                    Inativar Periférico
                  </Button>
                </AlertDialogTrigger>
                <AlertDialogContent>
                  <AlertDialogHeader>
                    <AlertDialogTitle>Inativar periférico #{peripheral.id}?</AlertDialogTitle>
                    <AlertDialogDescription>
                      {peripheral.tipo} — {peripheral.brand || '-'} {peripheral.model || ''} (S/N: {peripheral.identificador || '-'}).
                      {' '}Periféricos em uso não podem ser inativados. Esta ação não pode ser desfeita.
                    </AlertDialogDescription>
                  </AlertDialogHeader>
                  <AlertDialogFooter>
                    <AlertDialogCancel>Cancelar</AlertDialogCancel>
                    <AlertDialogAction onClick={() => onDeactivate(peripheral.id)}>
                      Inativar
                    </AlertDialogAction>
                  </AlertDialogFooter>
                </AlertDialogContent>
              </AlertDialog>
            </div>
          )}
        </div>
      </div>
    </div>,
    document.body
  )
}

export default function PeripheralsPage() {
  const queryClient = useQueryClient()
  const { hasRole } = useAuth()
  const canManage = hasRole('Gestor', 'Técnico')
  const { peripheralTypes, isLoading: constantsLoading } = useConstants()

  // Formulário de cadastro: recolhido por padrão, aberto pelo botão "Novo
  // Periférico" (T4/Step 3) — antes ocupava o topo da página o tempo todo.
  const [showCreateForm, setShowCreateForm] = useState(false)
  const [tipo, setTipo] = useState('')
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [identificador, setIdentificador] = useState('')

  // Filtros de Tipo e Status: `GET /api/peripherals` aceita `status` e `tipo`
  // (backend/app/routers/peripherals.py:19-31), então vão para o servidor. O
  // endpoint não aceita `search` nem `limit`/`offset` — busca e paginação
  // abaixo são resolvidas no cliente sobre a lista já filtrada pelo servidor.
  const [filterTipo, setFilterTipo] = useState('all')
  const [filterStatus, setFilterStatus] = useState('all')
  const [search, setSearch] = useState('')
  const [pageIndex, setPageIndex] = useState(0)

  const [selectedPeripheral, setSelectedPeripheral] = useState<Peripheral | null>(null)

  const hasFilters = filterTipo !== 'all' || filterStatus !== 'all' || search !== ''

  const { data: peripherals = [], isLoading } = useQuery({
    queryKey: ['peripherals', filterTipo, filterStatus],
    queryFn: () =>
      listPeripherals({
        status: filterStatus !== 'all' ? filterStatus : undefined,
        tipo: filterTipo !== 'all' ? filterTipo : undefined,
      }),
  })

  // Busca no cliente: id, tipo, marca, modelo e identificador. Sem debounce —
  // ao contrário de HistoryPage/StockPage, não dispara requisição nenhuma
  // (o endpoint não aceita `search`), então não há custo de rede a evitar.
  const filteredPeripherals = peripherals.filter((p) => {
    if (!search.trim()) return true
    const q = search.toLowerCase().trim()
    return (
      String(p.id).includes(q) ||
      (p.tipo ?? '').toLowerCase().includes(q) ||
      (p.brand ?? '').toLowerCase().includes(q) ||
      (p.model ?? '').toLowerCase().includes(q) ||
      (p.identificador ?? '').toLowerCase().includes(q)
    )
  })
  const total = filteredPeripherals.length
  const pagedPeripherals = filteredPeripherals.slice(pageIndex * PAGE_SIZE, (pageIndex + 1) * PAGE_SIZE)

  // Se um filtro/busca reduzir o total a ponto da página atual deixar de
  // existir, volta para a última página válida (mesmo padrão do HistoryPage).
  useEffect(() => {
    const maxPageIndex = Math.max(0, Math.ceil(total / PAGE_SIZE) - 1)
    if (pageIndex > maxPageIndex) setPageIndex(maxPageIndex)
  }, [total, pageIndex])

  function handleTipoChange(value: string) {
    setFilterTipo(value)
    setPageIndex(0)
  }

  function handleStatusChange(value: string) {
    setFilterStatus(value)
    setPageIndex(0)
  }

  function handleSearchChange(value: string) {
    setSearch(value)
    setPageIndex(0)
  }

  function handleClearFilters() {
    setFilterTipo('all')
    setFilterStatus('all')
    setSearch('')
    setPageIndex(0)
  }

  const mutation = useMutation({
    mutationFn: createPeripheral,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setTipo(''); setBrand(''); setModel(''); setIdentificador('')
      setShowCreateForm(false)
      toast('Periférico cadastrado com sucesso!')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao cadastrar.'), 'error')
    },
  })

  // T4: DELETE /api/peripherals/{id} (Gestor e Técnico) inativa um periférico
  // cadastrado por engano; o backend recusa com 400 se ele estiver em uso —
  // a mensagem exibida no toast de erro vem direto do `detail` da resposta.
  const deleteMutation = useMutation({
    mutationFn: (peripheralId: number) => deletePeripheral(peripheralId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setSelectedPeripheral(null)
      toast('Periférico inativado com sucesso.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao inativar periférico.'), 'error')
    },
  })

  const columns: ColumnDef<Peripheral, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'model', header: 'Modelo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'identificador', header: 'Identificador (S/N)', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'status', header: 'Status', cell: ({ row }) => <PeripheralStatusBadge status={row.original.status} /> },
    { accessorKey: 'date_registered', header: 'Cadastro', cell: ({ getValue }) => formatDate(getValue() as string) },
    // Ação de inativar só aparece para quem o backend de fato autoriza
    // (Gestor e Técnico), mesmo padrão de gating usado em StockPage.
    ...(canManage ? [{
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: Peripheral } }) => (
        <AlertDialog>
          <AlertDialogTrigger asChild>
            <Button size="sm" variant="ghost" className="text-muted-foreground hover:text-destructive" disabled={deleteMutation.isPending}>
              <Trash2 size={14} />
            </Button>
          </AlertDialogTrigger>
          <AlertDialogContent>
            <AlertDialogHeader>
              <AlertDialogTitle>Inativar periférico #{row.original.id}?</AlertDialogTitle>
              <AlertDialogDescription>
                {row.original.tipo} — {row.original.brand || '-'} {row.original.model || ''} (S/N: {row.original.identificador || '-'}).
                {' '}Periféricos em uso não podem ser inativados. Esta ação não pode ser desfeita.
              </AlertDialogDescription>
            </AlertDialogHeader>
            <AlertDialogFooter>
              <AlertDialogCancel>Cancelar</AlertDialogCancel>
              <AlertDialogAction onClick={() => deleteMutation.mutate(row.original.id)}>
                Inativar
              </AlertDialogAction>
            </AlertDialogFooter>
          </AlertDialogContent>
        </AlertDialog>
      ),
    } as ColumnDef<Peripheral, unknown>] : []),
  ]

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Ativos de Suporte"
        eyebrowDetail="Periféricos"
        title="Periféricos"
        description="Cadastre e gerencie periféricos como mouse, teclado, monitor etc."
        actions={
          canManage && (
            <Button variant="gradient" size="sm" onClick={() => setShowCreateForm((v) => !v)}>
              <Plus size={16} />
              Novo Periférico
            </Button>
          )
        }
      />

      {/* Formulário de cadastro — recolhido por padrão (Step 3). */}
      {showCreateForm && canManage && (
        <form
          onSubmit={(e) => { e.preventDefault(); mutation.mutate({ tipo, brand, model, identificador }) }}
          className="surface-panel p-6 space-y-4"
        >
          <PanelHeader
            title="Cadastrar Periférico"
            description="Tipo, identificador e dados opcionais do periférico."
            actions={
              <Button type="button" variant="ghost" size="sm" onClick={() => setShowCreateForm(false)}>
                <X size={14} />
                Fechar
              </Button>
            }
          />
          <div className="grid grid-cols-2 gap-4">
            <div className="flex flex-col gap-1.5">
              <Label>Tipo *</Label>
              <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
                <SelectTrigger><SelectValue placeholder="Selecione o tipo" /></SelectTrigger>
                <SelectContent>
                  {peripheralTypes.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Identificador (S/N) *</Label>
              <Input value={identificador} onChange={e => setIdentificador(e.target.value)} required placeholder="Número de série único" />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Marca</Label>
              <Input value={brand} onChange={e => setBrand(e.target.value)} />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Modelo</Label>
              <Input value={model} onChange={e => setModel(e.target.value)} />
            </div>
          </div>
          <div className="flex justify-end">
            <Button type="submit" disabled={mutation.isPending}>
              {mutation.isPending ? 'Cadastrando...' : 'Cadastrar Periférico'}
            </Button>
          </div>
        </form>
      )}

      {/* Filtros de Tipo e Status — resolvidos no servidor (Step 1). */}
      <div className="surface-panel p-3 space-y-3">
        <PanelHeader title="Filtros de Consulta" />
        <div className="flex flex-wrap items-center gap-3">
          <Select value={filterTipo} onValueChange={handleTipoChange} disabled={constantsLoading}>
            <SelectTrigger className="w-48">
              <SelectValue placeholder="Tipo de Periférico" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todos os tipos</SelectItem>
              {peripheralTypes.map((t) => (
                <SelectItem key={t} value={t}>{t}</SelectItem>
              ))}
            </SelectContent>
          </Select>

          <Select value={filterStatus} onValueChange={handleStatusChange}>
            <SelectTrigger className="w-48">
              <SelectValue placeholder="Status" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">Todos os status</SelectItem>
              {STATUS_OPTIONS.map((s) => (
                <SelectItem key={s} value={s}>{s}</SelectItem>
              ))}
            </SelectContent>
          </Select>

          {hasFilters && (
            <Button
              variant="ghost"
              size="sm"
              onClick={handleClearFilters}
              className="text-xs text-muted-foreground hover:text-foreground"
            >
              <X size={14} />
              Limpar Filtros
            </Button>
          )}
        </div>
      </div>

      {/* Lista */}
      {isLoading ? (
        <div className="py-8 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
          <Loader2 className="h-4 w-4 animate-spin" />
          Carregando...
        </div>
      ) : (
        <DataTable
          data={pagedPeripherals}
          columns={columns}
          onRowClick={(p) => setSelectedPeripheral(p)}
          searchPlaceholder="Buscar por ID, tipo, marca, modelo, identificador..."
          pagination={{
            total,
            pageIndex,
            pageSize: PAGE_SIZE,
            onPageChange: setPageIndex,
            search,
            onSearchChange: handleSearchChange,
          }}
        />
      )}

      <PeripheralDetailModal
        peripheral={selectedPeripheral}
        onClose={() => setSelectedPeripheral(null)}
        canManage={canManage}
        onDeactivate={(id) => deleteMutation.mutate(id)}
        isDeactivating={deleteMutation.isPending}
      />
    </div>
  )
}
