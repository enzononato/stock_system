import { useEffect, useState } from 'react'
import { useQueries, useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated, type Item } from '@/api/items'
import {
  listPeripherals,
  listItemPeripherals,
  linkPeripheral,
  unlinkPeripheral,
  replacePeripheral,
  type Peripheral,
} from '@/api/peripherals'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { FileUpload } from '@/components/ui/FileUpload'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { Link2, Unlink, RefreshCw } from 'lucide-react'

// Tipos de equipamento que aceitam periféricos. Regra própria desta tela
// (subconjunto curado, não um cadastro de domínio) — não é uma das listas
// duplicadas com o backend que a T2 pede para eliminar (revendas, setores,
// centros de custo, tipos de equipamento/periférico, motivos de remoção),
// por isso continua fixa aqui.
const LINK_ALLOWED_TYPES = ['Desktop', 'Notebook', 'Switch', 'Impressora']

// Esta tela só tem um bloco de dados, e é uma lista de SELEÇÃO — o operador
// procura o equipamento a que vai vincular o periférico, não folheia página
// por página (T3). Por isso liga a busca (com debounce) ao parâmetro
// `search` do servidor, com `limit` pequeno, em vez de buscar os 500
// primeiros e filtrar no cliente (o antigo limite fixo de 500, removido
// nesta task). `GET /api/items` só aceita um `tipo` por chamada, e esta tela
// precisa de quatro (LINK_ALLOWED_TYPES) — em vez de puxar um lote genérico e
// filtrar no cliente (o que deixaria tipos menos comuns sub-representados
// nos 20 primeiros resultados), dispara uma busca pequena por tipo em
// paralelo (`useQueries`), cada uma já filtrada e limitada no servidor.
const ITEM_SEARCH_LIMIT_PER_TYPE = 10

function PeripheralCard({
  peripheral,
  action,
  actionLabel,
  actionIcon,
  variant = 'default',
}: {
  peripheral: Peripheral
  action: () => void
  actionLabel: string
  actionIcon: React.ReactNode
  variant?: 'default' | 'destructive'
}) {
  return (
    <div className="flex items-center justify-between border border-border rounded p-3 surface-interactive">
      <div className="flex flex-col gap-0.5">
        <span className="text-body-sm font-medium text-foreground">
          {peripheral.tipo} — {peripheral.brand || '-'} {peripheral.model || ''}
        </span>
        <span className="text-[11px] text-muted-foreground">S/N: <span className="num">{peripheral.identificador || '-'}</span></span>
      </div>
      <div className="flex items-center gap-3">
        <Badge variant={peripheral.status === 'Disponível' ? 'success' : peripheral.status === 'Em Uso' ? 'warning' : 'danger'}>
          {peripheral.status}
        </Badge>
        <Button size="sm" variant={variant === 'destructive' ? 'destructive' : 'outline'} onClick={action}>
          {actionIcon}
          <span>{actionLabel}</span>
        </Button>
      </div>
    </div>
  )
}

export default function LinkPeripheralPage() {
  const queryClient = useQueryClient()
  const [selectedItemId, setSelectedItemId] = useState('')

  // Substituição
  const [replacingLinkId, setReplacingLinkId] = useState<number | null>(null)
  const [replacingOldId, setReplacingOldId] = useState<number | null>(null)
  const [replaceNewId, setReplaceNewId] = useState('')
  const [replaceReason, setReplaceReason] = useState('')
  const [replaceAttachment, setReplaceAttachment] = useState<File | null>(null)

  // Busca do equipamento a vincular: `itemSearch` é o que o operador digita,
  // `itemSearchAplicado` é o que vai para o servidor, com atraso, para não
  // disparar uma requisição por tecla (mesmo padrão do debounce de busca já
  // usado em HistoryPage/StockPage).
  const [itemSearch, setItemSearch] = useState('')
  const [itemSearchAplicado, setItemSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => setItemSearchAplicado(itemSearch.trim()), 400)
    return () => clearTimeout(timer)
  }, [itemSearch])

  const linkableQueries = useQueries({
    queries: LINK_ALLOWED_TYPES.map((tipo) => ({
      queryKey: ['items', 'select', 'linkable', tipo, itemSearchAplicado],
      queryFn: () =>
        listItemsPaginated({ tipo, search: itemSearchAplicado || undefined, limit: ITEM_SEARCH_LIMIT_PER_TYPE }),
      placeholderData: keepPreviousData,
    })),
  })
  const linkableItems = linkableQueries.flatMap((q) => q.data?.items ?? [])

  // Snapshot do equipamento selecionado: `linkableItems` é só o lote pequeno
  // (≤20) da busca atual, então some da lista assim que o operador digita
  // outra coisa. Guarda o item completo no momento da seleção para a linha
  // "Status / Revenda" abaixo continuar mostrando os dados do equipamento
  // escolhido, mesmo depois que a busca mudar.
  const [selectedItemSnapshot, setSelectedItemSnapshot] = useState<Item | null>(null)
  function handleSelectItem(val: string) {
    setSelectedItemId(val)
    const found = linkableItems.find(i => String(i.id) === val)
    if (found) setSelectedItemSnapshot(found)
  }

  const { data: linkedPeripherals = [], refetch: refetchLinked } = useQuery({
    queryKey: ['item-peripherals', selectedItemId],
    queryFn: () => listItemPeripherals(Number(selectedItemId)),
    enabled: !!selectedItemId,
  })

  const { data: availablePeripherals = [] } = useQuery({
    queryKey: ['peripherals', 'Disponível'],
    queryFn: () => listPeripherals({ status: 'Disponível' }),
  })

  const linkMutation = useMutation({
    mutationFn: (pid: number) => linkPeripheral(Number(selectedItemId), pid),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast('Periférico vinculado!')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao vincular.'), 'error')
    },
  })

  const unlinkMutation = useMutation({
    mutationFn: (linkId: number) => unlinkPeripheral(linkId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast('Periférico desvinculado.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao desvincular.'), 'error')
    },
  })

  const replaceMutation = useMutation({
    mutationFn: () =>
      replacePeripheral(
        Number(selectedItemId),
        replacingOldId!,
        Number(replaceNewId),
        replaceReason,
        replaceAttachment ?? undefined
      ),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setReplacingLinkId(null)
      setReplacingOldId(null)
      setReplaceNewId('')
      setReplaceReason('')
      setReplaceAttachment(null)
      toast('Periférico substituído com sucesso!')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao substituir.'), 'error')
    },
  })

  const selectedItem = linkableItems.find(i => String(i.id) === selectedItemId) ?? selectedItemSnapshot

  return (
    <div className="page-container-reading space-y-6">
      <PageHeader
        eyebrow="Ativos de Suporte"
        eyebrowDetail="Vínculo de Periféricos"
        title="Vincular Periféricos"
        description="Associe periféricos a equipamentos como desktops, notebooks, switches e impressoras."
      />

      {/* Seletor de equipamento */}
      <div className="surface-panel p-4 flex flex-col gap-2">
        <Label>Selecione o Equipamento</Label>
        <div className="max-w-md">
          {/* Busca no servidor (debounce de 400ms) controlando o próprio campo
              de busca do dropdown via `search`/`onSearchChange` — sem isso, o
              usuário veria uma segunda caixa de busca (a interna do
              SearchableSelect) além desta, filtrando só os 20 já carregados. */}
          <SearchableSelect
            options={linkableItems.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={handleSelectItem}
            placeholder="Selecione um equipamento..."
            searchPlaceholder="Buscar por marca, modelo ou identificador (patrimônio)..."
            search={itemSearch}
            onSearchChange={setItemSearch}
          />
        </div>
        {selectedItem && (
          <p className="text-body-sm text-muted-foreground">
            Status: <strong>{selectedItem.status}</strong> · Revenda: <strong>{selectedItem.revenda}</strong>
          </p>
        )}
      </div>

      {selectedItemId && (
        <div className="grid grid-cols-2 gap-6">
          {/* Periféricos vinculados */}
          <div className="surface-panel p-4 space-y-3">
            <PanelHeader
              title={<>Periféricos Vinculados (<span className="num">{linkedPeripherals.length}</span>)</>}
              actions={
                <Button size="sm" variant="ghost" onClick={() => refetchLinked()}>
                  <RefreshCw size={13} />
                </Button>
              }
            />
            {linkedPeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-10 text-center">Nenhum periférico vinculado.</p>
            ) : (
              <div className="space-y-2">
                {linkedPeripherals.map(p => (
                  <div key={p.link_id} className="flex flex-col gap-1">
                    <PeripheralCard
                      peripheral={p}
                      action={() => unlinkMutation.mutate(p.link_id!)}
                      actionLabel="Desvincular"
                      actionIcon={<Unlink size={13} />}
                      variant="destructive"
                    />
                    <button
                      className="text-xs text-muted-foreground hover:text-foreground hover:underline text-left ml-1"
                      onClick={() => { setReplacingLinkId(p.link_id!); setReplacingOldId(p.id) }}
                    >
                      Substituir...
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Periféricos disponíveis */}
          <div className="surface-panel p-4 space-y-3">
            <PanelHeader title={<>Periféricos Disponíveis (<span className="num">{availablePeripherals.length}</span>)</>} />
            {availablePeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-10 text-center">Nenhum periférico disponível.</p>
            ) : (
              <div className="space-y-2 max-h-[420px] overflow-y-auto pr-1">
                {availablePeripherals.map(p => (
                  <PeripheralCard
                    key={p.id}
                    peripheral={p}
                    action={() => linkMutation.mutate(p.id)}
                    actionLabel="Vincular"
                    actionIcon={<Link2 size={13} />}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Painel de substituição */}
      {replacingLinkId && replacingOldId && (
        <div className="figure-ground-panel space-y-4">
          <PanelHeader title="Substituição de Periférico" />
          <h3 className="text-heading-sm text-foreground">Substituir Periférico #<span className="num">{replacingOldId}</span></h3>
          <div className="grid grid-cols-2 gap-4">
            <div className="flex flex-col gap-1.5">
              <Label>Novo Periférico (ID)</Label>
              <Select value={replaceNewId} onValueChange={setReplaceNewId}>
                <SelectTrigger><SelectValue placeholder="Selecione o substituto" /></SelectTrigger>
                <SelectContent>
                  {availablePeripherals.map(p => (
                    <SelectItem key={p.id} value={String(p.id)}>
                      #{p.id} — {p.tipo} {p.brand} {p.model}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label>Motivo da Substituição *</Label>
              <Input value={replaceReason} onChange={e => setReplaceReason(e.target.value)} placeholder="Ex: Defeito, Upgrade..." />
            </div>
          </div>
          <FileUpload
            accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
            onFile={setReplaceAttachment}
            label="Comprovante (opcional)"
          />
          <div className="flex gap-3">
            <Button
              disabled={!replaceNewId || !replaceReason || replaceMutation.isPending}
              onClick={() => replaceMutation.mutate()}
            >
              {replaceMutation.isPending ? 'Substituindo...' : 'Confirmar Substituição'}
            </Button>
            <Button variant="ghost" onClick={() => { setReplacingLinkId(null); setReplacingOldId(null) }}>
              Cancelar
            </Button>
          </div>
        </div>
      )}
    </div>
  )
}
