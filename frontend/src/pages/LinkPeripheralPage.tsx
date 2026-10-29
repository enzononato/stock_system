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
import { Link2, Unlink, RefreshCw } from 'lucide-react'

// Tipos de equipamento que aceitam periféricos. Regra própria desta tela
// (subconjunto curado, não um cadastro de domínio), por isso continua fixa aqui.
const LINK_ALLOWED_TYPES = ['Desktop', 'Notebook', 'Switch', 'Impressora']

// A lista de equipamentos é de SELEÇÃO — o operador procura o equipamento a
// que vai vincular o periférico. Por isso liga a busca (com debounce) ao
// parâmetro `search` do servidor, com `limit` pequeno, em vez de buscar
// centenas de itens e filtrar no cliente. `GET /api/items` só aceita um `tipo`
// por chamada, e esta tela precisa de quatro (LINK_ALLOWED_TYPES): dispara uma
// busca pequena por tipo em paralelo, cada uma já filtrada e limitada no servidor.
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
    <div className="flex items-center justify-between border border-border rounded p-3 bg-surface hover:bg-surface-alt transition-colors">
      <div className="flex flex-col gap-0.5 min-w-0">
        <span className="text-body-sm font-medium text-foreground truncate">
          {peripheral.tipo} — {peripheral.brand || '—'} {peripheral.model || ''}
        </span>
        <span className="text-[11px] text-muted-foreground font-mono">
          S/N: {peripheral.identificador || '—'}
        </span>
      </div>
      <div className="flex items-center gap-3 shrink-0">
        <Badge
          variant={peripheral.status === 'Disponível' ? 'success' : peripheral.status === 'Em Uso' ? 'warning' : 'danger'}
          showDot
          className="text-[10px]"
        >
          {peripheral.status}
        </Badge>
        <Button
          size="sm"
          variant={variant === 'destructive' ? 'outline' : 'default'}
          onClick={action}
          className={variant === 'destructive' ? 'text-destructive border-border hover:bg-surface-alt text-xs h-7' : 'text-xs h-7'}
        >
          {actionIcon}
          <span className="ml-1">{actionLabel}</span>
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
  // disparar uma requisição por tecla.
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
  // da busca atual, então some da lista assim que o operador digita outra
  // coisa. Guarda o item completo no momento da seleção para o rótulo do
  // seletor e a linha "Status / Unidade" continuarem mostrando o equipamento
  // escolhido, mesmo depois que a busca mudar.
  const [selectedItemSnapshot, setSelectedItemSnapshot] = useState<Item | null>(null)
  function handleSelectItem(val: string) {
    setSelectedItemId(val)
    const found = linkableItems.find((i) => String(i.id) === val)
    if (found) setSelectedItemSnapshot(found)
  }

  const selectableItems = [
    ...(selectedItemSnapshot && !linkableItems.some((i) => i.id === selectedItemSnapshot.id)
      ? [selectedItemSnapshot]
      : []),
    ...linkableItems,
  ]

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
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast.success('Periférico vinculado!')
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao vincular.'))
    },
  })

  const unlinkMutation = useMutation({
    mutationFn: (linkId: number) => unlinkPeripheral(linkId),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast.success('Periférico desvinculado.')
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao desvincular.'))
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
      void queryClient.invalidateQueries({ queryKey: ['item-peripherals', selectedItemId] })
      void queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setReplacingLinkId(null)
      setReplacingOldId(null)
      setReplaceNewId('')
      setReplaceReason('')
      setReplaceAttachment(null)
      toast.success('Periférico substituído com sucesso!')
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao substituir.'))
    },
  })

  const selectedItem = selectableItems.find((i) => String(i.id) === selectedItemId)

  return (
    <div className="space-y-6 max-w-6xl mx-auto">
      {/* Cabeçalho */}
      <div className="border-b border-border pb-4">
        <div className="flex items-center gap-2">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            Ativos de Suporte
          </span>
          <span className="text-muted-foreground">•</span>
          <span className="text-caption text-foreground font-medium">Vínculo de Periféricos</span>
        </div>
        <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
          Vincular Periféricos
        </h1>
        <p className="text-body-sm text-muted-foreground mt-1">
          Associe periféricos a equipamentos como desktops, notebooks, switches e impressoras.
        </p>
      </div>

      {/* Seletor de equipamento */}
      <div className="rounded border border-border bg-surface p-5 space-y-3">
        <Label className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
          Selecione o Equipamento Principal
        </Label>
        <div className="max-w-xl">
          {/* Busca no servidor (debounce de 400ms) controlando o próprio campo
              de busca do dropdown via `search`/`onSearchChange`. */}
          <SearchableSelect
            options={selectableItems.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`.trim(),
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedItemId}
            onValueChange={handleSelectItem}
            placeholder="Selecione ou busque um equipamento…"
            searchPlaceholder="Buscar por marca, modelo ou patrimônio…"
            search={itemSearch}
            onSearchChange={setItemSearch}
          />
        </div>
        {selectedItem && (
          <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-caption text-muted-foreground pt-1">
            <span>
              Status: <strong className="text-foreground">{selectedItem.status || '—'}</strong>
            </span>
            <span aria-hidden="true">•</span>
            <span>
              Unidade: <strong className="text-foreground">{selectedItem.revenda || '—'}</strong>
            </span>
            {selectedItem.assigned_to && (
              <>
                <span aria-hidden="true">•</span>
                <span>
                  Alocado: <strong className="text-foreground">{selectedItem.assigned_to}</strong>
                </span>
              </>
            )}
          </div>
        )}
      </div>

      {selectedItemId && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Periféricos vinculados */}
          <div className="rounded border border-border bg-surface p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-border pb-3">
              <div>
                <h2 className="text-body font-semibold text-foreground">
                  Periféricos Vinculados ({linkedPeripherals.length})
                </h2>
                <p className="text-caption text-muted-foreground">Componentes conectados a este ativo.</p>
              </div>
              <Button
                size="sm"
                variant="ghost"
                onClick={() => void refetchLinked()}
                className="size-8 p-0"
                aria-label="Atualizar lista"
              >
                <RefreshCw className="size-3.5" />
              </Button>
            </div>

            {linkedPeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-12 text-center">Nenhum periférico vinculado.</p>
            ) : (
              <div className="space-y-2.5">
                {linkedPeripherals.map((p) => (
                  <div key={p.link_id} className="flex flex-col gap-1">
                    <PeripheralCard
                      peripheral={p}
                      action={() => unlinkMutation.mutate(p.link_id!)}
                      actionLabel="Desvincular"
                      actionIcon={<Unlink className="size-3" />}
                      variant="destructive"
                    />
                    <button
                      className="text-xs text-muted-foreground hover:text-foreground hover:underline text-left ml-1 cursor-pointer"
                      onClick={() => {
                        setReplacingLinkId(p.link_id!)
                        setReplacingOldId(p.id)
                      }}
                    >
                      Substituir por outro…
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Periféricos disponíveis */}
          <div className="rounded border border-border bg-surface p-5 space-y-4">
            <div className="border-b border-border pb-3">
              <h2 className="text-body font-semibold text-foreground">
                Periféricos Disponíveis ({availablePeripherals.length})
              </h2>
              <p className="text-caption text-muted-foreground">Itens em estoque livres para vinculação.</p>
            </div>

            {availablePeripherals.length === 0 ? (
              <p className="text-body-sm text-muted-foreground py-12 text-center">Nenhum periférico disponível.</p>
            ) : (
              <div className="space-y-2.5 max-h-[440px] overflow-y-auto pr-1">
                {availablePeripherals.map((p) => (
                  <PeripheralCard
                    key={p.id}
                    peripheral={p}
                    action={() => linkMutation.mutate(p.id)}
                    actionLabel="Vincular"
                    actionIcon={<Link2 className="size-3" />}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Painel de substituição */}
      {replacingLinkId && replacingOldId && (
        <div className="rounded border border-border bg-surface-alt p-6 space-y-4">
          <div className="border-b border-border pb-2">
            <h3 className="text-body-lg font-semibold text-foreground">
              Substituição de Periférico #{replacingOldId}
            </h3>
            <p className="text-caption text-muted-foreground">
              Selecione o novo periférico que assumirá o vínculo e informe a justificativa.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Novo Periférico Substituto *</Label>
              <Select value={replaceNewId} onValueChange={setReplaceNewId}>
                <SelectTrigger className="h-9">
                  <SelectValue placeholder="Selecione o substituto" />
                </SelectTrigger>
                <SelectContent>
                  {availablePeripherals.map((p) => (
                    <SelectItem key={p.id} value={String(p.id)}>
                      #{p.id} — {p.tipo} {p.brand} {p.model}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label className="text-caption font-medium">Motivo da Substituição *</Label>
              <Input
                value={replaceReason}
                onChange={(e) => setReplaceReason(e.target.value)}
                placeholder="Ex: Defeito, Upgrade, Avaria…"
                className="h-9"
              />
            </div>
          </div>

          <div className="pt-2">
            <FileUpload
              accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
              onFile={setReplaceAttachment}
              label="Comprovante / laudo (opcional)"
            />
          </div>

          <div className="flex justify-end gap-3 pt-2 border-t border-border">
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                setReplacingLinkId(null)
                setReplacingOldId(null)
              }}
            >
              Cancelar
            </Button>
            <Button
              size="sm"
              disabled={!replaceNewId || !replaceReason || replaceMutation.isPending}
              onClick={() => replaceMutation.mutate()}
            >
              {replaceMutation.isPending ? 'Substituindo…' : 'Confirmar Substituição'}
            </Button>
          </div>
        </div>
      )}
    </div>
  )
}
