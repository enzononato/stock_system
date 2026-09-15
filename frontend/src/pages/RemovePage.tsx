import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { listItemsPaginated, removeItem } from '@/api/items'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { PageHeader } from '@/components/layout/PageHeader'
import { useConstants } from '@/hooks/useConstants'
import { Trash2 } from 'lucide-react'

// Esta tela só tem um bloco de dados, e é uma lista de SELEÇÃO — o operador
// procura o equipamento "Disponível" a remover, não folheia página por
// página. Por isso liga a busca (com debounce) ao parâmetro `search` do
// servidor, com `limit` pequeno, em vez de buscar os 500 primeiros e filtrar
// no cliente (T3).
const ITEM_SEARCH_LIMIT = 20

export default function RemovePage() {
  const queryClient = useQueryClient()
  const { removalReasons, removalReasonsAttachment, isLoading: constantsLoading } = useConstants()
  const [selectedId, setSelectedId] = useState('')
  const [reason, setReason] = useState('')
  const [attachment, setAttachment] = useState<File | null>(null)

  // Busca do equipamento disponível: `itemSearch` é o que o operador digita,
  // `itemSearchAplicado` é o que vai para o servidor, com atraso, para não
  // disparar uma requisição por tecla (mesmo padrão do debounce de busca já
  // usado em HistoryPage/StockPage).
  const [itemSearch, setItemSearch] = useState('')
  const [itemSearchAplicado, setItemSearchAplicado] = useState('')
  useEffect(() => {
    const timer = setTimeout(() => setItemSearchAplicado(itemSearch.trim()), 400)
    return () => clearTimeout(timer)
  }, [itemSearch])

  const { data } = useQuery({
    queryKey: ['items', 'select', 'Disponível', itemSearchAplicado],
    queryFn: () =>
      listItemsPaginated({ status: 'Disponível', search: itemSearchAplicado || undefined, limit: ITEM_SEARCH_LIMIT }),
    placeholderData: keepPreviousData,
  })
  const disponiveis = data?.items ?? []

  // T2: o mapa de quais motivos exigem comprovante vem de /api/constants
  // (removalReasonsAttachment), não mais de um objeto hardcoded local.
  const needsAttachment = reason ? removalReasonsAttachment[reason] : false

  const mutation = useMutation({
    mutationFn: () => removeItem(Number(selectedId), reason, attachment ?? undefined),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setSelectedId('')
      setReason('')
      setAttachment(null)
      toast('Item removido do estoque.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao remover item.'
      toast(msg, 'error')
    },
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!selectedId) { toast('Selecione um item.', 'error'); return }
    if (!reason) { toast('Selecione o motivo.', 'error'); return }
    if (needsAttachment && !attachment) { toast('Este motivo exige um comprovante.', 'error'); return }
    mutation.mutate()
  }

  return (
    <div className="page-container-reading space-y-6">
      <PageHeader
        eyebrow="Baixa de Ativos"
        eyebrowDetail="Ação Irreversível"
        title="Remover Equipamento"
        description="Remove permanentemente o item do estoque. Esta ação não pode ser desfeita."
      />

      <form onSubmit={handleSubmit} className="figure-ground-panel space-y-4">
        <div className="flex flex-col gap-1.5">
          <Label>Equipamento *</Label>
          {/* Busca no servidor (debounce de 400ms) controlando o próprio campo
              de busca do dropdown via `search`/`onSearchChange` — sem isso, o
              usuário veria uma segunda caixa de busca (a interna do
              SearchableSelect) além desta, filtrando só os 20 já carregados. */}
          <SearchableSelect
            options={disponiveis.map((i) => ({
              value: String(i.id),
              label: `#${i.id} — ${i.tipo} ${i.brand || ''} ${i.model || ''}`,
              subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
            }))}
            value={selectedId}
            onValueChange={setSelectedId}
            placeholder="Selecione um equipamento disponível..."
            searchPlaceholder="Buscar por marca, modelo ou identificador (patrimônio)..."
            search={itemSearch}
            onSearchChange={setItemSearch}
          />
        </div>

        <div className="flex flex-col gap-1.5">
          <Label>Motivo da Remoção *</Label>
          <Select value={reason} onValueChange={setReason} required disabled={constantsLoading}>
            <SelectTrigger><SelectValue placeholder="Selecione o motivo" /></SelectTrigger>
            <SelectContent>
              {removalReasons.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}
            </SelectContent>
          </Select>
        </div>

        {reason && (
          <div className="flex flex-col gap-1.5">
            <Label>
              Comprovante {needsAttachment ? '(obrigatório)' : '(opcional)'}
            </Label>
            <FileUpload
              accept={{ 'application/pdf': ['.pdf'], 'image/*': ['.jpg', '.jpeg', '.png'] }}
              onFile={setAttachment}
              label="Arraste ou clique para anexar comprovante"
            />
          </div>
        )}

        <Button type="submit" variant="destructive" disabled={mutation.isPending} className="w-full">
          <Trash2 size={14} />
          {mutation.isPending ? 'Removendo...' : 'Confirmar Remoção'}
        </Button>
      </form>
    </div>
  )
}
