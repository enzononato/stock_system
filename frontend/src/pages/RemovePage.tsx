import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { getItem, listItemsPaginated, removeItem } from '@/api/items'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { FileUpload } from '@/components/ui/FileUpload'
import { Checkbox } from '@/components/ui/checkbox'
import { StatusBadge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { useConstants } from '@/hooks/useConstants'
import { formatDate } from '@/lib/utils'
import { Trash2, ShieldAlert } from 'lucide-react'

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
  // Confirmação de ciência: este é o portão da única ação irreversível do
  // sistema — o botão de confirmar só destrava com os três requisitos
  // preenchidos (equipamento, motivo e esta ciência) ao mesmo tempo.
  const [confirmedUnderstanding, setConfirmedUnderstanding] = useState(false)

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

  // Ficha completa do ativo selecionado, buscada por id direto no servidor —
  // não a partir de `disponiveis` (que é só a página pequena, filtrada pela
  // busca, vinda do servidor). Sem isto, a ficha sumiria assim que o usuário
  // digitasse algo novo na busca depois de já ter selecionado um item.
  const { data: selectedItem } = useQuery({
    queryKey: ['items', 'detail', selectedId],
    queryFn: () => getItem(Number(selectedId)),
    enabled: !!selectedId,
  })

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
      setConfirmedUnderstanding(false)
      toast('Item removido do estoque.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao remover item.'), 'error')
    },
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!selectedId) { toast('Selecione um item.', 'error'); return }
    if (!reason) { toast('Selecione o motivo.', 'error'); return }
    if (needsAttachment && !attachment) { toast('Este motivo exige um comprovante.', 'error'); return }
    if (!confirmedUnderstanding) { toast('Confirme a ciência de irreversibilidade desta operação.', 'error'); return }
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

      {/* Aviso de irreversibilidade — reforça em texto o que o eyebrow do
          PageHeader já diz, antes de qualquer campo do formulário. */}
      <div className="figure-ground-panel flex items-start gap-3">
        <div className="p-2 rounded bg-foreground text-background shrink-0 mt-0.5">
          <ShieldAlert size={16} />
        </div>
        <div>
          <p className="text-body-sm font-semibold text-foreground">
            Ação permanente com rastro em auditoria
          </p>
          <p className="text-caption text-muted-foreground mt-0.5 leading-relaxed">
            A baixa remove o equipamento do estoque ativo e encerra seu ciclo patrimonial. Um
            registro histórico com data, operador e motivo é gerado e não pode ser desfeito.
          </p>
        </div>
      </div>

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

        {/* Ficha completa do ativo — sem isto o operador confirmava a baixa
            vendo só o texto curto da opção escolhida no dropdown. */}
        {selectedItem && (
          <div className="rounded border border-border bg-surface-alt p-4 space-y-3">
            <PanelHeader title="Ficha do Ativo Selecionado" />
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3">
              <div>
                <p className="text-caption text-muted-foreground">Tipo / Modelo</p>
                <p className="text-body-sm text-foreground font-medium">
                  {selectedItem.tipo} {selectedItem.brand} {selectedItem.model}
                </p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Unidade</p>
                <p className="text-body-sm text-foreground">{selectedItem.revenda || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Identificador</p>
                <p className="text-body-sm text-foreground num">{selectedItem.identificador || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Nota Fiscal</p>
                <p className="text-body-sm text-foreground num">{selectedItem.nota_fiscal || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Data Cadastro</p>
                <p className="text-body-sm text-foreground num">{formatDate(selectedItem.date_registered)}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Status</p>
                <StatusBadge status={selectedItem.status} />
              </div>
            </div>
          </div>
        )}

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

        {/* Confirmação de ciência — o portão final antes da baixa definitiva. */}
        <div className="flex items-start gap-2.5 border-t border-border pt-4">
          <Checkbox
            checked={confirmedUnderstanding}
            onCheckedChange={setConfirmedUnderstanding}
            className="mt-0.5"
          />
          <p className="text-body-sm text-foreground leading-snug">
            Estou ciente de que a remoção do equipamento{' '}
            <strong className="num">#{selectedId || '—'}</strong> é definitiva, irreversível e
            será registrada em auditoria sob minhas credenciais.
          </p>
        </div>

        <Button
          type="submit"
          variant="destructive"
          disabled={mutation.isPending || !selectedId || !reason || !confirmedUnderstanding}
          className="w-full"
        >
          <Trash2 size={14} />
          {mutation.isPending ? 'Removendo...' : 'Confirmar Remoção'}
        </Button>
      </form>
    </div>
  )
}
