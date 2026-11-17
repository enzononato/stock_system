import { useEffect, useState, type FormEvent } from 'react'
import { useMutation, useQuery, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { Trash2, ArrowLeft, ShieldAlert } from 'lucide-react'
import { useNavigate, useSearchParams } from 'react-router-dom'
import { toast } from '@/components/ui/toast'

import { getItem, listItemsPaginated, removeItem } from '@/api/items'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { FileUpload } from '@/components/ui/FileUpload'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { StatusPill } from '@/components/ui/StatusPill'
import { Checkbox } from '@/components/ui/checkbox'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  ApagamentoDadosFields,
  apagamentoCompleto,
  apagamentoInicial,
  paraApagamentoInput,
  type ApagamentoState,
} from '@/features/ciclo-vida/ApagamentoDadosFields'
import { useConstants } from '@/hooks/useConstants'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import { formatDate } from '@/lib/utils'

// Esta tela só tem um bloco de dados, e é uma lista de SELEÇÃO — o operador
// procura o equipamento "Disponível" a remover, não folheia página por
// página. Por isso liga a busca (com debounce) ao parâmetro `search` do
// servidor, com `limit` pequeno, em vez de buscar os 500 primeiros e filtrar
// no cliente (T3).
const ITEM_SEARCH_LIMIT = 20

export default function RemovePage() {
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const [searchParams] = useSearchParams()
  const {
    removalReasons,
    removalReasonsAttachment,
    tiposExigemApagamento,
    motivosSemPosse,
    metodosApagamento,
    isLoading: constantsLoading,
  } = useConstants()

  // `?item=<id>` abre a tela com o equipamento já escolhido (atalho "Dar baixa" da aba
  // Manutenções para equipamento sem conserto, que não aparece na busca de Disponível).
  const [selectedId, setSelectedId] = useState(() => {
    const doLink = searchParams.get('item') ?? ''
    return /^\d+$/.test(doLink) ? doLink : ''
  })
  const [reason, setReason] = useState('')
  const [attachment, setAttachment] = useState<File | null>(null)
  const [apagamento, setApagamento] = useState<ApagamentoState>(apagamentoInicial)
  // Confirmação de ciência: este é o portão da única ação irreversível do
  // sistema — o botão de confirmar só destrava com os requisitos preenchidos
  // (equipamento, motivo, comprovante quando exigido e esta ciência).
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
  const needsAttachment = reason ? Boolean(removalReasonsAttachment[reason]) : false

  // Apagamento de dados (LGPD): obrigatório para os tipos com armazenamento; "não se aplica"
  // só vale para roubo ou perda. Os dois conjuntos vêm de /api/constants.
  const exige = tiposExigemApagamento.includes(selectedItem?.tipo ?? '')
  const permiteNaoAplica = motivosSemPosse.includes(reason)
  const apagamentoOk = !exige || apagamentoCompleto(apagamento, permiteNaoAplica)

  function trocarMotivo(novo: string) {
    setReason(novo)
    // Trocar para um motivo que não aceita mais "não se aplica" zera a escolha.
    if (!motivosSemPosse.includes(novo) && apagamento.dadosApagados === 'nao_aplica') {
      setApagamento({ ...apagamento, dadosApagados: '' })
    }
  }

  // O item vindo do link (`?item=`) pode não estar na lista de Disponíveis: entra como opção
  // para o gatilho mostrar o equipamento escolhido.
  const opcoes = [...disponiveis]
  if (selectedItem && !opcoes.some((i) => i.id === selectedItem.id)) opcoes.unshift(selectedItem)

  const mutation = useMutation({
    mutationFn: () =>
      removeItem(
        Number(selectedId),
        reason,
        attachment ?? undefined,
        exige ? paraApagamentoInput(apagamento) : undefined
      ),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setSelectedId('')
      setReason('')
      setAttachment(null)
      setApagamento(apagamentoInicial())
      setConfirmedUnderstanding(false)
      toast('Equipamento baixado e removido do estoque com sucesso.')
    },
    onError: (err: unknown) => {
      toast(getErrorMessage(err, 'Erro ao remover item do estoque.'), 'error')
    },
  })

  function handleSubmit(e: FormEvent) {
    e.preventDefault()
    if (!selectedId) {
      toast('Selecione um equipamento para remoção.', 'error')
      return
    }
    if (!reason) {
      toast('Selecione o motivo da baixa.', 'error')
      return
    }
    if (needsAttachment && !attachment) {
      toast('Este motivo de baixa exige o upload de comprovante (laudo/B.O./nota).', 'error')
      return
    }
    if (!apagamentoOk) {
      toast('Registre o apagamento dos dados do equipamento (LGPD).', 'error')
      return
    }
    if (!confirmedUnderstanding) {
      toast('Confirme a ciência de irreversibilidade desta operação.', 'error')
      return
    }
    mutation.mutate()
  }

  return (
    <div className="space-y-6 max-w-4xl mx-auto">
      {/* Cabeçalho */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Operação de Baixa
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Decisão Crítica</span>
          </div>
          <h1 className="text-heading-lg font-semibold tracking-tight text-foreground mt-0.5">
            Baixa de Equipamento
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Remoção permanente de patrimônio por descarte, furto, venda ou avaria irreversível.
          </p>
        </div>

        <Button
          variant="outline"
          size="sm"
          onClick={() => void navigate('/stock')}
          className="self-start"
        >
          <ArrowLeft className="mr-1.5 size-3.5" />
          Cancelar e Voltar
        </Button>
      </div>

      <div className="space-y-6">
        {/* Aviso de irreversibilidade */}
        <div className="rounded border border-border bg-surface p-4 flex items-start gap-3.5">
          <div className="p-2 rounded bg-foreground text-background shrink-0 mt-0.5">
            <ShieldAlert className="size-4" />
          </div>
          <div>
            <h2 className="text-body-sm font-semibold text-foreground">
              Ação Permanente com Rastreabilidade em Auditoria
            </h2>
            <p className="text-caption text-muted-foreground mt-0.5 leading-relaxed">
              A baixa remove o equipamento do estoque ativo e encerra seu ciclo patrimonial. Um
              registro histórico imutável com data, operador e motivo será gerado.
            </p>
          </div>
        </div>

        {/* Formulário de decisão crítica */}
        <form
          onSubmit={handleSubmit}
          className="rounded border border-border bg-surface p-6 space-y-6"
        >
          {/* Seleção do ativo */}
          <div className="space-y-2">
            <Label className="text-caption font-medium text-foreground">
              Selecione o Equipamento Disponível *
            </Label>
            {/* Busca no servidor (debounce de 400ms) controlando o próprio campo
                de busca do dropdown via `search`/`onSearchChange` — sem isso, o
                usuário veria uma segunda caixa de busca (a interna do
                SearchableSelect) filtrando só os 20 já carregados. O texto do
                gatilho cita só os campos que o servidor realmente pesquisa
                (marca, modelo, identificador) — o da FT_STC prometia busca por
                ID e revenda, que o `search` do servidor não cobre. */}
            <SearchableSelect
              options={opcoes.map((i) => ({
                value: String(i.id),
                label: `#${i.id} — ${i.tipo ?? ''} ${i.brand ?? ''} ${i.model ?? ''}`.trim(),
                subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
              }))}
              value={selectedId}
              onValueChange={setSelectedId}
              placeholder="Pesquisar por marca, modelo ou identificador…"
              searchPlaceholder="Digite para filtrar equipamentos disponíveis…"
              search={itemSearch}
              onSearchChange={setItemSearch}
            />
          </div>

          {/* Resumo do ativo selecionado */}
          {selectedItem && (
            <div className="rounded border border-border bg-surface-alt p-4 space-y-3">
              <div className="flex items-center justify-between border-b border-border pb-2">
                <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
                  Ficha do Ativo Selecionado
                </span>
                <span className="font-mono text-caption text-foreground font-semibold">
                  ID #{selectedItem.id}
                </span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-caption">
                <div>
                  <span className="text-muted-foreground block">Tipo / Modelo</span>
                  <span className="font-medium text-foreground">
                    {selectedItem.tipo} {selectedItem.brand} {selectedItem.model}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Revenda</span>
                  <span className="font-medium text-foreground">{upper(selectedItem.revenda) || '—'}</span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Identificador</span>
                  <span className="font-mono text-foreground">
                    {selectedItem.identificador || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Nota Fiscal</span>
                  <span className="font-mono text-foreground">
                    {selectedItem.nota_fiscal || '—'}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Data Cadastro</span>
                  <span className="font-mono text-foreground">
                    {formatDate(selectedItem.date_registered)}
                  </span>
                </div>
                <div>
                  <span className="text-muted-foreground block">Status</span>
                  {/* Status real vindo do servidor (a ficha é buscada por id). */}
                  <StatusPill status={selectedItem.status} />
                </div>
              </div>
            </div>
          )}

          {/* Motivo da baixa */}
          <div className="space-y-2">
            <Label className="text-caption font-medium text-foreground">
              Motivo Formal da Remoção *
            </Label>
            <Select value={reason} onValueChange={trocarMotivo} required disabled={constantsLoading}>
              <SelectTrigger className="h-9">
                <SelectValue placeholder="Selecione a justificativa da baixa" />
              </SelectTrigger>
              <SelectContent>
                {removalReasons.map((r) => (
                  <SelectItem key={r} value={r}>
                    {r}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          {/* Comprovante / laudo */}
          {reason && (
            <div className="space-y-2 pt-2 border-t border-border">
              <div className="flex items-center justify-between">
                <Label className="text-caption font-medium text-foreground">
                  Comprovante / Laudo Técnico{' '}
                  {needsAttachment ? '(Obrigatório para este motivo)' : '(Opcional)'}
                </Label>
                {needsAttachment && (
                  <Badge
                    variant="danger"
                    showDot
                    className="font-mono text-[10px] border-border text-foreground"
                  >
                    DOCUMENTO OBRIGATÓRIO
                  </Badge>
                )}
              </div>
              <FileUpload
                accept={{
                  'application/pdf': ['.pdf'],
                  'image/*': ['.png', '.jpg', '.jpeg'],
                }}
                onFile={setAttachment}
                label="Arraste ou clique para anexar laudo, B.O. ou comprovante de baixa (PDF / Imagem)"
              />
            </div>
          )}

          {/* Apagamento de dados (LGPD): tipos com armazenamento exigem o registro; os demais, não. */}
          {selectedItem && exige && (
            <ApagamentoDadosFields
              tipo={selectedItem.tipo ?? ''}
              value={apagamento}
              onChange={setApagamento}
              metodos={metodosApagamento}
              permiteNaoAplica={permiteNaoAplica}
            />
          )}
          {selectedItem && !exige && (
            <p className="text-caption text-muted-foreground">
              Este tipo de equipamento não armazena dados pessoais: o apagamento será registrado como "não se aplica".
            </p>
          )}

          {/* Confirmação de ciência — o portão final antes da baixa definitiva.
              Diferente da FT_STC, o texto NÃO fica dentro de um <label>
              envolvente: só o quadrado do Checkbox (já corrigido) é clicável. */}
          <div className="space-y-3 pt-4 border-t border-border">
            <div className="flex items-start gap-3 select-none">
              <Checkbox
                checked={confirmedUnderstanding}
                onCheckedChange={setConfirmedUnderstanding}
                className="mt-0.5"
              />
              <span className="text-caption text-foreground leading-snug">
                Estou ciente de que a remoção do equipamento{' '}
                <strong className="font-mono">#{selectedId || '—'}</strong> é definitiva,
                irreversível e será registrada em auditoria sob minhas credenciais.
              </span>
            </div>
          </div>

          {/* Ações */}
          <div className="pt-2 flex justify-end gap-3 border-t border-border">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => void navigate('/stock')}
            >
              Cancelar
            </Button>
            <Button
              type="submit"
              size="sm"
              loading={mutation.isPending}
              disabled={
                !selectedId ||
                !reason ||
                (needsAttachment && !attachment) ||
                !apagamentoOk ||
                !confirmedUnderstanding
              }
            >
              <Trash2 className="mr-1.5 size-3.5" />
              Confirmar Baixa Definitiva
            </Button>
          </div>
        </form>
      </div>
    </div>
  )
}
