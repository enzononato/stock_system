import { useState, type FormEvent } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Trash2, ArrowLeft, ShieldAlert } from 'lucide-react'
import { useNavigate } from 'react-router-dom'
import { toast } from '@/components/ui/toast'

import { listItemsPaginated, removeItem } from '@/api/items'
import { SearchableSelect } from '@/components/ui/SearchableSelect'
import { FileUpload } from '@/components/ui/FileUpload'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { Checkbox } from '@/components/ui/checkbox'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { useConstants } from '@/hooks/useConstants'
import { getErrorMessage } from '@/lib/api-error'
import { formatDate } from '@/lib/utils'

const FETCH_ALL_LIMIT = 500

export default function RemovePage() {
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const { removalReasons, removalReasonsAttachment, isLoading: constantsLoading } = useConstants()

  const [selectedId, setSelectedId] = useState('')
  const [reason, setReason] = useState('')
  const [attachment, setAttachment] = useState<File | null>(null)
  const [confirmedUnderstanding, setConfirmedUnderstanding] = useState(false)

  const { data } = useQuery({
    queryKey: ['items', 'remove-all'],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  })
  const items = data?.items ?? []
  const disponiveis = items.filter((i) => i.status === 'Disponível')

  const selectedItem = disponiveis.find((i) => String(i.id) === selectedId)
  const needsAttachment = reason ? Boolean(removalReasonsAttachment[reason]) : false

  const mutation = useMutation({
    mutationFn: () => removeItem(Number(selectedId), reason, attachment ?? undefined),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setSelectedId('')
      setReason('')
      setAttachment(null)
      setConfirmedUnderstanding(false)
      toast.success('Equipamento baixado e removido do estoque com sucesso.')
      void navigate('/stock')
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao remover item do estoque.'))
    },
  })

  function handleSubmit(e: FormEvent) {
    e.preventDefault()
    if (!selectedId) {
      toast.error('Selecione um equipamento para remoção.')
      return
    }
    if (!reason) {
      toast.error('Selecione o motivo da baixa.')
      return
    }
    if (needsAttachment && !attachment) {
      toast.error('Este motivo de baixa exige o upload de comprovante (laudo/B.O./nota).')
      return
    }
    if (!confirmedUnderstanding) {
      toast.error('Confirme a ciência de irreversibilidade desta operação.')
      return
    }
    mutation.mutate()
  }

  return (
    <div className="space-y-6 max-w-4xl mx-auto">
      {/* Header */}
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
        {/* Aviso de Irreversibilidade */}
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

        {/* Formulário de Decisão Crítica */}
        <form
          onSubmit={handleSubmit}
          className="rounded border border-border bg-surface p-6 space-y-6"
        >
          {/* Seleção do Ativo */}
          <div className="space-y-2">
            <Label className="text-caption font-medium text-foreground">
              Selecione o Equipamento Disponível *
            </Label>
            <SearchableSelect
              options={disponiveis.map((i) => ({
                value: String(i.id),
                label: `#${i.id} — ${i.tipo ?? ''} ${i.brand ?? ''} ${i.model ?? ''}`.trim(),
                subtitle: [i.revenda, i.identificador].filter(Boolean).join(' • '),
              }))}
              value={selectedId}
              onValueChange={setSelectedId}
              placeholder="Pesquisar por ID, modelo, serial ou revenda…"
              searchPlaceholder="Digite para filtrar equipamentos disponíveis…"
            />
          </div>

          {/* Resumo do Ativo Selecionado */}
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
                  <span className="font-medium text-foreground">{selectedItem.revenda || '—'}</span>
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
                  <Badge
                    variant="success"
                    showDot
                    className="font-mono text-[10px] border-border text-foreground"
                  >
                    DISPONÍVEL
                  </Badge>
                </div>
              </div>
            </div>
          )}

          {/* Motivo da Baixa */}
          <div className="space-y-2">
            <Label className="text-caption font-medium text-foreground">
              Motivo Formal da Remoção *
            </Label>
            <Select value={reason} onValueChange={setReason} required disabled={constantsLoading}>
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

          {/* Comprovante / Laudo */}
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

          {/* Confirmação de 2 Etapas */}
          <div className="space-y-3 pt-4 border-t border-border">
            <label className="flex items-start gap-3 cursor-pointer select-none">
              <Checkbox
                checked={confirmedUnderstanding}
                onCheckedChange={(checked) => setConfirmedUnderstanding(Boolean(checked))}
                className="mt-0.5"
              />
              <span className="text-caption text-foreground leading-snug">
                Estou ciente de que a remoção do equipamento{' '}
                <strong className="font-mono">#{selectedId || '—'}</strong> é definitiva,
                irreversível e será registrada em auditoria sob minhas credenciais.
              </span>
            </label>
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
              disabled={
                mutation.isPending ||
                !selectedId ||
                !reason ||
                (needsAttachment && !attachment) ||
                !confirmedUnderstanding
              }
            >
              <Trash2 className="mr-1.5 size-3.5" />
              {mutation.isPending ? 'Processando…' : 'Confirmar Baixa Definitiva'}
            </Button>
          </div>
        </form>
      </div>
    </div>
  )
}
