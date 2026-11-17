import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { createLinha, OPERADORAS, updateLinha, type Linha, type LinhaInput } from '@/api/linhas'
import { listUnidades } from '@/api/unidades'
import { AparelhoSelect } from '@/components/linhas/AparelhoSelect'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { custoParaInput, descreverAparelho, parseCusto } from '@/features/linhas/format'
import { getErrorMessage } from '@/lib/api-error'
import { formatTelefone, upper } from '@/lib/format'

const selectClass =
  'flex h-9 w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

const MSG_CUSTO = 'Informe o custo no formato 0,00.'

const apenasDigitos = (v: string | null | undefined) => (v ?? '').replace(/\D/g, '')

interface LinhaFormDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  /** `null` = cadastro de uma nova linha; senão, edição. */
  linha: Linha | null
}

export function LinhaFormDialog({ open, onOpenChange, linha }: LinhaFormDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[90vh] max-w-2xl overflow-y-auto">
        <FormBody linha={linha} onClose={() => onOpenChange(false)} />
      </DialogContent>
    </Dialog>
  )
}

function FormBody({ linha, onClose }: { linha: Linha | null; onClose: () => void }) {
  const queryClient = useQueryClient()
  const editando = linha !== null

  const [form, setForm] = useState({
    numero: formatTelefone(linha?.numero),
    operadora: linha?.operadora ?? '',
    plano: linha?.plano ?? '',
    custo: custoParaInput(linha?.custo_mensal),
    unidadeId: linha ? String(linha.unidade_id) : '',
    iccid: linha?.iccid ?? '',
    observacoes: linha?.observacoes ?? '',
    motivo: '',
  })
  const [itemId, setItemId] = useState<number | null>(linha?.item_id ?? null)
  const [colaboradorId, setColaboradorId] = useState<number | null>(linha?.colaborador_id ?? null)
  const [erro, setErro] = useState<string | null>(null)
  const set = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) => {
    setForm((f) => ({ ...f, [key]: value }))
    setErro(null)
  }

  const { data: unidadesAtivas = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
  })
  // Na edição, a unidade atual continua selecionável mesmo que tenha sido desativada.
  const unidades =
    linha && !unidadesAtivas.some((u) => u.id === linha.unidade_id)
      ? [...unidadesAtivas, { id: linha.unidade_id, nome: linha.unidade_nome }]
      : unidadesAtivas

  const mutation = useMutation({
    mutationFn: (data: LinhaInput) => (linha ? updateLinha(linha.id, data) : createLinha(data)),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['linhas'] })
      toast(editando ? 'Linha atualizada.' : 'Linha cadastrada.')
      onClose()
    },
    onError: (err: unknown) =>
      setErro(getErrorMessage(err, editando ? 'Erro ao atualizar a linha.' : 'Erro ao cadastrar a linha.')),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setErro(null)

    const custo = parseCusto(form.custo)
    if (custo === null) {
      setErro(MSG_CUSTO)
      return
    }
    if (!form.unidadeId) {
      setErro('Escolha a unidade da linha.')
      return
    }
    const unidadeId = Number(form.unidadeId)
    const plano = form.plano.trim() || null
    const iccid = form.iccid.replace(/\s/g, '') || null
    const observacoes = form.observacoes.trim() || null

    if (!linha) {
      mutation.mutate({
        numero: form.numero.trim(),
        operadora: form.operadora,
        plano,
        custo_mensal: custo,
        unidade_id: unidadeId,
        item_id: itemId,
        colaborador_id: colaboradorId,
        iccid,
        observacoes,
      })
      return
    }

    // Edição: "chave ausente = sem mudança, null = limpar". Só vai o que mudou.
    const mudancas: LinhaInput = {}
    if (apenasDigitos(form.numero) !== apenasDigitos(linha.numero)) mudancas.numero = form.numero.trim()
    if (form.operadora !== linha.operadora) mudancas.operadora = form.operadora
    if (plano !== (linha.plano ?? null)) mudancas.plano = plano
    if (custo !== parseCusto(String(linha.custo_mensal))) mudancas.custo_mensal = custo
    if (unidadeId !== linha.unidade_id) mudancas.unidade_id = unidadeId
    if (itemId !== (linha.item_id ?? null)) mudancas.item_id = itemId
    if (colaboradorId !== (linha.colaborador_id ?? null)) mudancas.colaborador_id = colaboradorId
    if (iccid !== (linha.iccid ?? null)) mudancas.iccid = iccid
    if (observacoes !== (linha.observacoes ?? null)) mudancas.observacoes = observacoes
    if (Object.keys(mudancas).length === 0) {
      setErro('Nenhuma alteração para salvar.')
      return
    }
    const motivo = form.motivo.trim()
    if (motivo) mudancas.motivo = motivo
    mutation.mutate(mudancas)
  }

  const rotuloAparelhoAtual =
    linha && linha.item_id !== null && itemId === linha.item_id
      ? `${descreverAparelho({
          tipo: linha.item_tipo,
          brand: linha.item_brand,
          model: linha.item_model,
          identificador: linha.item_identificador,
        })}${linha.item_ativo === false ? ' (baixado)' : ''}`
      : undefined

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>{editando ? `Editar linha ${formatTelefone(linha.numero)}` : 'Nova linha'}</DialogTitle>
        <DialogDescription>
          {editando
            ? 'Trocas de aparelho, usuário e operadora ficam registradas no histórico da linha.'
            : 'Cadastre a linha com o número, o plano e, se houver, o aparelho e o colaborador.'}
        </DialogDescription>
      </DialogHeader>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-numero">Número *</Label>
          <Input
            id="linha-numero"
            inputMode="tel"
            value={form.numero}
            onChange={(e) => set('numero', e.target.value)}
            placeholder="(00) 00000-0000"
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-operadora">Operadora *</Label>
          <select
            id="linha-operadora"
            className={selectClass}
            value={form.operadora}
            onChange={(e) => set('operadora', e.target.value)}
          >
            <option value="">Selecione</option>
            {OPERADORAS.map((o) => (
              <option key={o} value={o}>
                {o}
              </option>
            ))}
          </select>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-plano">Plano</Label>
          <Input id="linha-plano" value={form.plano} onChange={(e) => set('plano', e.target.value)} maxLength={150} />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-custo">Custo mensal *</Label>
          <Input
            id="linha-custo"
            inputMode="decimal"
            value={form.custo}
            onChange={(e) => set('custo', e.target.value)}
            placeholder="0,00"
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-unidade">Unidade *</Label>
          <select
            id="linha-unidade"
            className={selectClass}
            value={form.unidadeId}
            onChange={(e) => set('unidadeId', e.target.value)}
          >
            <option value="">Selecione</option>
            {unidades.map((u) => (
              <option key={u.id} value={u.id}>
                {upper(u.nome)}
              </option>
            ))}
          </select>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="linha-iccid">ICCID</Label>
          <Input id="linha-iccid" value={form.iccid} onChange={(e) => set('iccid', e.target.value)} maxLength={30} />
        </div>
        <div className="flex flex-col gap-1.5 sm:col-span-2">
          <Label id="linha-aparelho-label">Aparelho (celular ou tablet)</Label>
          <div role="group" aria-labelledby="linha-aparelho-label">
            <AparelhoSelect
              value={itemId}
              onChange={(a) => {
                setItemId(a?.id ?? null)
                setErro(null)
              }}
              linhaId={linha?.id}
              rotuloAtual={rotuloAparelhoAtual}
            />
          </div>
        </div>
        <div className="flex flex-col gap-1.5 sm:col-span-2">
          <Label id="linha-colaborador-label">Colaborador</Label>
          <div role="group" aria-labelledby="linha-colaborador-label">
            <ColaboradorSelect
              value={colaboradorId}
              onChange={(c) => {
                setColaboradorId(c?.id ?? null)
                setErro(null)
              }}
              placeholder="Sem colaborador"
            />
          </div>
        </div>
        <div className="flex flex-col gap-1.5 sm:col-span-2">
          <Label htmlFor="linha-observacoes">Observações</Label>
          <textarea
            id="linha-observacoes"
            rows={2}
            maxLength={500}
            className="flex w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring"
            value={form.observacoes}
            onChange={(e) => set('observacoes', e.target.value)}
          />
        </div>
        {editando && (
          <div className="flex flex-col gap-1.5 sm:col-span-2">
            <Label htmlFor="linha-motivo">Motivo da alteração (opcional)</Label>
            <Input
              id="linha-motivo"
              value={form.motivo}
              onChange={(e) => set('motivo', e.target.value)}
              maxLength={255}
            />
          </div>
        )}
      </div>

      {erro && (
        <p role="alert" className="text-body-sm text-destructive">
          {erro}
        </p>
      )}

      <DialogFooter className="gap-2 sm:gap-0">
        <Button type="button" variant="ghost" onClick={onClose}>
          Cancelar
        </Button>
        <Button type="submit" loading={mutation.isPending}>
          Salvar
        </Button>
      </DialogFooter>
    </form>
  )
}
