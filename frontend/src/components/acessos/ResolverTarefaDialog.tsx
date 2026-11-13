import { useRef, useState } from 'react'
import { useMutation } from '@tanstack/react-query'
import { resolverTarefa, type EventoDetalhe, type Tarefa } from '@/api/acessos'
import { AvisoSemSenha } from '@/components/acessos/AvisoSemSenha'
import { EquipamentoDoColaboradorSelect } from '@/components/acessos/EquipamentoDoColaboradorSelect'
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
import { getErrorMessage } from '@/lib/api-error'

export type ModoResolucao = 'feito' | 'nao_se_aplica' | 'manter'

interface ResolverTarefaDialogProps {
  tarefa: Tarefa | null
  modo: ModoResolucao | null
  colaboradorId: number
  onClose: () => void
  onResolved: (evento: EventoDetalhe) => void
}

export function ResolverTarefaDialog({
  tarefa,
  modo,
  colaboradorId,
  onClose,
  onResolved,
}: ResolverTarefaDialogProps) {
  const aberto = tarefa !== null && modo !== null
  return (
    <Dialog open={aberto} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        {tarefa && modo && (
          <FormBody
            key={`${tarefa.id}-${modo}`}
            tarefa={tarefa}
            modo={modo}
            colaboradorId={colaboradorId}
            onClose={onClose}
            onResolved={onResolved}
          />
        )}
      </DialogContent>
    </Dialog>
  )
}

function FormBody({
  tarefa,
  modo,
  colaboradorId,
  onClose,
  onResolved,
}: {
  tarefa: Tarefa
  modo: ModoResolucao
  colaboradorId: number
  onClose: () => void
  onResolved: (evento: EventoDetalhe) => void
}) {
  const [login, setLogin] = useState('')
  const [itemId, setItemId] = useState<number | null>(null)
  const [observacao, setObservacao] = useState('')
  const [erro, setErro] = useState<string | null>(null)
  // Trava síncrona: dois cliques seguidos chegam antes de o React repintar o botão em `loading`.
  const enviando = useRef(false)

  const mutation = useMutation({
    mutationFn: () => {
      if (modo === 'feito') {
        return resolverTarefa(
          tarefa.id,
          tarefa.acao === 'criar'
            ? { status: 'feito', login: login.trim(), item_id: itemId }
            : { status: 'feito' }
        )
      }
      return resolverTarefa(tarefa.id, { status: modo, observacao: observacao.trim() })
    },
    onSuccess: (evento) => {
      toast('Tarefa atualizada.')
      onResolved(evento)
      onClose()
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao atualizar a tarefa.'), 'error'),
    onSettled: () => {
      enviando.current = false
    },
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (enviando.current) return
    if (modo === 'feito' && tarefa.acao === 'criar' && !login.trim()) {
      setErro('Informe o login.')
      return
    }
    if (modo === 'nao_se_aplica' && !observacao.trim()) {
      setErro('Informe a observação.')
      return
    }
    if (modo === 'manter' && !observacao.trim()) {
      setErro('Informe o motivo para manter o acesso.')
      return
    }
    setErro(null)
    enviando.current = true
    mutation.mutate()
  }

  const criar = modo === 'feito' && tarefa.acao === 'criar'
  const desativar = modo === 'feito' && tarefa.acao === 'desativar'
  const titulo =
    modo === 'manter'
      ? 'Manter acesso'
      : modo === 'nao_se_aplica'
        ? `Não se aplica: ${tarefa.sistema_nome}`
        : criar
          ? `Registrar acesso: ${tarefa.sistema_nome}`
          : `Desativar acesso: ${tarefa.sistema_nome}`

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>{titulo}</DialogTitle>
        <DialogDescription>
          {criar
            ? 'Informe o login criado no sistema. Equipamento é opcional.'
            : 'A tarefa só é concluída com a sua confirmação.'}
        </DialogDescription>
      </DialogHeader>

      {criar && (
        <>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="tarefa-login">Login</Label>
            <Input
              id="tarefa-login"
              value={login}
              onChange={(e) => setLogin(e.target.value)}
              maxLength={150}
              autoComplete="off"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="tarefa-equipamento">Equipamento (opcional)</Label>
            <EquipamentoDoColaboradorSelect
              id="tarefa-equipamento"
              colaboradorId={colaboradorId}
              value={itemId}
              onChange={setItemId}
            />
          </div>
          <AvisoSemSenha />
        </>
      )}

      {desativar && (
        <p className="text-sm">
          {`Confirme que o acesso "${tarefa.login ?? ''}" no sistema ${tarefa.sistema_nome} foi desativado.`}
        </p>
      )}

      {(modo === 'nao_se_aplica' || modo === 'manter') && (
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="tarefa-observacao">{modo === 'manter' ? 'Motivo' : 'Observação'}</Label>
          <Input
            id="tarefa-observacao"
            value={observacao}
            onChange={(e) => setObservacao(e.target.value)}
            maxLength={450}
          />
        </div>
      )}

      {erro && (
        <p role="alert" className="text-sm text-destructive">
          {erro}
        </p>
      )}

      <DialogFooter>
        <Button type="button" variant="outline" onClick={onClose}>
          Voltar
        </Button>
        <Button type="submit" loading={mutation.isPending}>
          {desativar ? 'Confirmar desativação' : 'Confirmar'}
        </Button>
      </DialogFooter>
    </form>
  )
}
