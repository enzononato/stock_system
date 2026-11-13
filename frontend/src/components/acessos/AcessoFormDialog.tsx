import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { createAcesso, listSistemas, updateAcesso, type Acesso } from '@/api/acessos'
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

const selectClass =
  'flex h-9 w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

interface AcessoFormDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  colaboradorId: number
  /** `null` = registrar um acesso novo; senão, edição (o sistema não muda). */
  acesso: Acesso | null
}

export function AcessoFormDialog({ open, onOpenChange, colaboradorId, acesso }: AcessoFormDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md">
        <FormBody
          key={acesso?.id ?? 'novo'}
          colaboradorId={colaboradorId}
          acesso={acesso}
          onClose={() => onOpenChange(false)}
        />
      </DialogContent>
    </Dialog>
  )
}

function FormBody({
  colaboradorId,
  acesso,
  onClose,
}: {
  colaboradorId: number
  acesso: Acesso | null
  onClose: () => void
}) {
  const queryClient = useQueryClient()
  const [sistemaId, setSistemaId] = useState('')
  const [login, setLogin] = useState(acesso?.login ?? '')
  const [itemId, setItemId] = useState<number | null>(acesso?.item_id ?? null)
  const [erro, setErro] = useState<string | null>(null)

  const { data: sistemas = [] } = useQuery({
    queryKey: ['acessos', 'sistemas', { include_inactive: false }],
    queryFn: () => listSistemas(),
    enabled: acesso === null,
  })

  const mutation = useMutation({
    mutationFn: () =>
      acesso
        ? updateAcesso(acesso.id, { login: login.trim(), item_id: itemId })
        : createAcesso({
            colaborador_id: colaboradorId,
            sistema_id: Number(sistemaId),
            login: login.trim(),
            item_id: itemId,
          }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast(acesso ? 'Acesso atualizado.' : 'Acesso registrado.')
      onClose()
    },
    onError: (err: unknown) =>
      toast(getErrorMessage(err, acesso ? 'Erro ao atualizar o acesso.' : 'Erro ao registrar o acesso.'), 'error'),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (mutation.isPending) return
    if (!acesso && !sistemaId) {
      setErro('Selecione o sistema.')
      return
    }
    if (!login.trim()) {
      setErro('Informe o login.')
      return
    }
    setErro(null)
    mutation.mutate()
  }

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>{acesso ? `Editar acesso: ${acesso.sistema_nome}` : 'Adicionar acesso'}</DialogTitle>
        <DialogDescription>Registre o login do colaborador no sistema e, se quiser, o equipamento.</DialogDescription>
      </DialogHeader>

      <div className="flex flex-col gap-1.5">
        <Label htmlFor="acesso-sistema">Sistema</Label>
        {acesso ? (
          <Input id="acesso-sistema" value={acesso.sistema_nome} readOnly disabled />
        ) : (
          <select
            id="acesso-sistema"
            className={selectClass}
            value={sistemaId}
            onChange={(e) => setSistemaId(e.target.value)}
          >
            <option value="">Selecione...</option>
            {sistemas.map((s) => (
              <option key={s.id} value={s.id}>
                {s.nome}
              </option>
            ))}
          </select>
        )}
      </div>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="acesso-login">Login</Label>
        <Input
          id="acesso-login"
          value={login}
          onChange={(e) => setLogin(e.target.value)}
          maxLength={150}
          autoComplete="off"
        />
      </div>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="acesso-equipamento">Equipamento (opcional)</Label>
        <EquipamentoDoColaboradorSelect
          id="acesso-equipamento"
          colaboradorId={colaboradorId}
          value={itemId}
          onChange={setItemId}
        />
      </div>

      <AvisoSemSenha />

      {erro && (
        <p role="alert" className="text-sm text-destructive">
          {erro}
        </p>
      )}

      <DialogFooter>
        <Button type="button" variant="outline" onClick={onClose}>
          Cancelar
        </Button>
        <Button type="submit" loading={mutation.isPending}>
          Salvar
        </Button>
      </DialogFooter>
    </form>
  )
}
