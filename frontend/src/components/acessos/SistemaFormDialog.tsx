import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { createSistema, updateSistema, type Sistema } from '@/api/acessos'
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

interface SistemaFormDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  /** `null` = novo sistema; senão, edição. */
  sistema: Sistema | null
}

export function SistemaFormDialog({ open, onOpenChange, sistema }: SistemaFormDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md">
        <FormBody key={sistema?.id ?? 'novo'} sistema={sistema} onClose={() => onOpenChange(false)} />
      </DialogContent>
    </Dialog>
  )
}

function FormBody({ sistema, onClose }: { sistema: Sistema | null; onClose: () => void }) {
  const queryClient = useQueryClient()
  const [nome, setNome] = useState(sistema?.nome ?? '')
  const [descricao, setDescricao] = useState(sistema?.descricao ?? '')

  const mutation = useMutation({
    mutationFn: () =>
      sistema
        ? updateSistema(sistema.id, { nome: nome.trim(), descricao: descricao.trim() || null })
        : createSistema({ nome: nome.trim(), descricao: descricao.trim() || null }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast(sistema ? 'Sistema atualizado.' : 'Sistema cadastrado.')
      onClose()
    },
    onError: (err: unknown) =>
      toast(getErrorMessage(err, sistema ? 'Erro ao atualizar sistema.' : 'Erro ao cadastrar sistema.'), 'error'),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (!nome.trim()) {
      toast('Informe o nome do sistema.', 'error')
      return
    }
    mutation.mutate()
  }

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>{sistema ? `Editar ${sistema.nome}` : 'Novo sistema'}</DialogTitle>
        <DialogDescription>
          Sistemas do catálogo entram nos pacotes por função e na ficha de acessos dos colaboradores.
        </DialogDescription>
      </DialogHeader>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="sistema-nome">Nome *</Label>
        <Input id="sistema-nome" value={nome} onChange={(e) => setNome(e.target.value)} maxLength={100} />
      </div>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="sistema-descricao">Descrição</Label>
        <Input
          id="sistema-descricao"
          value={descricao}
          onChange={(e) => setDescricao(e.target.value)}
          maxLength={255}
        />
      </div>
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
