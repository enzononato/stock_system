import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { alterarStatusLinha, type Linha } from '@/api/linhas'
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
import { STATUS_DA_ACAO, type AcaoStatusLinha } from '@/features/linhas/status'
import { getErrorMessage } from '@/lib/api-error'
import { formatTelefone } from '@/lib/format'

const TEXTOS: Record<AcaoStatusLinha, { titulo: string; confirmar: string; sucesso: string; aviso?: string }> = {
  suspender: {
    titulo: 'Suspender linha',
    confirmar: 'Suspender',
    sucesso: 'Linha suspensa.',
    aviso: 'A linha deixa de contar no custo mensal até ser reativada.',
  },
  reativar: { titulo: 'Reativar linha', confirmar: 'Reativar', sucesso: 'Linha reativada.' },
  cancelar: {
    titulo: 'Cancelar linha',
    confirmar: 'Cancelar linha',
    sucesso: 'Linha cancelada.',
    aviso: 'A linha deixa de contar no custo mensal e o aparelho vinculado é liberado.',
  },
}

interface LinhaStatusDialogProps {
  linha: Linha
  acao: AcaoStatusLinha
  open: boolean
  onOpenChange: (open: boolean) => void
}

export function LinhaStatusDialog({ linha, acao, open, onOpenChange }: LinhaStatusDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md">
        <Corpo linha={linha} acao={acao} onClose={() => onOpenChange(false)} />
      </DialogContent>
    </Dialog>
  )
}

function Corpo({ linha, acao, onClose }: { linha: Linha; acao: AcaoStatusLinha; onClose: () => void }) {
  const queryClient = useQueryClient()
  const textos = TEXTOS[acao]
  const [observacao, setObservacao] = useState('')
  const [erro, setErro] = useState<string | null>(null)

  const mutation = useMutation({
    mutationFn: () => alterarStatusLinha(linha.id, STATUS_DA_ACAO[acao], observacao.trim() || undefined),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['linhas'] })
      toast(textos.sucesso)
      onClose()
    },
    onError: (err: unknown) => setErro(getErrorMessage(err, 'Não foi possível alterar o status da linha.')),
  })

  return (
    <form
      noValidate
      className="space-y-4"
      onSubmit={(e) => {
        e.preventDefault()
        setErro(null)
        mutation.mutate()
      }}
    >
      <DialogHeader>
        <DialogTitle>{textos.titulo}</DialogTitle>
        <DialogDescription>{formatTelefone(linha.numero)}</DialogDescription>
      </DialogHeader>
      {textos.aviso && <p className="text-sm text-foreground">{textos.aviso}</p>}
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="linha-status-observacao">Observação (opcional)</Label>
        <Input
          id="linha-status-observacao"
          value={observacao}
          onChange={(e) => setObservacao(e.target.value)}
          maxLength={255}
        />
      </div>
      {erro && (
        <p role="alert" className="text-body-sm text-destructive">
          {erro}
        </p>
      )}
      <DialogFooter className="gap-2 sm:gap-0">
        <Button type="button" variant="ghost" onClick={onClose}>
          Voltar
        </Button>
        <Button type="submit" variant={acao === 'cancelar' ? 'destructive' : 'default'} loading={mutation.isPending}>
          {textos.confirmar}
        </Button>
      </DialogFooter>
    </form>
  )
}
