import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { anexarTermoAssinado, type TermoSolicitacao } from '@/api/termos'
import { Button } from '@/components/ui/button'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'

interface AnexarAssinadoDialogProps {
  termo: TermoSolicitacao
  open: boolean
  onOpenChange: (open: boolean) => void
}

/** Anexa à mão o PDF assinado de uma solicitação (o mesmo efeito da assinatura pelo RevSign). */
export function AnexarAssinadoDialog({ termo, open, onOpenChange }: AnexarAssinadoDialogProps) {
  const queryClient = useQueryClient()
  const [file, setFile] = useState<File | null>(null)

  const mutation = useMutation({
    mutationFn: () => anexarTermoAssinado(termo.id as number, file as File),
    onSuccess: () => {
      toast.success('Termo assinado anexado.')
      void queryClient.invalidateQueries({ queryKey: ['item-termos', termo.item_id] })
      void queryClient.invalidateQueries({ queryKey: ['item', termo.item_id] })
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setFile(null)
      onOpenChange(false)
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Anexar termo assinado</DialogTitle>
          <DialogDescription>
            O PDF assinado conclui a solicitação, como se o colaborador tivesse assinado pelo RevSign.
          </DialogDescription>
        </DialogHeader>
        <FileUpload onFile={setFile} />
        <DialogFooter className="gap-2">
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            Cancelar
          </Button>
          <Button onClick={() => mutation.mutate()} disabled={!file} loading={mutation.isPending}>
            Anexar
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}
