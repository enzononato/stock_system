import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { solicitarTermo } from '@/api/termos'
import { Button } from '@/components/ui/button'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { formatDateTime } from '@/lib/utils'

interface SolicitarAssinaturaDialogProps {
  itemId: number
  open: boolean
  onOpenChange: (open: boolean) => void
  revsignAtivo: boolean
  /** Há uma solicitação pendente que a nova vai cancelar. */
  temPendente: boolean
}

/** Pede a assinatura do termo do equipamento agora ou em data agendada (só com o RevSign ligado). */
export function SolicitarAssinaturaDialog({
  itemId,
  open,
  onOpenChange,
  revsignAtivo,
  temPendente,
}: SolicitarAssinaturaDialogProps) {
  const queryClient = useQueryClient()
  const [modo, setModo] = useState<'agora' | 'agendar'>('agora')
  const [quando, setQuando] = useState('')

  const mutation = useMutation({
    mutationFn: () => solicitarTermo(itemId, modo === 'agendar' ? { agendado_para: quando } : {}),
    onSuccess: (termo) => {
      if (termo.status === 'agendado') {
        toast.success(`Envio agendado para ${formatDateTime(termo.agendado_para)}.`)
      } else if (termo.status === 'enviado') {
        toast.success(`Termo enviado para assinatura em ${termo.destinatario_email}.`)
      } else {
        toast.success(`Solicitação registrada para assinatura manual: ${termo.motivo ?? ''}`)
      }
      void queryClient.invalidateQueries({ queryKey: ['item-termos', itemId] })
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      setModo('agora')
      setQuando('')
      onOpenChange(false)
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  const agendarSemData = modo === 'agendar' && !quando

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Solicitar assinatura do termo</DialogTitle>
          <DialogDescription>O colaborador recebe o termo do equipamento para assinar.</DialogDescription>
        </DialogHeader>

        <div className="space-y-3">
          <label className="flex items-center gap-2 text-sm text-foreground">
            <input type="radio" name="modo-envio" checked={modo === 'agora'} onChange={() => setModo('agora')} />
            Enviar agora
          </label>
          <label className="flex items-center gap-2 text-sm text-foreground">
            <input
              type="radio"
              name="modo-envio"
              checked={modo === 'agendar'}
              disabled={!revsignAtivo}
              onChange={() => setModo('agendar')}
            />
            Agendar envio
          </label>
          {!revsignAtivo && (
            <p className="text-xs text-muted-foreground">Agendamento exige a integração com o RevSign ligada.</p>
          )}
          {modo === 'agendar' && (
            <Input
              type="datetime-local"
              aria-label="Data e hora do envio"
              value={quando}
              onChange={(e) => setQuando(e.target.value)}
            />
          )}
          {temPendente && (
            <p className="text-xs text-muted-foreground">A solicitação pendente deste equipamento será cancelada.</p>
          )}
        </div>

        <DialogFooter className="gap-2">
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            Fechar
          </Button>
          <Button onClick={() => mutation.mutate()} disabled={agendarSemData} loading={mutation.isPending}>
            Solicitar
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}
