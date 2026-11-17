import { useQuery } from '@tanstack/react-query'
import { listEventosLinha, type Linha } from '@/api/linhas'
import { LinhaEventosList } from '@/components/linhas/LinhaEventosList'
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from '@/components/ui/dialog'
import { formatTelefone } from '@/lib/format'

interface LinhaHistoricoDialogProps {
  linha: Linha | null
  open: boolean
  onOpenChange: (open: boolean) => void
}

export function LinhaHistoricoDialog({ linha, open, onOpenChange }: LinhaHistoricoDialogProps) {
  const { data, isLoading, isError } = useQuery({
    queryKey: ['linhas', 'eventos', linha?.id],
    queryFn: () => listEventosLinha(linha!.id),
    enabled: open && linha !== null,
  })

  return (
    <Dialog open={open && linha !== null} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[90vh] max-w-xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Histórico da linha {formatTelefone(linha?.numero)}</DialogTitle>
          <DialogDescription>Trocas, mudanças de status e edições, da mais recente para a mais antiga.</DialogDescription>
        </DialogHeader>
        {isLoading ? (
          <p className="text-sm text-muted-foreground">Carregando histórico...</p>
        ) : isError ? (
          <p className="text-sm text-muted-foreground">Não foi possível carregar o histórico.</p>
        ) : (
          <LinhaEventosList eventos={data ?? []} />
        )}
      </DialogContent>
    </Dialog>
  )
}
