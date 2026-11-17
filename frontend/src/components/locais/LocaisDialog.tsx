import type { Unidade } from '@/api/unidades'
import { LocaisManager } from '@/components/locais/LocaisManager'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { upper } from '@/lib/format'

interface LocaisDialogProps {
  unidade: Pick<Unidade, 'id' | 'nome'> | null
  open: boolean
  onOpenChange: (open: boolean) => void
}

/** Lista, cria, renomeia e ativa/desativa os locais de uma unidade. */
export function LocaisDialog({ unidade, open, onOpenChange }: LocaisDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        {unidade && (
          <div className="space-y-4">
            <DialogHeader>
              <DialogTitle>Locais — {upper(unidade.nome)}</DialogTitle>
              <DialogDescription>
                Locais físicos desta unidade (salas, almoxarifados...). Locais desativados deixam de ser
                oferecidos nos cadastros, mas continuam nos itens já vinculados.
              </DialogDescription>
            </DialogHeader>
            <LocaisManager unidade={unidade} />
          </div>
        )}
      </DialogContent>
    </Dialog>
  )
}
