import { useState } from 'react'
import type { Leitura } from '@/api/inventarios'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog'
import { Button } from '@/components/ui/button'
import { StatusPill } from '@/components/ui/StatusPill'
import { ROTULO_RESULTADO, TOM_RESULTADO } from '@/features/inventario/resultados'
import { formatDateTime } from '@/lib/utils'

interface UltimaLeituraCardProps {
  leitura: Leitura
  onInformarUsuario: (colaboradorId: number) => void
  onSemUsuario: () => void
  onDesfazer: () => void
  /** Uma correção ou exclusão está em andamento. */
  ocupado?: boolean
}

/** Resultado da última leitura, com as correções que o técnico pode fazer na hora. */
export function UltimaLeituraCard({
  leitura,
  onInformarUsuario,
  onSemUsuario,
  onDesfazer,
  ocupado = false,
}: UltimaLeituraCardProps) {
  const [escolhendo, setEscolhendo] = useState(false)
  const [confirmando, setConfirmando] = useState(false)
  const temItem = leitura.item_id != null

  const titulo = temItem
    ? `#${leitura.item_id} · ${leitura.item_tipo ?? ''} · ${[leitura.item_brand, leitura.item_model]
        .filter(Boolean)
        .join(' ')}`.replace(/( · )+$/, '')
    : leitura.codigo_lido

  return (
    <section aria-label="Última leitura" className="surface-panel space-y-3 p-4">
      <div className="flex flex-wrap items-center gap-2">
        <StatusPill status={ROTULO_RESULTADO[leitura.resultado]} tone={TOM_RESULTADO[leitura.resultado]} />
        <span className="text-body-sm font-medium">{titulo}</span>
      </div>

      {leitura.detalhe && <p className="text-body-sm text-foreground">{leitura.detalhe}</p>}
      {temItem && (
        <p className="text-body-sm text-muted-foreground">
          {`Usuário cadastrado: ${leitura.usuario_esperado ?? 'nenhum'}`}
        </p>
      )}
      {leitura.duplicada && (
        <p className="text-body-sm text-muted-foreground">
          {`Já lido em ${formatDateTime(leitura.lido_em)} por ${leitura.lido_por} — não contou de novo.`}
        </p>
      )}

      {escolhendo && (
        <div className="space-y-2">
          <ColaboradorSelect
            value={leitura.colaborador_id_encontrado ?? null}
            placeholder="Quem está usando?"
            onChange={(c) => {
              if (!c) return
              setEscolhendo(false)
              onInformarUsuario(c.id)
            }}
          />
        </div>
      )}

      <div className="flex flex-wrap gap-2">
        {temItem && (
          <>
            <Button variant="outline" size="lg" disabled={ocupado} onClick={() => setEscolhendo((v) => !v)}>
              Informar usuário
            </Button>
            <Button variant="outline" size="lg" disabled={ocupado} onClick={onSemUsuario}>
              Ninguém usando
            </Button>
          </>
        )}
        <Button variant="outline" size="lg" disabled={ocupado} onClick={() => setConfirmando(true)}>
          Desfazer leitura
        </Button>
      </div>

      <AlertDialog open={confirmando} onOpenChange={setConfirmando}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Desfazer esta leitura?</AlertDialogTitle>
            <AlertDialogDescription>
              A leitura sai do inventário e o equipamento volta a contar como ainda não lido. Você pode lê-lo de novo.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Manter</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                setConfirmando(false)
                onDesfazer()
              }}
            >
              Desfazer
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </section>
  )
}
