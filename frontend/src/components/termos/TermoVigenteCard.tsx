import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { downloadAuthenticated } from '@/api/client'
import { downloadSignedTerm } from '@/api/loans'
import {
  STATUS_TERMO_PENDENTES,
  TIPO_TERMO_ROTULO,
  cancelarTermo,
  getArquivosTermo,
  reenviarTermo,
  type TermoSolicitacao,
} from '@/api/termos'
import { Button } from '@/components/ui/button'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { formatDateTime } from '@/lib/utils'
import { AnexarAssinadoDialog } from './AnexarAssinadoDialog'
import { SolicitarAssinaturaDialog } from './SolicitarAssinaturaDialog'
import { TermoStatusPill } from './TermoStatusPill'

const STATUS_ITEM_COM_TERMO = ['Pendente', 'Indisponível', 'Pendente Devolução']

interface TermoVigenteCardProps {
  itemId: number
  itemStatus?: string | null
  vigente: TermoSolicitacao | null
  revsignAtivo: boolean
}

function salvarBlob(blob: Blob, nome: string) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = nome
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function Campo({ rotulo, children }: { rotulo: string; children: React.ReactNode }) {
  return (
    <div className="min-w-0">
      <dt className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{rotulo}</dt>
      <dd className="mt-0.5 break-words text-sm font-medium text-foreground">{children}</dd>
    </div>
  )
}

/** Termo vigente do equipamento, com as ações de assinatura (solicitar, reenviar, cancelar, anexar, baixar). */
export function TermoVigenteCard({ itemId, itemStatus, vigente, revsignAtivo }: TermoVigenteCardProps) {
  const queryClient = useQueryClient()
  const [solicitarAberto, setSolicitarAberto] = useState(false)
  const [anexarAberto, setAnexarAberto] = useState(false)

  const podeSolicitar = !!itemStatus && STATUS_ITEM_COM_TERMO.includes(itemStatus)
  const real = vigente && vigente.id !== null ? vigente : null
  const pendente = !!real && STATUS_TERMO_PENDENTES.includes(real.status)

  function atualizar() {
    void queryClient.invalidateQueries({ queryKey: ['item-termos', itemId] })
    void queryClient.invalidateQueries({ queryKey: ['items'] })
  }

  const cancelar = useMutation({
    mutationFn: () => cancelarTermo(real?.id as number),
    onSuccess: () => {
      toast.success('Solicitação cancelada.')
      atualizar()
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  const reenviar = useMutation({
    mutationFn: () => reenviarTermo(real?.id as number),
    onSuccess: () => {
      toast.success('Termo reenviado.')
      atualizar()
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  const baixarTermo = useMutation({
    mutationFn: async () => {
      const arquivos = await getArquivosTermo(real?.id as number)
      if (!arquivos.documento_url) throw new Error('O arquivo do termo não está mais disponível.')
      await downloadAuthenticated(arquivos.documento_url, `termo_${itemId}.docx`)
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  const verAssinado = useMutation({
    mutationFn: async () => {
      if (vigente?.legado) {
        salvarBlob(await downloadSignedTerm(itemId), `termo_assinado_${itemId}.pdf`)
        return
      }
      const arquivos = await getArquivosTermo(real?.id as number)
      if (!arquivos.assinado_url) throw new Error('O PDF assinado não está mais disponível.')
      await downloadAuthenticated(arquivos.assinado_url, `termo_assinado_${itemId}.pdf`)
    },
    onError: (error) => toast.error(getErrorMessage(error)),
  })

  return (
    <section className="space-y-4 rounded-lg border border-border p-4">
      <h3 className="text-heading-sm text-foreground">Termo vigente</h3>

      {vigente ? (
        <dl className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-2 lg:grid-cols-3">
          <Campo rotulo="Tipo">{TIPO_TERMO_ROTULO[vigente.tipo]}</Campo>
          <Campo rotulo="Versão do modelo">{vigente.modelo_versao || '—'}</Campo>
          <Campo rotulo="Status">
            <TermoStatusPill status={vigente.status} />
          </Campo>
          <Campo rotulo="Solicitação">
            Solicitado por {vigente.solicitado_por} em {formatDateTime(vigente.solicitado_em)}
          </Campo>
          {vigente.enviado_em && <Campo rotulo="Enviado em">{formatDateTime(vigente.enviado_em)}</Campo>}
          {vigente.agendado_para && <Campo rotulo="Agendado para">{formatDateTime(vigente.agendado_para)}</Campo>}
          {vigente.assinado_em && <Campo rotulo="Assinado em">{formatDateTime(vigente.assinado_em)}</Campo>}
          {vigente.destinatario_email && <Campo rotulo="Destinatário">{vigente.destinatario_email}</Campo>}
          {vigente.motivo && <Campo rotulo="Motivo">{vigente.motivo}</Campo>}
        </dl>
      ) : (
        <p className="text-sm text-muted-foreground">Nenhum termo vigente para este equipamento.</p>
      )}

      <div className="flex flex-wrap items-center gap-2">
        <Button size="sm" onClick={() => setSolicitarAberto(true)} disabled={!podeSolicitar}>
          Solicitar assinatura
        </Button>
        {real && (real.status === 'enviado' || (real.status === 'manual' && revsignAtivo)) && (
          <Button size="sm" variant="outline" onClick={() => reenviar.mutate()} loading={reenviar.isPending}>
            Reenviar
          </Button>
        )}
        {pendente && (
          <Button size="sm" variant="outline" onClick={() => cancelar.mutate()} loading={cancelar.isPending}>
            Cancelar
          </Button>
        )}
        {real && ['enviado', 'manual', 'erro'].includes(real.status) && (
          <Button size="sm" variant="outline" onClick={() => setAnexarAberto(true)}>
            Anexar assinado
          </Button>
        )}
        {real?.tem_documento && (
          <Button size="sm" variant="outline" onClick={() => baixarTermo.mutate()} loading={baixarTermo.isPending}>
            Baixar termo
          </Button>
        )}
        {vigente?.tem_assinado && (
          <Button size="sm" variant="outline" onClick={() => verAssinado.mutate()} loading={verAssinado.isPending}>
            Ver assinado
          </Button>
        )}
      </div>
      {!podeSolicitar && (
        <p className="text-xs text-muted-foreground">Disponível só para equipamento com colaborador.</p>
      )}

      <SolicitarAssinaturaDialog
        itemId={itemId}
        open={solicitarAberto}
        onOpenChange={setSolicitarAberto}
        revsignAtivo={revsignAtivo}
        temPendente={pendente}
      />
      {real && <AnexarAssinadoDialog termo={real} open={anexarAberto} onOpenChange={setAnexarAberto} />}
    </section>
  )
}
