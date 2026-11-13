import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import {
  baixarAnexoManutencao,
  cancelarManutencao,
  enviarManutencao,
  marcarEmReparo,
  registrarRetorno,
  type Manutencao,
} from '@/api/manutencoes'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { FileUpload } from '@/components/ui/FileUpload'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { anexoGrandeDemais, isDataBR, maskDataBR, maskMoedaInput, parseMoedaBR } from '@/features/ciclo-vida/format'
import { getErrorMessage } from '@/lib/api-error'
import { STATUS_EM_ANDAMENTO } from './status'

type Dialogo = 'envio' | 'retorno' | 'cancelar' | null

const ACCEPT = { 'application/pdf': ['.pdf'], 'image/png': ['.png'], 'image/jpeg': ['.jpg', '.jpeg'] }

/** Chaves de consulta afetadas por qualquer mudança numa manutenção. */
const CHAVES_AFETADAS = [['manutencoes'], ['item'], ['items'], ['item-linha-do-tempo']]

/** Botões (e diálogos) das ações possíveis para o status atual da manutenção. */
export function ManutencaoAcoes({ manutencao }: { manutencao: Manutencao }) {
  const queryClient = useQueryClient()
  const [dialogo, setDialogo] = useState<Dialogo>(null)

  // Envio
  const [assistencia, setAssistencia] = useState(manutencao.assistencia ?? '')
  const [previsao, setPrevisao] = useState('')
  // Retorno
  const [resultado, setResultado] = useState<'retornada' | 'sem_conserto'>('retornada')
  const [custo, setCusto] = useState('')
  const [observacoes, setObservacoes] = useState('')
  const [anexo, setAnexo] = useState<File | null>(null)
  const [anexoKey, setAnexoKey] = useState(0)
  // Cancelamento
  const [motivo, setMotivo] = useState('')

  const emAndamento = STATUS_EM_ANDAMENTO.includes(manutencao.status)

  function sucesso(detail: string) {
    toast(detail)
    for (const queryKey of CHAVES_AFETADAS) void queryClient.invalidateQueries({ queryKey })
    setDialogo(null)
  }
  const falha = (fallback: string) => (err: unknown) => toast(getErrorMessage(err, fallback), 'error')

  const envio = useMutation({
    mutationFn: () =>
      enviarManutencao(manutencao.id, {
        assistencia: assistencia.trim(),
        ...(previsao.trim() ? { previsao_retorno: previsao.trim() } : {}),
      }),
    onSuccess: (r) => sucesso(r.detail),
    onError: falha('Não foi possível registrar o envio.'),
  })

  const emReparo = useMutation({
    mutationFn: () => marcarEmReparo(manutencao.id, {}),
    onSuccess: (r) => sucesso(r.detail),
    onError: falha('Não foi possível marcar em reparo.'),
  })

  const retorno = useMutation({
    mutationFn: (valor: number | null) =>
      registrarRetorno(manutencao.id, {
        resultado,
        custo: valor,
        observacoes: observacoes.trim() || undefined,
        anexo,
      }),
    onSuccess: (r) => sucesso(r.detail),
    onError: falha('Não foi possível registrar o retorno.'),
  })

  const cancelamento = useMutation({
    mutationFn: () => cancelarManutencao(manutencao.id, { motivo: motivo.trim() }),
    onSuccess: (r) => sucesso(r.detail),
    onError: falha('Não foi possível cancelar a manutenção.'),
  })

  function confirmarEnvio() {
    if (previsao.trim() && !isDataBR(previsao)) {
      toast('Previsão de retorno inválida. Use dd/mm/aaaa.', 'error')
      return
    }
    envio.mutate()
  }

  function confirmarRetorno() {
    const valor = custo.trim() ? parseMoedaBR(custo) : null
    if (custo.trim() && valor === null) {
      toast('Custo inválido. Use números com até duas casas decimais.', 'error')
      return
    }
    retorno.mutate(valor)
  }

  function selecionarAnexo(file: File | null) {
    if (file && anexoGrandeDemais(file)) {
      toast('O anexo deve ter no máximo 10 MB.', 'error')
      setAnexo(null)
      setAnexoKey((k) => k + 1)
      return
    }
    setAnexo(file)
  }

  return (
    <div className="flex flex-wrap gap-2">
      {manutencao.status === 'aberta' && (
        <Button type="button" variant="outline" size="sm" onClick={() => setDialogo('envio')}>
          Registrar envio
        </Button>
      )}
      {(manutencao.status === 'aberta' || manutencao.status === 'enviada') && (
        <Button type="button" variant="outline" size="sm" loading={emReparo.isPending} onClick={() => emReparo.mutate()}>
          Marcar em reparo
        </Button>
      )}
      {emAndamento && (
        <>
          <Button type="button" variant="outline" size="sm" onClick={() => setDialogo('retorno')}>
            Registrar retorno
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={() => setDialogo('cancelar')}>
            Cancelar manutenção
          </Button>
        </>
      )}
      {manutencao.anexo_key && (
        <Button
          type="button"
          variant="ghost"
          size="sm"
          onClick={() =>
            void baixarAnexoManutencao(manutencao).catch(falha('Não foi possível baixar o anexo.'))
          }
        >
          Baixar anexo
        </Button>
      )}

      <Dialog open={dialogo === 'envio'} onOpenChange={(aberto) => !aberto && setDialogo(null)}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Registrar envio à assistência</DialogTitle>
            <DialogDescription>{manutencao.defeito}</DialogDescription>
          </DialogHeader>
          <form
            className="space-y-4"
            onSubmit={(e) => {
              e.preventDefault()
              if (assistencia.trim()) confirmarEnvio()
            }}
          >
            <div className="flex flex-col gap-1.5">
              <Label htmlFor={`envio-assistencia-${manutencao.id}`} className="text-caption font-medium">
                Assistência técnica
              </Label>
              <Input
                id={`envio-assistencia-${manutencao.id}`}
                value={assistencia}
                maxLength={150}
                onChange={(e) => setAssistencia(e.target.value)}
                className="h-9"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor={`envio-previsao-${manutencao.id}`} className="text-caption font-medium">
                Previsão de retorno
              </Label>
              <Input
                id={`envio-previsao-${manutencao.id}`}
                value={previsao}
                onChange={(e) => setPrevisao(maskDataBR(e.target.value))}
                placeholder="DD/MM/AAAA (opcional)"
                inputMode="numeric"
                className="h-9 font-mono"
              />
            </div>
            <DialogFooter className="gap-2 sm:gap-0">
              <Button type="button" variant="outline" size="sm" onClick={() => setDialogo(null)}>
                Voltar
              </Button>
              <Button type="submit" size="sm" loading={envio.isPending} disabled={!assistencia.trim()}>
                Confirmar envio
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      <Dialog open={dialogo === 'retorno'} onOpenChange={(aberto) => !aberto && setDialogo(null)}>
        <DialogContent className="max-h-[90vh] max-w-lg overflow-y-auto">
          <DialogHeader>
            <DialogTitle>Registrar retorno</DialogTitle>
            <DialogDescription>{manutencao.defeito}</DialogDescription>
          </DialogHeader>
          <form
            className="space-y-4"
            onSubmit={(e) => {
              e.preventDefault()
              confirmarRetorno()
            }}
          >
            <div className="flex flex-col gap-2">
              <label className="flex items-center gap-2 text-body-sm text-foreground">
                <input
                  type="radio"
                  name={`resultado-${manutencao.id}`}
                  className="size-4"
                  checked={resultado === 'retornada'}
                  onChange={() => setResultado('retornada')}
                />
                Consertado — volta para Disponível
              </label>
              <label className="flex items-center gap-2 text-body-sm text-foreground">
                <input
                  type="radio"
                  name={`resultado-${manutencao.id}`}
                  className="size-4"
                  checked={resultado === 'sem_conserto'}
                  onChange={() => setResultado('sem_conserto')}
                />
                Sem conserto
              </label>
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor={`retorno-custo-${manutencao.id}`} className="text-caption font-medium">
                Custo (R$)
              </Label>
              <Input
                id={`retorno-custo-${manutencao.id}`}
                value={custo}
                onChange={(e) => setCusto(maskMoedaInput(e.target.value))}
                placeholder="0,00"
                inputMode="decimal"
                className="h-9 font-mono"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <Label htmlFor={`retorno-obs-${manutencao.id}`} className="text-caption font-medium">
                Observações
              </Label>
              <Input
                id={`retorno-obs-${manutencao.id}`}
                value={observacoes}
                maxLength={500}
                onChange={(e) => setObservacoes(e.target.value)}
                className="h-9"
              />
            </div>
            <FileUpload
              key={anexoKey}
              accept={ACCEPT}
              onFile={selecionarAnexo}
              label="Anexe a ordem de serviço ou nota (PDF ou imagem, até 10 MB)"
            />
            <DialogFooter className="gap-2 sm:gap-0">
              <Button type="button" variant="outline" size="sm" onClick={() => setDialogo(null)}>
                Voltar
              </Button>
              <Button type="submit" size="sm" loading={retorno.isPending}>
                Confirmar retorno
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      <Dialog open={dialogo === 'cancelar'} onOpenChange={(aberto) => !aberto && setDialogo(null)}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Cancelar manutenção</DialogTitle>
            <DialogDescription>O equipamento volta a Disponível.</DialogDescription>
          </DialogHeader>
          <form
            className="space-y-4"
            onSubmit={(e) => {
              e.preventDefault()
              if (motivo.trim()) cancelamento.mutate()
            }}
          >
            <div className="flex flex-col gap-1.5">
              <Label htmlFor={`cancelar-motivo-${manutencao.id}`} className="text-caption font-medium">
                Motivo do cancelamento
              </Label>
              <Input
                id={`cancelar-motivo-${manutencao.id}`}
                value={motivo}
                maxLength={255}
                onChange={(e) => setMotivo(e.target.value)}
                className="h-9"
              />
            </div>
            <DialogFooter className="gap-2 sm:gap-0">
              <Button type="button" variant="outline" size="sm" onClick={() => setDialogo(null)}>
                Voltar
              </Button>
              <Button type="submit" size="sm" loading={cancelamento.isPending} disabled={!motivo.trim()}>
                Confirmar cancelamento
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}
