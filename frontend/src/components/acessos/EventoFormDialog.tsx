import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { createEvento, type EventoDetalhe, type TipoEvento } from '@/api/acessos'
import { listCargos, type Colaborador } from '@/api/colaboradores'
import { AvisoSemSenha } from '@/components/acessos/AvisoSemSenha'
import { ROTULO_TIPO_EVENTO } from '@/components/acessos/rotulos'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
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

interface EventoFormDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  colaboradorInicial?: Colaborador | null
  onCreated: (evento: EventoDetalhe) => void
}

export function EventoFormDialog({ open, onOpenChange, colaboradorInicial, onCreated }: EventoFormDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-lg max-h-[90vh] overflow-y-auto">
        <FormBody
          colaboradorInicial={colaboradorInicial ?? null}
          onClose={() => onOpenChange(false)}
          onCreated={onCreated}
        />
      </DialogContent>
    </Dialog>
  )
}

function CargoSelect({
  id,
  value,
  onChange,
  cargos,
  vazio,
}: {
  id: string
  value: string
  onChange: (v: string) => void
  cargos: string[]
  vazio: string
}) {
  const opcoes = value && !cargos.includes(value) ? [value, ...cargos] : cargos
  return (
    <select id={id} className={selectClass} value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="">{vazio}</option>
      {opcoes.map((c) => (
        <option key={c} value={c}>
          {c}
        </option>
      ))}
    </select>
  )
}

function FormBody({
  colaboradorInicial,
  onClose,
  onCreated,
}: {
  colaboradorInicial: Colaborador | null
  onClose: () => void
  onCreated: (evento: EventoDetalhe) => void
}) {
  const queryClient = useQueryClient()
  const [tipo, setTipo] = useState<TipoEvento>('admissao')
  const [colaborador, setColaborador] = useState<Colaborador | null>(colaboradorInicial)
  const [cargoAnterior, setCargoAnterior] = useState('')
  const [cargoNovo, setCargoNovo] = useState(colaboradorInicial?.cargo ?? '')
  const [observacao, setObservacao] = useState('')

  const { data: cargos = [] } = useQuery({
    queryKey: ['colaboradores', 'cargos'],
    queryFn: listCargos,
  })

  const mutation = useMutation({
    mutationFn: () =>
      createEvento({
        colaborador_id: (colaborador as Colaborador).id,
        tipo,
        cargo_anterior: tipo === 'mudanca_funcao' ? cargoAnterior || null : null,
        cargo_novo: tipo === 'desligamento' ? null : cargoNovo || null,
        observacao: observacao.trim() || null,
      }),
    onSuccess: (evento) => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast('Evento registrado.')
      onClose()
      onCreated(evento)
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao registrar evento.'), 'error'),
  })

  function escolherColaborador(c: Colaborador | null) {
    setColaborador(c)
    setCargoNovo(c?.cargo ?? '')
  }

  const incompleto = !colaborador || (tipo === 'mudanca_funcao' && !cargoAnterior)

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (incompleto) return
    mutation.mutate()
  }

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>Registrar evento de acesso</DialogTitle>
        <DialogDescription>
          O evento gera a lista de tarefas para criar ou desativar os acessos do colaborador.
        </DialogDescription>
      </DialogHeader>

      <div className="flex flex-col gap-1.5">
        <Label htmlFor="evento-tipo">Tipo de evento</Label>
        <select
          id="evento-tipo"
          className={selectClass}
          value={tipo}
          onChange={(e) => setTipo(e.target.value as TipoEvento)}
        >
          {(Object.keys(ROTULO_TIPO_EVENTO) as TipoEvento[]).map((t) => (
            <option key={t} value={t}>
              {ROTULO_TIPO_EVENTO[t]}
            </option>
          ))}
        </select>
      </div>

      <div className="flex flex-col gap-1.5">
        <Label>Colaborador</Label>
        <ColaboradorSelect value={colaborador?.id ?? null} onChange={escolherColaborador} somenteAtivos={false} />
      </div>

      {tipo === 'admissao' && (
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="evento-cargo">Cargo</Label>
          <CargoSelect
            id="evento-cargo"
            value={cargoNovo}
            onChange={setCargoNovo}
            cargos={cargos}
            vazio="Cargo atual do colaborador"
          />
        </div>
      )}

      {tipo === 'mudanca_funcao' && (
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="evento-cargo-anterior">Cargo anterior *</Label>
            <CargoSelect
              id="evento-cargo-anterior"
              value={cargoAnterior}
              onChange={setCargoAnterior}
              cargos={cargos}
              vazio="Selecione..."
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="evento-cargo-novo">Cargo novo</Label>
            <CargoSelect
              id="evento-cargo-novo"
              value={cargoNovo}
              onChange={setCargoNovo}
              cargos={cargos}
              vazio="Cargo atual do colaborador"
            />
          </div>
        </div>
      )}

      {tipo === 'desligamento' && (
        <p className="text-sm text-muted-foreground">
          Serão geradas tarefas para desativar todos os acessos ativos do colaborador, e o evento só termina
          quando ele não tiver mais equipamentos.
        </p>
      )}

      <div className="flex flex-col gap-1.5">
        <Label htmlFor="evento-observacao">Observação</Label>
        <Input
          id="evento-observacao"
          value={observacao}
          onChange={(e) => setObservacao(e.target.value)}
          maxLength={500}
        />
      </div>

      <AvisoSemSenha />

      <DialogFooter>
        <Button type="button" variant="outline" onClick={onClose}>
          Cancelar
        </Button>
        <Button type="submit" loading={mutation.isPending} disabled={incompleto}>
          Registrar evento
        </Button>
      </DialogFooter>
    </form>
  )
}
