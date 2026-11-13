import { useState } from 'react'
import { Link, useNavigate } from 'react-router-dom'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Plus } from 'lucide-react'
import { desativarAcesso, listAcessos, listEventos, type Acesso } from '@/api/acessos'
import type { Colaborador } from '@/api/colaboradores'
import { AcessoFormDialog } from '@/components/acessos/AcessoFormDialog'
import { EventoFormDialog } from '@/components/acessos/EventoFormDialog'
import { ROTULO_TIPO_EVENTO, STATUS_ACESSO, STATUS_EVENTO } from '@/components/acessos/rotulos'
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
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { formatDateTime } from '@/lib/utils'
import type { ColaboradorDetailTab } from '../types'

const linkClass =
  'rounded-sm font-semibold text-brand-text hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2'

function AcessosTab({ colaborador }: { colaborador: Colaborador }) {
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const [formAberto, setFormAberto] = useState(false)
  const [editando, setEditando] = useState<Acesso | null>(null)
  const [desativando, setDesativando] = useState<Acesso | null>(null)
  const [eventoAberto, setEventoAberto] = useState(false)

  const ficha = useQuery({
    queryKey: ['acessos', 'ficha', { colaborador_id: colaborador.id }],
    queryFn: () => listAcessos({ colaborador_id: colaborador.id }),
  })
  const eventos = useQuery({
    queryKey: ['acessos', 'eventos', { colaborador_id: colaborador.id, limit: 20, offset: 0 }],
    queryFn: () => listEventos({ colaborador_id: colaborador.id, limit: 20, offset: 0 }),
  })

  const desativar = useMutation({
    mutationFn: (id: number) => desativarAcesso(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast('Acesso desativado.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao desativar o acesso.'), 'error'),
  })

  function abrirForm(acesso: Acesso | null) {
    setEditando(acesso)
    setFormAberto(true)
  }

  return (
    <div className="space-y-8">
      <section aria-labelledby="titulo-ficha" className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h2 id="titulo-ficha" className="text-base font-semibold">
            Ficha de acessos
          </h2>
          <Button size="sm" onClick={() => abrirForm(null)}>
            <Plus aria-hidden="true" />
            Adicionar acesso
          </Button>
        </div>

        {ficha.isLoading ? (
          <p className="text-sm text-muted-foreground">Carregando acessos...</p>
        ) : ficha.isError ? (
          <p className="text-sm text-muted-foreground">Não foi possível carregar os acessos.</p>
        ) : (ficha.data ?? []).length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum acesso registrado para este colaborador.</p>
        ) : (
          <ul className="divide-y divide-border rounded-lg border border-border">
            {(ficha.data ?? []).map((a) => {
              const st = STATUS_ACESSO[a.status]
              return (
                <li key={a.id} className="flex flex-wrap items-start justify-between gap-3 px-4 py-3">
                  <div className="min-w-0 space-y-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-sm font-semibold">{a.sistema_nome}</span>
                      <span className="text-sm">{a.login}</span>
                      <StatusPill status={st.label} tone={st.tone} />
                    </div>
                    <p className="text-xs text-muted-foreground">
                      {'Equipamento: '}
                      {a.item_id !== null ? (
                        <Link to={`/items/${a.item_id}`} className={linkClass}>
                          {a.item_descricao ?? `Equipamento #${a.item_id}`}
                        </Link>
                      ) : (
                        '—'
                      )}
                    </p>
                    <p className="text-xs text-muted-foreground">
                      {`Criado por ${a.criado_por} em ${formatDateTime(a.criado_em)}`}
                    </p>
                    <p className="text-xs text-muted-foreground">
                      {a.desativado_por
                        ? `Desativado por ${a.desativado_por} em ${formatDateTime(a.desativado_em)}`
                        : 'Desativado: —'}
                    </p>
                  </div>
                  <div className="flex gap-2">
                    <Button
                      size="sm"
                      variant="outline"
                      aria-label={`Editar acesso ${a.sistema_nome} ${a.login}`}
                      onClick={() => abrirForm(a)}
                    >
                      Editar
                    </Button>
                    {a.status === 'ativo' && (
                      <Button
                        size="sm"
                        variant="outline"
                        aria-label={`Desativar acesso ${a.sistema_nome} ${a.login}`}
                        className="text-destructive"
                        loading={desativar.isPending && desativar.variables === a.id}
                        onClick={() => setDesativando(a)}
                      >
                        Desativar
                      </Button>
                    )}
                  </div>
                </li>
              )
            })}
          </ul>
        )}
      </section>

      <section aria-labelledby="titulo-eventos" className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h2 id="titulo-eventos" className="text-base font-semibold">
            Eventos de acesso
          </h2>
          <Button size="sm" variant="outline" onClick={() => setEventoAberto(true)}>
            Registrar evento
          </Button>
        </div>
        {eventos.isLoading ? (
          <p className="text-sm text-muted-foreground">Carregando eventos...</p>
        ) : eventos.isError ? (
          <p className="text-sm text-muted-foreground">Não foi possível carregar os eventos.</p>
        ) : (eventos.data?.items ?? []).length === 0 ? (
          <p className="text-sm text-muted-foreground">Nenhum evento de acesso para este colaborador.</p>
        ) : (
          <ul className="divide-y divide-border rounded-lg border border-border">
            {(eventos.data?.items ?? []).map((e) => {
              const st = STATUS_EVENTO[e.status]
              return (
                <li key={e.id} className="flex flex-wrap items-center justify-between gap-2 px-4 py-3">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-sm font-semibold">{ROTULO_TIPO_EVENTO[e.tipo]}</span>
                    <StatusPill status={st.label} tone={st.tone} />
                    <span className="text-xs text-muted-foreground">{formatDateTime(e.aberto_em)}</span>
                  </div>
                  <Link to={`/acessos/eventos/${e.id}`} aria-label={`Abrir evento #${e.id}`} className={`text-sm ${linkClass}`}>
                    Abrir
                  </Link>
                </li>
              )
            })}
          </ul>
        )}
      </section>

      <AcessoFormDialog
        open={formAberto}
        onOpenChange={setFormAberto}
        colaboradorId={colaborador.id}
        acesso={editando}
      />
      <EventoFormDialog
        open={eventoAberto}
        onOpenChange={setEventoAberto}
        colaboradorInicial={colaborador}
        onCreated={(evento) => navigate(`/acessos/eventos/${evento.id}`)}
      />
      <AlertDialog open={desativando !== null} onOpenChange={(o) => !o && setDesativando(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Desativar acesso</AlertDialogTitle>
            <AlertDialogDescription>
              {desativando &&
                `Desativar o acesso "${desativando.login}" no sistema ${desativando.sistema_nome}?`}
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancelar</AlertDialogCancel>
            <AlertDialogAction
              onClick={() => {
                if (desativando) desativar.mutate(desativando.id)
              }}
            >
              Desativar
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  )
}

const tab: ColaboradorDetailTab = {
  id: 'acessos',
  label: 'Acessos',
  order: 50,
  roles: ['Gestor', 'Técnico'],
  Component: AcessosTab,
}
export default tab
