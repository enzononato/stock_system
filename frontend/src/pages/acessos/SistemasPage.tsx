import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Ban, Pencil, Plus, RotateCcw } from 'lucide-react'
import { listSistemas, updateSistema, type Sistema } from '@/api/acessos'
import { SistemaFormDialog } from '@/components/acessos/SistemaFormDialog'
import { STATUS_ACESSO } from '@/components/acessos/rotulos'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'

/** Catálogo de sistemas (só Gestor). Sem paginação: o catálogo é curto e limitado. */
export default function SistemasPage() {
  const queryClient = useQueryClient()
  const [formAberto, setFormAberto] = useState(false)
  const [editando, setEditando] = useState<Sistema | null>(null)

  const { data: sistemas = [], isLoading, isError, error, refetch } = useQuery({
    queryKey: ['acessos', 'sistemas', { include_inactive: true }],
    queryFn: () => listSistemas({ include_inactive: true }),
  })

  const alterarSituacao = useMutation({
    mutationFn: ({ id, ativo }: { id: number; ativo: boolean }) => updateSistema(id, { ativo }),
    onSuccess: (_, { ativo }) => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast(ativo ? 'Sistema reativado.' : 'Sistema desativado.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao atualizar sistema.'), 'error'),
  })
  const idEmAndamento = alterarSituacao.isPending ? alterarSituacao.variables?.id : undefined

  function abrir(sistema: Sistema | null) {
    setEditando(sistema)
    setFormAberto(true)
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Acessos"
        title="Sistemas"
        description="Catálogo de sistemas usados nos pacotes por função e na ficha de acessos."
        actions={
          <Button onClick={() => abrir(null)}>
            <Plus aria-hidden="true" />
            Novo sistema
          </Button>
        }
      />

      {isLoading ? (
        <LoadingState label="Carregando sistemas..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : sistemas.length === 0 ? (
        <EmptyState title="Nenhum sistema cadastrado" description="Cadastre o primeiro sistema do catálogo." />
      ) : (
        <div className="surface-panel overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="border-b border-border bg-surface-alt">
                <tr>
                  {['Nome', 'Descrição', 'Situação', ''].map((h, i) => (
                    <th key={h || i} className="whitespace-nowrap px-4 py-2.5 text-left text-caption text-muted-foreground">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sistemas.map((s) => {
                  const situacao = STATUS_ACESSO[s.ativo ? 'ativo' : 'desativado']
                  return (
                    <tr key={s.id} className="border-b border-border last:border-0 hover:bg-surface-alt">
                      <td className="px-4 py-2.5 text-body-sm font-medium">{s.nome}</td>
                      <td className="px-4 py-2.5 text-body-sm text-muted-foreground">{s.descricao || '—'}</td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <StatusPill status={situacao.label} tone={situacao.tone} />
                      </td>
                      <td className="whitespace-nowrap px-4 py-2.5">
                        <div className="flex gap-1">
                          <Button size="sm" variant="ghost" aria-label={`Editar ${s.nome}`} onClick={() => abrir(s)}>
                            <Pencil size={14} />
                            Editar
                          </Button>
                          {s.ativo ? (
                            <Button
                              size="sm"
                              variant="ghost"
                              aria-label={`Desativar ${s.nome}`}
                              className="text-muted-foreground hover:text-destructive"
                              loading={idEmAndamento === s.id}
                              onClick={() => alterarSituacao.mutate({ id: s.id, ativo: false })}
                            >
                              <Ban size={14} />
                              Desativar
                            </Button>
                          ) : (
                            <Button
                              size="sm"
                              variant="ghost"
                              aria-label={`Reativar ${s.nome}`}
                              loading={idEmAndamento === s.id}
                              onClick={() => alterarSituacao.mutate({ id: s.id, ativo: true })}
                            >
                              <RotateCcw size={14} />
                              Reativar
                            </Button>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <SistemaFormDialog open={formAberto} onOpenChange={setFormAberto} sistema={editando} />
    </div>
  )
}
