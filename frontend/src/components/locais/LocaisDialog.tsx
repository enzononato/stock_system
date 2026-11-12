import { useEffect, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { Ban, Check, Pencil, Plus, RotateCcw, X } from 'lucide-react'
import { createLocal, listLocais, updateLocal, type Local, type LocalUpdateInput } from '@/api/locais'
import type { Unidade } from '@/api/unidades'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Pagination } from '@/components/ui/Pagination'
import { EmptyState, ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { usePageSize } from '@/hooks/usePageSize'
import { getErrorMessage } from '@/lib/api-error'
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
        {unidade && <Corpo unidade={unidade} />}
      </DialogContent>
    </Dialog>
  )
}

function Corpo({ unidade }: { unidade: Pick<Unidade, 'id' | 'nome'> }) {
  const queryClient = useQueryClient()
  const [pageSize, setPageSize] = usePageSize('locais')
  const [page, setPage] = useState(1)
  const [nome, setNome] = useState('')
  const [descricao, setDescricao] = useState('')
  const [editandoId, setEditandoId] = useState<number | null>(null)
  const [nomeEditado, setNomeEditado] = useState('')

  const { data: locais = [], isLoading, isError, error, refetch } = useQuery({
    queryKey: ['locais', { unidadeId: unidade.id, includeInactive: true }],
    queryFn: () => listLocais({ unidade_id: unidade.id, include_inactive: true }),
  })

  const lastPage = Math.max(1, Math.ceil(locais.length / pageSize))
  // Se a lista encolheu ou o tamanho de página cresceu, volta para a última página existente.
  useEffect(() => {
    if (page > lastPage) setPage(lastPage)
  }, [page, lastPage])
  const visiveis = locais.slice((page - 1) * pageSize, page * pageSize)

  const criar = useMutation({
    mutationFn: () =>
      createLocal({
        unidade_id: unidade.id,
        nome: nome.trim(),
        ...(descricao.trim() ? { descricao: descricao.trim() } : {}),
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['locais'] })
      setNome('')
      setDescricao('')
      toast('Local criado com sucesso!')
    },
    // O `detail` do backend (ex.: nome duplicado na unidade) é a mensagem exata a exibir.
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao criar local.'), 'error'),
  })

  const atualizar = useMutation({
    mutationFn: ({ id, dados }: { id: number; dados: LocalUpdateInput }) => updateLocal(id, dados),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['locais'] })
      setEditandoId(null)
      toast('Local atualizado com sucesso!')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao atualizar local.'), 'error'),
  })
  // Só a linha clicada mostra o spinner (as demais continuam usáveis assim que a mutação termina).
  const idEmAtualizacao = atualizar.isPending ? atualizar.variables?.id : undefined

  function handleCriar(e: React.FormEvent) {
    e.preventDefault()
    if (!nome.trim()) {
      toast('Informe o nome do local.', 'error')
      return
    }
    criar.mutate()
  }

  function iniciarEdicao(l: Local) {
    setEditandoId(l.id)
    setNomeEditado(l.nome)
  }

  function salvarNome(l: Local) {
    const novo = nomeEditado.trim()
    if (!novo) {
      toast('Informe o nome do local.', 'error')
      return
    }
    if (novo === l.nome) {
      setEditandoId(null)
      return
    }
    atualizar.mutate({ id: l.id, dados: { nome: novo } })
  }

  return (
    <div className="space-y-4">
      <DialogHeader>
        <DialogTitle>Locais — {upper(unidade.nome)}</DialogTitle>
        <DialogDescription>
          Locais físicos desta unidade (salas, almoxarifados...). Locais desativados deixam de ser
          oferecidos nos cadastros, mas continuam nos itens já vinculados.
        </DialogDescription>
      </DialogHeader>

      <form onSubmit={handleCriar} noValidate className="flex flex-wrap items-end gap-2">
        <Input
          aria-label="Nome do novo local"
          placeholder="Nome do local"
          value={nome}
          onChange={(e) => setNome(e.target.value)}
          className="min-w-[160px] flex-1"
        />
        <Input
          aria-label="Descrição do novo local"
          placeholder="Descrição (opcional)"
          value={descricao}
          onChange={(e) => setDescricao(e.target.value)}
          className="min-w-[160px] flex-1"
        />
        <Button type="submit" loading={criar.isPending}>
          <Plus aria-hidden="true" />
          Adicionar local
        </Button>
      </form>

      {isLoading ? (
        <LoadingState label="Carregando locais..." />
      ) : isError ? (
        <ErrorState error={error} onRetry={() => refetch()} />
      ) : locais.length === 0 ? (
        <EmptyState
          title="Nenhum local cadastrado"
          description="Adicione o primeiro local desta unidade acima."
        />
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
                {visiveis.map((l) => (
                  <tr key={l.id} className="border-b border-border last:border-0 hover:bg-surface-alt">
                    <td className="px-4 py-2.5 text-body-sm font-medium">
                      {editandoId === l.id ? (
                        <form
                          className="flex items-center gap-1"
                          onSubmit={(e) => {
                            e.preventDefault()
                            salvarNome(l)
                          }}
                        >
                          <Input
                            aria-label={`Novo nome do local ${l.nome}`}
                            value={nomeEditado}
                            onChange={(e) => setNomeEditado(e.target.value)}
                            autoFocus
                          />
                          <Button
                            type="submit"
                            size="sm"
                            aria-label="Salvar nome"
                            loading={idEmAtualizacao === l.id}
                          >
                            <Check aria-hidden="true" />
                          </Button>
                          <Button
                            type="button"
                            size="sm"
                            variant="ghost"
                            aria-label="Cancelar renomeação"
                            onClick={() => setEditandoId(null)}
                          >
                            <X aria-hidden="true" />
                          </Button>
                        </form>
                      ) : (
                        l.nome
                      )}
                    </td>
                    <td className="px-4 py-2.5 text-body-sm text-muted-foreground">{l.descricao || '—'}</td>
                    <td className="whitespace-nowrap px-4 py-2.5">
                      <StatusPill
                        status={l.ativo ? 'Ativo' : 'Inativo'}
                        tone={l.ativo ? 'available' : 'retired'}
                      />
                    </td>
                    <td className="whitespace-nowrap px-4 py-2.5">
                      <div className="flex gap-1">
                        <Button
                          size="sm"
                          variant="ghost"
                          aria-label={`Renomear local ${l.nome}`}
                          onClick={() => iniciarEdicao(l)}
                        >
                          <Pencil size={14} />
                        </Button>
                        {l.ativo ? (
                          <Button
                            size="sm"
                            variant="ghost"
                            aria-label={`Desativar local ${l.nome}`}
                            className="text-muted-foreground hover:text-destructive"
                            loading={idEmAtualizacao === l.id}
                            onClick={() => atualizar.mutate({ id: l.id, dados: { ativo: false } })}
                          >
                            <Ban size={14} />
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="ghost"
                            aria-label={`Reativar local ${l.nome}`}
                            loading={idEmAtualizacao === l.id}
                            onClick={() => atualizar.mutate({ id: l.id, dados: { ativo: true } })}
                          >
                            <RotateCcw size={14} />
                          </Button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="border-t border-border px-4 py-3">
            <Pagination
              page={page}
              pageSize={pageSize}
              total={locais.length}
              onPageChange={setPage}
              onPageSizeChange={(n) => {
                setPageSize(n)
                setPage(1)
              }}
            />
          </div>
        </div>
      )}
    </div>
  )
}
