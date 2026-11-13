import { useMemo, useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import { listPacotes, listSistemas, savePacote, type Pacote, type Sistema } from '@/api/acessos'
import { listCargos } from '@/api/colaboradores'
import { PageHeader } from '@/components/layout/PageHeader'
import { Button } from '@/components/ui/button'
import { Checkbox } from '@/components/ui/checkbox'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { ErrorState, LoadingState } from '@/components/ui/StateBlocks'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { cn } from '@/lib/utils'

/** Chave de comparação do cargo: sem acento, sem caixa e sem espaços repetidos. */
function chave(cargo: string): string {
  return cargo
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
}

function resumo(pacote: Pacote | undefined): string {
  const n = pacote?.sistemas.length ?? 0
  if (n === 0) return 'Sem pacote'
  return n === 1 ? '1 sistema' : `${n} sistemas`
}

/** Pacotes de sistemas por cargo (só Gestor): o que cada função recebe numa admissão. */
export default function PacotesPage() {
  const [busca, setBusca] = useState('')
  const [escolhido, setEscolhido] = useState<string | null>(null)

  const cargosQuery = useQuery({ queryKey: ['colaboradores', 'cargos'], queryFn: listCargos })
  const pacotesQuery = useQuery({ queryKey: ['acessos', 'pacotes'], queryFn: listPacotes })
  const sistemasQuery = useQuery({
    queryKey: ['acessos', 'sistemas', { include_inactive: false }],
    queryFn: () => listSistemas(),
  })

  const pacotes = pacotesQuery.data
  const linhas = useMemo(() => {
    const porChave = new Map<string, { cargo: string; semColaboradores: boolean; pacote?: Pacote }>()
    for (const cargo of cargosQuery.data ?? []) {
      if (!porChave.has(chave(cargo))) porChave.set(chave(cargo), { cargo, semColaboradores: false })
    }
    for (const pacote of pacotes ?? []) {
      const atual = porChave.get(chave(pacote.cargo))
      if (atual) atual.pacote = pacote
      else porChave.set(chave(pacote.cargo), { cargo: pacote.cargo, semColaboradores: true, pacote })
    }
    return [...porChave.values()].sort((a, b) => a.cargo.localeCompare(b.cargo, 'pt-BR'))
  }, [cargosQuery.data, pacotes])

  const visiveis = linhas.filter((l) => chave(l.cargo).includes(chave(busca)))
  const linhaEscolhida = linhas.find((l) => l.cargo === escolhido)

  if (cargosQuery.isLoading || pacotesQuery.isLoading || sistemasQuery.isLoading) {
    return <LoadingState label="Carregando pacotes..." />
  }
  const erro = cargosQuery.error ?? pacotesQuery.error ?? sistemasQuery.error
  if (erro) {
    return (
      <ErrorState
        error={erro}
        onRetry={() => {
          cargosQuery.refetch()
          pacotesQuery.refetch()
          sistemasQuery.refetch()
        }}
      />
    )
  }

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Acessos"
        title="Pacotes por função"
        description="Sistemas padrão de cada cargo: geram as tarefas de criar acesso numa admissão ou mudança de função."
      />

      <div className="grid gap-4 md:grid-cols-2">
        <section aria-label="Cargos" className="surface-panel space-y-3 p-4">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="busca-cargo">Buscar cargo</Label>
            <Input
              id="busca-cargo"
              aria-label="Buscar cargo"
              value={busca}
              onChange={(e) => setBusca(e.target.value)}
              placeholder="Digite parte do cargo"
            />
          </div>
          <ul className="space-y-1">
            {visiveis.map((l) => (
              <li key={l.cargo}>
                <button
                  type="button"
                  aria-pressed={l.cargo === escolhido}
                  onClick={() => setEscolhido(l.cargo)}
                  className={cn(
                    'flex w-full flex-col rounded border px-3 py-2 text-left text-sm transition-colors focus-visible:ring-1 focus-visible:ring-ring',
                    l.cargo === escolhido
                      ? 'border-border-strong bg-surface-alt'
                      : 'border-border hover:bg-surface-alt'
                  )}
                >
                  <span className="font-medium">{l.cargo}</span>
                  <span className="text-xs text-muted-foreground">
                    {resumo(l.pacote)}
                    {l.semColaboradores && ' · sem colaboradores com este cargo'}
                  </span>
                </button>
              </li>
            ))}
          </ul>
        </section>

        <section aria-label="Pacote do cargo" className="surface-panel p-4">
          {linhaEscolhida ? (
            <PacoteEditor
              key={linhaEscolhida.cargo}
              cargo={linhaEscolhida.cargo}
              pacote={linhaEscolhida.pacote}
              sistemasAtivos={sistemasQuery.data ?? []}
            />
          ) : (
            <p className="text-sm text-muted-foreground">Selecione um cargo para definir os sistemas padrão.</p>
          )}
        </section>
      </div>
    </div>
  )
}

function PacoteEditor({
  cargo,
  pacote,
  sistemasAtivos,
}: {
  cargo: string
  pacote: Pacote | undefined
  sistemasAtivos: Sistema[]
}) {
  const queryClient = useQueryClient()
  const idsAtivos = new Set(sistemasAtivos.map((s) => s.id))
  const [marcados, setMarcados] = useState<Set<number>>(
    () => new Set((pacote?.sistemas ?? []).filter((s) => idsAtivos.has(s.id)).map((s) => s.id))
  )
  const desativados = (pacote?.sistemas ?? []).filter((s) => !idsAtivos.has(s.id))

  const salvar = useMutation({
    mutationFn: () => savePacote({ cargo, sistema_ids: [...marcados].sort((a, b) => a - b) }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['acessos'] })
      toast('Pacote salvo.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao salvar pacote.'), 'error'),
  })

  function alternar(id: number, marcado: boolean) {
    setMarcados((atual) => {
      const novo = new Set(atual)
      if (marcado) novo.add(id)
      else novo.delete(id)
      return novo
    })
  }

  return (
    <div className="space-y-4">
      <h2 className="text-base font-semibold">{cargo}</h2>
      <ul className="space-y-2">
        {sistemasAtivos.map((s) => (
          <li key={s.id} className="flex items-center gap-2">
            <Checkbox
              id={`pacote-sistema-${s.id}`}
              aria-label={s.nome}
              checked={marcados.has(s.id)}
              onCheckedChange={(v) => alternar(s.id, v)}
            />
            <Label htmlFor={`pacote-sistema-${s.id}`} className="cursor-pointer">
              {s.nome}
            </Label>
          </li>
        ))}
      </ul>
      {desativados.length > 0 && (
        <p className="text-sm text-muted-foreground">
          Sistemas desativados neste pacote (não geram tarefas): {desativados.map((s) => s.nome).join(', ')}
        </p>
      )}
      <Button onClick={() => salvar.mutate()} loading={salvar.isPending}>
        Salvar pacote
      </Button>
    </div>
  )
}
