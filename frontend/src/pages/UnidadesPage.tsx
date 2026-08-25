import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import {
  listUnidades,
  createUnidade,
  updateUnidade,
  deactivateUnidade,
  reactivateUnidade,
  getIndicadoresUnidade,
  type Unidade,
  type UnidadeInput,
  type IndicadoresUnidade,
} from '@/api/unidades'
import { Button } from '@/components/ui/button'
import { KpiCard } from '@/components/ui/KpiCard'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { DataTable } from '@/components/ui/DataTable'
import { toast } from '@/components/ui/toast'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from '@/components/ui/alert-dialog'
import type { ColumnDef } from '@tanstack/react-table'
import { Building2, Pencil, Ban, BarChart3, RotateCcw } from 'lucide-react'

// --- Validação e máscaras -----------------------------------------------------
// `lib/utils.ts` (fora da posse deste módulo) já tem `isValidCpf` seguindo o
// mesmo módulo 11 usado aqui para CNPJ — a lógica é replicada localmente em vez
// de importada, seguindo apenas o estilo daquela implementação como referência.

/**
 * Valida um CNPJ conferindo os dois dígitos verificadores (módulo 11).
 * Rejeita sequências com os 14 dígitos repetidos (ex: 11.111.111/1111-11),
 * que passariam pelo cálculo dos dígitos verificadores mas nunca são CNPJs reais.
 */
export function isValidCnpj(cnpj: string | null | undefined): boolean {
  if (!cnpj) return false
  const digits = cnpj.replace(/\D/g, '')
  if (digits.length !== 14) return false
  if (/^(\d)\1{13}$/.test(digits)) return false

  const calcCheckDigit = (base: string, weights: number[]): number => {
    let total = 0
    for (let i = 0; i < base.length; i++) {
      total += Number(base[i]) * weights[i]
    }
    const remainder = total % 11
    return remainder < 2 ? 0 : 11 - remainder
  }

  const weights1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
  const weights2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

  const dv1 = calcCheckDigit(digits.slice(0, 12), weights1)
  const dv2 = calcCheckDigit(digits.slice(0, 12) + String(dv1), weights2)

  return dv1 === Number(digits[12]) && dv2 === Number(digits[13])
}

/** Máscara progressiva de CNPJ (00.000.000/0000-00) aplicada enquanto o usuário digita. */
export function maskCnpjInput(value: string): string {
  return value
    .replace(/\D/g, '')
    .slice(0, 14)
    .replace(/(\d{2})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d)/, '$1.$2')
    .replace(/(\d{3})(\d)/, '$1/$2')
    .replace(/(\d{4})(\d{1,2})$/, '$1-$2')
}

/** CEP no formato 00000-000 — validação usada só quando o campo está preenchido (é opcional). */
export function isValidCep(cep: string | null | undefined): boolean {
  if (!cep) return false
  return /^\d{5}-\d{3}$/.test(cep)
}

/** Máscara progressiva de CEP (00000-000) aplicada enquanto o usuário digita. */
export function maskCepInput(value: string): string {
  return value.replace(/\D/g, '').slice(0, 8).replace(/(\d{5})(\d{1,3})$/, '$1-$2')
}

/** UF: exatamente 2 letras. A máscara já força maiúsculas e trunca em 2 caracteres. */
export function isValidUf(uf: string | null | undefined): boolean {
  if (!uf) return false
  return /^[A-Z]{2}$/.test(uf)
}

/** Restringe a digitação da UF a no máximo 2 letras, sempre maiúsculas. */
export function maskUfInput(value: string): string {
  return value.replace(/[^A-Za-z]/g, '').slice(0, 2).toUpperCase()
}

// --- Indicadores ---------------------------------------------------------------

/**
 * Ordem fixa dos 4 status de item exibidos no painel de indicadores — o
 * contrato da API garante que `itens_por_status` sempre traz essas 4 chaves,
 * mas ainda assim usamos `?? 0` como defesa caso alguma falte.
 */
const STATUS_ORDER = ['Disponível', 'Indisponível', 'Pendente', 'Pendente Devolução']

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-lg border bg-muted p-3">
      <p className="text-xs text-muted-foreground">{label}</p>
      <p className="text-lg font-semibold text-foreground">{value}</p>
    </div>
  )
}

function IndicadoresPanel({ dados }: { dados: IndicadoresUnidade }) {
  // "Sem movimento" não é um erro — é o estado normal de uma unidade recém
  // criada ou que nunca recebeu equipamentos. Sem este aviso explícito, uma
  // tela cheia de zeros pode passar a impressão de que os dados não carregaram.
  const semMovimento =
    dados.termos_emitidos === 0 &&
    dados.termos_confirmados === 0 &&
    dados.devolucoes_concluidas === 0 &&
    dados.itens_total === 0 &&
    dados.emprestimos_ativos === 0 &&
    dados.perifericos_vinculados === 0

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        <Stat label="Termos Emitidos" value={dados.termos_emitidos} />
        <Stat label="Termos Confirmados" value={dados.termos_confirmados} />
        <Stat label="Devoluções Concluídas" value={dados.devolucoes_concluidas} />
        <Stat label="Total de Itens" value={dados.itens_total} />
        <Stat label="Empréstimos Ativos" value={dados.emprestimos_ativos} />
        <Stat label="Periféricos Vinculados" value={dados.perifericos_vinculados} />
      </div>

      <div>
        <p className="text-sm font-medium text-muted-foreground mb-2">Itens por Status</p>
        <div className="flex flex-wrap gap-2">
          {STATUS_ORDER.map((status) => (
            <Badge key={status}>
              {status}: {dados.itens_por_status[status] ?? 0}
            </Badge>
          ))}
        </div>
      </div>

      {semMovimento && (
        <p className="text-sm text-muted-foreground italic">
          Esta unidade ainda não tem nenhuma movimentação registrada.
        </p>
      )}
    </div>
  )
}

// --- Página ---------------------------------------------------------------------

function errorDetail(err: unknown, fallback: string): string {
  return (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? fallback
}

const emptyForm = { nome: '', razaoSocial: '', cnpj: '', endereco: '', cep: '', cidade: '', uf: '' }

export default function UnidadesPage() {
  const queryClient = useQueryClient()

  const [editingId, setEditingId] = useState<number | null>(null)
  const [originalNome, setOriginalNome] = useState('')
  const [form, setForm] = useState(emptyForm)

  const [confirmRenameOpen, setConfirmRenameOpen] = useState(false)
  const [pendingData, setPendingData] = useState<UnidadeInput | null>(null)

  const [indicadoresUnidadeId, setIndicadoresUnidadeId] = useState<number | null>(null)

  // Por padrão a lista só traz as unidades ativas (o dia a dia de
  // Cadastrar/Emprestar). Sem esta alternância, uma unidade inativada some da
  // tela e o botão de reativar nunca fica alcançável de novo — o Gestor
  // precisa poder trazê-la de volta à vista para reativar.
  const [showInactive, setShowInactive] = useState(false)

  const { data: unidades = [], isLoading } = useQuery({
    queryKey: ['unidades', { includeInactive: showInactive }],
    queryFn: () => listUnidades(showInactive ? { include_inactive: true } : undefined),
  })

  const { data: indicadores, isLoading: indicadoresLoading } = useQuery({
    queryKey: ['unidades', indicadoresUnidadeId, 'indicadores'],
    queryFn: () => getIndicadoresUnidade(indicadoresUnidadeId as number),
    enabled: indicadoresUnidadeId !== null,
  })

  // Criar, editar e inativar uma unidade muda o que os seletores de
  // Cadastrar/Emprestar oferecem (GET /api/constants → revendas) — sem esta
  // invalidação, quem cria uma unidade aqui só a veria em outras telas depois
  // do staleTime de `useConstants` expirar (ver hooks/useConstants.ts).
  function invalidateAfterWrite() {
    queryClient.invalidateQueries({ queryKey: ['unidades'] })
    queryClient.invalidateQueries({ queryKey: ['constants'] })
  }

  function resetForm() {
    setEditingId(null)
    setOriginalNome('')
    setForm(emptyForm)
  }

  const createMutation = useMutation({
    mutationFn: (data: UnidadeInput) => createUnidade(data),
    onSuccess: () => {
      invalidateAfterWrite()
      resetForm()
      toast('Unidade criada com sucesso!')
    },
    onError: (err: unknown) => toast(errorDetail(err, 'Erro ao criar unidade.'), 'error'),
  })

  const updateMutation = useMutation({
    mutationFn: ({ id, data }: { id: number; data: UnidadeInput }) => updateUnidade(id, data),
    onSuccess: () => {
      invalidateAfterWrite()
      resetForm()
      toast('Unidade atualizada com sucesso!')
    },
    onError: (err: unknown) => toast(errorDetail(err, 'Erro ao atualizar unidade.'), 'error'),
  })

  // O backend recusa inativar unidade com itens ativos (400) — exibimos o
  // `detail` que ele devolve, e não uma mensagem genérica, porque é ele quem
  // sabe explicar o motivo exato (ex.: quantidade e tipos de itens presos).
  const deactivateMutation = useMutation({
    mutationFn: (id: number) => deactivateUnidade(id),
    onSuccess: () => {
      invalidateAfterWrite()
      toast('Unidade inativada.')
    },
    onError: (err: unknown) => toast(errorDetail(err, 'Erro ao inativar unidade.'), 'error'),
  })

  // Reverte a inativação (POST /api/unidades/{id}/reativar) — sem isso não
  // havia caminho de volta para uma unidade inativada por engano. Invalida
  // ['constants'] pelo mesmo motivo que as demais escritas: a unidade volta a
  // ser opção nos seletores de Cadastrar/Emprestar.
  const reactivateMutation = useMutation({
    mutationFn: (id: number) => reactivateUnidade(id),
    onSuccess: () => {
      invalidateAfterWrite()
      toast('Unidade reativada.')
    },
    onError: (err: unknown) => toast(errorDetail(err, 'Erro ao reativar unidade.'), 'error'),
  })

  function startEdit(u: Unidade) {
    setEditingId(u.id)
    setOriginalNome(u.nome)
    setForm({
      nome: u.nome,
      razaoSocial: u.razao_social,
      cnpj: u.cnpj,
      endereco: u.endereco ?? '',
      cep: u.cep ?? '',
      cidade: u.cidade ?? '',
      uf: u.uf ?? '',
    })
  }

  function buildAndValidate(): UnidadeInput | null {
    if (!form.nome.trim()) {
      toast('Informe o nome da unidade.', 'error')
      return null
    }
    if (!form.razaoSocial.trim()) {
      toast('Informe a razão social.', 'error')
      return null
    }
    if (!isValidCnpj(form.cnpj)) {
      toast('CNPJ inválido. Confira os dígitos digitados.', 'error')
      return null
    }
    if (form.cep.trim() && !isValidCep(form.cep)) {
      toast('CEP deve estar no formato 00000-000.', 'error')
      return null
    }
    if (form.uf.trim() && !isValidUf(form.uf)) {
      toast('UF deve ter exatamente 2 letras.', 'error')
      return null
    }
    return {
      nome: form.nome.trim(),
      razao_social: form.razaoSocial.trim(),
      cnpj: form.cnpj.trim(),
      endereco: form.endereco.trim() || undefined,
      cep: form.cep.trim() || undefined,
      cidade: form.cidade.trim() || undefined,
      uf: form.uf.trim() || undefined,
    }
  }

  function submit(data: UnidadeInput) {
    if (editingId !== null) {
      updateMutation.mutate({ id: editingId, data })
    } else {
      createMutation.mutate(data)
    }
  }

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    const data = buildAndValidate()
    if (!data) return

    // O nome da unidade é a chave que liga equipamentos e histórico a ela.
    // O backend cascateia a renomeação, mas o usuário precisa confirmar
    // explicitamente antes — sem isso, um erro de digitação renomeia tudo
    // silenciosamente.
    if (editingId !== null && data.nome !== originalNome) {
      setPendingData(data)
      setConfirmRenameOpen(true)
      return
    }
    submit(data)
  }

  function confirmRename() {
    if (pendingData) submit(pendingData)
    setConfirmRenameOpen(false)
    setPendingData(null)
  }

  function cancelRename() {
    setConfirmRenameOpen(false)
    setPendingData(null)
  }

  const isSaving = createMutation.isPending || updateMutation.isPending

  const columns: ColumnDef<Unidade, unknown>[] = [
    {
      accessorKey: 'nome',
      header: 'Nome',
      cell: ({ row }) => (
        <div className="flex items-center gap-2">
          <span className={row.original.is_active ? '' : 'text-muted-foreground line-through'}>
            {row.original.nome}
          </span>
          {!row.original.is_active && <Badge variant="danger">Inativa</Badge>}
        </div>
      ),
    },
    { accessorKey: 'razao_social', header: 'Razão Social' },
    { accessorKey: 'cnpj', header: 'CNPJ' },
    {
      id: 'localizacao',
      header: 'Cidade/UF',
      cell: ({ row }) => {
        const { cidade, uf } = row.original
        if (!cidade && !uf) return <span className="text-muted-foreground">—</span>
        return <span>{cidade || '—'}/{uf || '—'}</span>
      },
    },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button
            size="sm"
            variant="ghost"
            aria-label={`Indicadores da unidade ${row.original.nome}`}
            onClick={() => setIndicadoresUnidadeId(row.original.id)}
          >
            <BarChart3 size={14} />
          </Button>
          <Button
            size="sm"
            variant="ghost"
            aria-label={`Editar unidade ${row.original.nome}`}
            onClick={() => startEdit(row.original)}
          >
            <Pencil size={14} />
          </Button>
          {row.original.is_active && (
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button
                  size="sm"
                  variant="ghost"
                  aria-label={`Inativar unidade ${row.original.nome}`}
                  className="text-red-500 hover:text-red-700"
                  disabled={deactivateMutation.isPending}
                >
                  <Ban size={14} />
                </Button>
              </AlertDialogTrigger>
              <AlertDialogContent>
                <AlertDialogHeader>
                  <AlertDialogTitle>Inativar unidade "{row.original.nome}"?</AlertDialogTitle>
                  <AlertDialogDescription>
                    Unidades com itens ativos vinculados não podem ser inativadas. Esta ação não
                    pode ser desfeita.
                  </AlertDialogDescription>
                </AlertDialogHeader>
                <AlertDialogFooter>
                  <AlertDialogCancel>Cancelar</AlertDialogCancel>
                  <AlertDialogAction onClick={() => deactivateMutation.mutate(row.original.id)}>
                    Inativar
                  </AlertDialogAction>
                </AlertDialogFooter>
              </AlertDialogContent>
            </AlertDialog>
          )}
          {!row.original.is_active && (
            <Button
              size="sm"
              variant="ghost"
              aria-label={`Reativar unidade ${row.original.nome}`}
              className="text-green-600 hover:text-green-800"
              disabled={reactivateMutation.isPending}
              onClick={() => reactivateMutation.mutate(row.original.id)}
            >
              <RotateCcw size={14} />
            </Button>
          )}
        </div>
      ),
    },
  ]

  const unidadeSelecionadaIndicadores = unidades.find((u) => u.id === indicadoresUnidadeId)

  return (
    <div className="space-y-8">
      <div>
        <h2 className="text-h1 text-foreground">Unidades de Revenda</h2>
        <p className="text-caption mt-1">
          Dados jurídicos e fiscais das unidades (antigas "revendas") e indicadores por unidade
        </p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <KpiCard label="Unidades Cadastradas" value={unidades.length} tone="primary" highlighted />
        <KpiCard label="Ativas" value={unidades.filter((u) => u.is_active !== false).length} tone="success" />
        <KpiCard label="Inativas" value={unidades.filter((u) => u.is_active === false).length} tone="warning" />
      </div>

      {/* Criar / editar unidade */}
      <form onSubmit={handleSubmit} className="bg-card rounded-xl border p-6 space-y-4 max-w-5xl">
        <h3 className="font-medium flex items-center gap-2">
          <Building2 size={16} />
          {editingId !== null ? `Editar Unidade #${editingId}` : 'Nova Unidade'}
        </h3>

        {editingId !== null && (
          <p className="text-xs text-warning bg-warning/10 border border-warning/25 rounded-md px-3 py-2">
            O nome da unidade liga os equipamentos e todo o histórico a ela. Alterá-lo atualiza
            automaticamente essas referências — confirme antes de salvar se for esse o caso.
          </p>
        )}

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-nome">Nome *</Label>
            <Input
              id="unidade-nome"
              value={form.nome}
              onChange={(e) => setForm((f) => ({ ...f, nome: e.target.value }))}
              required
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-razao-social">Razão Social *</Label>
            <Input
              id="unidade-razao-social"
              value={form.razaoSocial}
              onChange={(e) => setForm((f) => ({ ...f, razaoSocial: e.target.value }))}
              required
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-cnpj">CNPJ *</Label>
            <Input
              id="unidade-cnpj"
              value={form.cnpj}
              onChange={(e) => setForm((f) => ({ ...f, cnpj: maskCnpjInput(e.target.value) }))}
              placeholder="00.000.000/0000-00"
              maxLength={18}
              required
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-endereco">Endereço</Label>
            <Input
              id="unidade-endereco"
              value={form.endereco}
              onChange={(e) => setForm((f) => ({ ...f, endereco: e.target.value }))}
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <div className="flex items-center gap-2">
              <Label htmlFor="unidade-cep">CEP</Label>
              {editingId !== null && !form.cep && (
                <Badge variant="warning" className="text-[10px]">Não informado</Badge>
              )}
            </div>
            <Input
              id="unidade-cep"
              value={form.cep}
              onChange={(e) => setForm((f) => ({ ...f, cep: maskCepInput(e.target.value) }))}
              placeholder="00000-000 (opcional)"
              maxLength={9}
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-cidade">Cidade</Label>
            <Input
              id="unidade-cidade"
              value={form.cidade}
              onChange={(e) => setForm((f) => ({ ...f, cidade: e.target.value }))}
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="unidade-uf">UF</Label>
            <Input
              id="unidade-uf"
              value={form.uf}
              onChange={(e) => setForm((f) => ({ ...f, uf: maskUfInput(e.target.value) }))}
              placeholder="Ex.: PE"
              maxLength={2}
            />
          </div>
        </div>

        <div className="flex justify-end gap-2">
          {editingId !== null && (
            <Button type="button" variant="ghost" onClick={resetForm}>
              Cancelar edição
            </Button>
          )}
          <Button type="submit" disabled={isSaving}>
            {isSaving ? 'Salvando...' : editingId !== null ? 'Salvar Alterações' : 'Criar Unidade'}
          </Button>
        </div>
      </form>

      {/* Confirmação explícita de renomeação — disparada programaticamente pelo
          submit acima, não por um AlertDialogTrigger comum. */}
      <AlertDialog
        open={confirmRenameOpen}
        onOpenChange={(open) => {
          if (!open) cancelRename()
        }}
      >
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Confirmar renomeação da unidade</AlertDialogTitle>
            <AlertDialogDescription>
              Você está alterando o nome de "{originalNome}" para "{pendingData?.nome}". O nome é a
              chave que liga os equipamentos e o histórico a esta unidade: o sistema atualiza essas
              referências automaticamente, mas confirme que é isso mesmo que deseja fazer.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel onClick={cancelRename}>Cancelar</AlertDialogCancel>
            <AlertDialogAction onClick={confirmRename}>Confirmar Renomeação</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* Lista */}
      <div className="flex items-center justify-between">
        <h3 className="font-medium text-foreground">Unidades cadastradas</h3>
        <label className="flex items-center gap-2 text-sm text-muted-foreground cursor-pointer select-none">
          <input
            type="checkbox"
            checked={showInactive}
            onChange={(e) => setShowInactive(e.target.checked)}
            className="rounded border-border"
          />
          Mostrar unidades inativas
        </label>
      </div>
      {isLoading ? (
        <div className="py-8 text-center text-muted-foreground">Carregando...</div>
      ) : (
        <DataTable data={unidades} columns={columns} searchPlaceholder="Buscar unidade..." />
      )}

      {/* Indicadores */}
      {indicadoresUnidadeId !== null && (
        <div className="bg-card rounded-xl border p-6 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-medium flex items-center gap-2">
              <BarChart3 size={16} />
              Indicadores — {unidadeSelecionadaIndicadores?.nome ?? `Unidade #${indicadoresUnidadeId}`}
            </h3>
            <Button variant="ghost" size="sm" onClick={() => setIndicadoresUnidadeId(null)}>
              Fechar
            </Button>
          </div>

          {indicadoresLoading ? (
            <p className="text-caption mt-1">Carregando indicadores...</p>
          ) : indicadores ? (
            <IndicadoresPanel dados={indicadores} />
          ) : (
            <p className="text-caption mt-1">Não foi possível carregar os indicadores.</p>
          )}
        </div>
      )}
    </div>
  )
}
