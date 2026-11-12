import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import {
  createColaborador,
  listCargos,
  updateColaborador,
  type Colaborador,
  type ColaboradorInput,
  type SituacaoColaborador,
} from '@/api/colaboradores'
import { listUnidades } from '@/api/unidades'
import { ColaboradorSelect } from '@/components/colaboradores/ColaboradorSelect'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { getErrorMessage } from '@/lib/api-error'
import { formatTelefone, upper } from '@/lib/format'
import { formatCpf, isValidCpf, maskCpfInput } from '@/lib/utils'

const AVISO_RH = 'Vem da base de RH'

const selectClass =
  'flex h-9 w-full rounded border border-input bg-surface-alt px-3 py-2 text-sm text-foreground focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-40'

interface ColaboradorFormDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  /** `null` = cadastro manual de um novo colaborador; senão, edição. */
  colaborador: Colaborador | null
}

export function ColaboradorFormDialog({ open, onOpenChange, colaborador }: ColaboradorFormDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        <FormBody colaborador={colaborador} onClose={() => onOpenChange(false)} />
      </DialogContent>
    </Dialog>
  )
}

function FormBody({ colaborador, onClose }: { colaborador: Colaborador | null; onClose: () => void }) {
  const queryClient = useQueryClient()
  const editando = colaborador !== null
  // Campos da base de RH só podem ser alterados em colaboradores cadastrados à mão;
  // para `origem='rh'` a sincronização é a fonte da verdade.
  const baseEditavel = !colaborador || colaborador.origem === 'manual'

  const [form, setForm] = useState({
    cpf: colaborador ? formatCpf(colaborador.cpf) : '',
    nome: colaborador?.nome ?? '',
    matricula: colaborador?.matricula ?? '',
    cargo: colaborador?.cargo ?? '',
    setor: colaborador?.setor ?? '',
    unidadeId: colaborador?.unidade_id ? String(colaborador.unidade_id) : '',
    situacao: (colaborador?.situacao ?? 'ativo') as SituacaoColaborador,
    email: colaborador?.email ?? '',
    telefone: formatTelefone(colaborador?.telefone),
  })
  const [superiorId, setSuperiorId] = useState<number | null>(colaborador?.superior_id ?? null)
  const set = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) =>
    setForm((f) => ({ ...f, [key]: value }))

  const { data: unidades = [] } = useQuery({
    queryKey: ['unidades', { includeInactive: false }],
    queryFn: () => listUnidades(),
    enabled: baseEditavel,
  })
  const { data: cargos = [] } = useQuery({
    queryKey: ['colaboradores', 'cargos'],
    queryFn: listCargos,
    enabled: baseEditavel,
  })

  const mutation = useMutation({
    mutationFn: (data: ColaboradorInput) =>
      colaborador ? updateColaborador(colaborador.id, data) : createColaborador(data),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['colaboradores'] })
      toast(editando ? 'Colaborador atualizado com sucesso!' : 'Colaborador cadastrado com sucesso!')
      onClose()
    },
    onError: (err: unknown) =>
      toast(getErrorMessage(err, editando ? 'Erro ao atualizar colaborador.' : 'Erro ao cadastrar colaborador.'), 'error'),
  })

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    if (baseEditavel) {
      if (!editando && !isValidCpf(form.cpf)) {
        toast('CPF inválido. Confira os dígitos digitados.', 'error')
        return
      }
      if (!form.nome.trim()) {
        toast('Informe o nome do colaborador.', 'error')
        return
      }
    }

    const contato: ColaboradorInput = {
      email: form.email.trim() || null,
      telefone: form.telefone.trim() || null,
      superior_id: superiorId,
    }
    if (!baseEditavel) {
      mutation.mutate(contato)
      return
    }
    const base: ColaboradorInput = {
      nome: form.nome.trim(),
      cargo: form.cargo.trim() || null,
      setor: form.setor.trim() || null,
      unidade_id: form.unidadeId ? Number(form.unidadeId) : null,
      situacao: form.situacao,
    }
    mutation.mutate(
      editando
        ? { ...base, ...contato }
        : { ...base, ...contato, cpf: form.cpf.replace(/\D/g, ''), matricula: form.matricula.trim() || null }
    )
  }

  const avisoRh = !baseEditavel && (
    <span className="ml-2 text-caption font-normal text-muted-foreground">{AVISO_RH}</span>
  )

  return (
    <form onSubmit={handleSubmit} noValidate className="space-y-4">
      <DialogHeader>
        <DialogTitle>{editando ? `Editar ${upper(colaborador.nome)}` : 'Novo colaborador'}</DialogTitle>
        <DialogDescription>
          {editando
            ? baseEditavel
              ? 'Colaborador cadastrado manualmente: todos os campos podem ser editados.'
              : 'Os dados da base de RH não podem ser alterados aqui; e-mail, telefone e superior podem.'
            : 'Cadastro manual, para quem ainda não está na base de RH.'}
        </DialogDescription>
      </DialogHeader>

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-cpf">CPF{!editando && ' *'}{avisoRh}</Label>
          <Input
            id="colab-cpf"
            value={form.cpf}
            onChange={(e) => set('cpf', maskCpfInput(e.target.value))}
            placeholder="000.000.000-00"
            maxLength={14}
            readOnly={editando}
            disabled={editando}
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-nome">Nome{baseEditavel && ' *'}{avisoRh}</Label>
          <Input
            id="colab-nome"
            value={form.nome}
            onChange={(e) => set('nome', e.target.value)}
            disabled={!baseEditavel}
          />
        </div>
        {!editando && (
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="colab-matricula">Matrícula</Label>
            <Input id="colab-matricula" value={form.matricula} onChange={(e) => set('matricula', e.target.value)} />
          </div>
        )}
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-cargo">Cargo{avisoRh}</Label>
          <Input
            id="colab-cargo"
            list="colab-cargos"
            value={form.cargo}
            onChange={(e) => set('cargo', e.target.value)}
            disabled={!baseEditavel}
          />
          <datalist id="colab-cargos">
            {cargos.map((c) => (
              <option key={c} value={c} />
            ))}
          </datalist>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-setor">Setor{avisoRh}</Label>
          <Input
            id="colab-setor"
            value={form.setor}
            onChange={(e) => set('setor', e.target.value)}
            disabled={!baseEditavel}
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-unidade">Unidade{avisoRh}</Label>
          {baseEditavel ? (
            <select
              id="colab-unidade"
              className={selectClass}
              value={form.unidadeId}
              onChange={(e) => set('unidadeId', e.target.value)}
            >
              <option value="">Sem unidade</option>
              {unidades.map((u) => (
                <option key={u.id} value={u.id}>
                  {upper(u.nome)}
                </option>
              ))}
            </select>
          ) : (
            <Input
              id="colab-unidade"
              value={upper(colaborador?.unidade_nome ?? colaborador?.unidade_rh)}
              disabled
              readOnly
            />
          )}
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-situacao">Situação{avisoRh}</Label>
          <select
            id="colab-situacao"
            className={selectClass}
            value={form.situacao}
            onChange={(e) => set('situacao', e.target.value as SituacaoColaborador)}
            disabled={!baseEditavel}
          >
            <option value="ativo">Ativo</option>
            <option value="afastado">Afastado</option>
            <option value="desligado">Desligado</option>
          </select>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-email">E-mail</Label>
          <Input
            id="colab-email"
            type="email"
            value={form.email}
            onChange={(e) => set('email', e.target.value)}
            placeholder="nome@empresa.com.br"
          />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="colab-telefone">Telefone</Label>
          <Input
            id="colab-telefone"
            value={form.telefone}
            onChange={(e) => set('telefone', e.target.value)}
            placeholder="(00) 00000-0000"
          />
        </div>
        <div className="flex flex-col gap-1.5 sm:col-span-2">
          <Label>Superior</Label>
          <ColaboradorSelect
            value={superiorId}
            onChange={(c) => setSuperiorId(c?.id ?? null)}
            placeholder="Sem superior definido"
          />
        </div>
      </div>

      <DialogFooter className="gap-2 sm:gap-0">
        <Button type="button" variant="ghost" onClick={onClose}>
          Cancelar
        </Button>
        <Button type="submit" loading={mutation.isPending}>
          Salvar
        </Button>
      </DialogFooter>
    </form>
  )
}
