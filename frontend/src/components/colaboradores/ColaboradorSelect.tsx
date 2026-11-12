import { useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { X } from 'lucide-react'
import { getColaborador, listColaboradores, type Colaborador } from '@/api/colaboradores'
import { SearchableSelect, type SearchableOption } from '@/components/ui/SearchableSelect'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { upper } from '@/lib/format'
import { formatCpf } from '@/lib/utils'

export interface ColaboradorSelectProps {
  /** `id` do colaborador selecionado (ou `null`). */
  value: number | null
  onChange: (colaborador: Colaborador | null) => void
  /** Por padrão só oferece colaboradores `ativo` — desligados/afastados não recebem equipamento. */
  somenteAtivos?: boolean
  placeholder?: string
  disabled?: boolean
}

function toOption(c: Colaborador): SearchableOption {
  const partes = [formatCpf(c.cpf), c.cargo, upper(c.unidade_nome ?? c.unidade_rh)].filter(Boolean)
  return { value: String(c.id), label: upper(c.nome), subtitle: partes.join(' · ') }
}

/**
 * Seletor de colaborador com busca no servidor (nome, CPF, matrícula...). A
 * lista nunca é carregada inteira: cada pausa de digitação (300 ms) refaz a
 * consulta com `limit: 20`. O colaborador do `value` é buscado à parte quando
 * não está entre os resultados atuais, para o rótulo nunca sumir.
 */
export function ColaboradorSelect({
  value,
  onChange,
  somenteAtivos = true,
  placeholder = 'Selecione um colaborador...',
  disabled,
}: ColaboradorSelectProps) {
  const [search, setSearch] = useState('')
  const busca = useDebouncedValue(search.trim(), 300)

  const { data } = useQuery({
    queryKey: ['colaboradores', 'select', { busca, somenteAtivos }],
    queryFn: () =>
      listColaboradores({
        busca: busca || undefined,
        situacao: somenteAtivos ? 'ativo' : undefined,
        limit: 20,
      }),
    placeholderData: keepPreviousData,
  })
  const items = data?.items ?? []

  const selecionadoNaLista = value !== null && items.some((c) => c.id === value)
  const { data: selecionado } = useQuery({
    queryKey: ['colaboradores', 'detalhe', value],
    queryFn: () => getColaborador(value as number),
    enabled: value !== null && !selecionadoNaLista,
  })

  const todos = selecionado && !selecionadoNaLista && value === selecionado.id ? [selecionado, ...items] : items

  function handleChange(val: string) {
    const c = todos.find((x) => String(x.id) === val) ?? null
    setSearch('')
    onChange(c)
  }

  return (
    <div className="flex items-center gap-1">
      <SearchableSelect
        options={todos.map(toOption)}
        value={value === null ? '' : String(value)}
        onValueChange={handleChange}
        placeholder={placeholder}
        searchPlaceholder="Buscar por nome, CPF ou matrícula..."
        disabled={disabled}
        search={search}
        onSearchChange={setSearch}
      />
      {value !== null && !disabled && (
        <button
          type="button"
          aria-label="Limpar colaborador"
          onClick={() => onChange(null)}
          className="shrink-0 rounded p-2 text-muted-foreground hover:bg-surface-alt hover:text-foreground"
        >
          <X size={14} aria-hidden="true" />
        </button>
      )}
    </div>
  )
}
