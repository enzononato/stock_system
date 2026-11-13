import { useState } from 'react'
import { keepPreviousData, useQuery } from '@tanstack/react-query'
import { X } from 'lucide-react'
import { getItem } from '@/api/items'
import { listAparelhos, type AparelhoOpcao } from '@/api/linhas'
import { SearchableSelect, type SearchableOption } from '@/components/ui/SearchableSelect'
import { useDebouncedValue } from '@/hooks/useDebouncedValue'
import { descreverAparelho } from '@/features/linhas/format'
import { upper } from '@/lib/format'

export interface AparelhoSelectProps {
  /** `id` do aparelho selecionado (ou `null`). */
  value: number | null
  onChange: (aparelho: AparelhoOpcao | null) => void
  /** Linha em edição: o aparelho dela continua na lista mesmo estando "ocupado". */
  linhaId?: number
  disabled?: boolean
  placeholder?: string
  /** Rótulo do aparelho atual quando ele já é conhecido (evita buscar o item, útil para aparelho baixado). */
  rotuloAtual?: string
}

function toOption(a: AparelhoOpcao): SearchableOption {
  const subtitulo = [upper(a.revenda), upper(a.responsavel_nome), a.codigo_patrimonial].filter(Boolean).join(' · ')
  return { value: String(a.id), label: descreverAparelho(a), subtitle: subtitulo || undefined }
}

/**
 * Seletor de celular/tablet livre, com busca no servidor (300 ms de debounce, `limit: 20`).
 * O aparelho do `value` é buscado à parte quando não está nos resultados, para o rótulo não sumir.
 */
export function AparelhoSelect({
  value,
  onChange,
  linhaId,
  disabled,
  placeholder = 'Selecione um celular ou tablet...',
  rotuloAtual,
}: AparelhoSelectProps) {
  const [search, setSearch] = useState('')
  const busca = useDebouncedValue(search.trim(), 300)

  const { data } = useQuery({
    queryKey: ['linhas', 'aparelhos', { busca, linhaId }],
    queryFn: () => listAparelhos({ busca: busca || undefined, linha_id: linhaId, limit: 20 }),
    placeholderData: keepPreviousData,
  })
  const opcoes = data ?? []

  const naLista = value !== null && opcoes.some((a) => a.id === value)
  const { data: item } = useQuery({
    queryKey: ['items', 'detalhe', value],
    queryFn: () => getItem(value as number),
    enabled: value !== null && !naLista && !rotuloAtual,
  })

  let todos = opcoes
  if (value !== null && !naLista) {
    if (item && item.id === value) {
      const atual: AparelhoOpcao = {
        id: item.id,
        tipo: item.tipo ?? '',
        brand: item.brand ?? null,
        model: item.model ?? null,
        identificador: item.identificador ?? null,
        codigo_patrimonial: item.codigo_patrimonial ?? null,
        revenda: item.revenda ?? null,
        status: item.status ?? null,
        responsavel_nome: item.colaborador_nome ?? item.assigned_to ?? null,
      }
      todos = [atual, ...opcoes]
    } else if (rotuloAtual) {
      todos = [
        {
          id: value, tipo: '', brand: null, model: null, identificador: null,
          codigo_patrimonial: null, revenda: null, status: null, responsavel_nome: null,
        },
        ...opcoes,
      ]
    }
  }
  const opcoesDoSelect = todos.map((a) =>
    a.id === value && !naLista && !item && rotuloAtual ? { value: String(a.id), label: rotuloAtual } : toOption(a)
  )

  function handleChange(val: string) {
    const a = todos.find((x) => String(x.id) === val) ?? null
    setSearch('')
    onChange(a)
  }

  return (
    <div className="flex items-center gap-1">
      <SearchableSelect
        options={opcoesDoSelect}
        value={value === null ? '' : String(value)}
        onValueChange={handleChange}
        placeholder={placeholder}
        searchPlaceholder="Buscar por marca, modelo, identificador ou responsável..."
        disabled={disabled}
        emptyText="Nenhum celular ou tablet livre encontrado."
        search={search}
        onSearchChange={setSearch}
      />
      {value !== null && !disabled && (
        <button
          type="button"
          aria-label="Limpar aparelho"
          onClick={() => onChange(null)}
          className="shrink-0 rounded p-2 text-muted-foreground hover:bg-surface-alt hover:text-foreground"
        >
          <X size={14} aria-hidden="true" />
        </button>
      )}
    </div>
  )
}
