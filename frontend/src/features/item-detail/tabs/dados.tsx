import type { ReactNode } from 'react'
import type { Item } from '@/api/items'
import { upper } from '@/lib/format'
import { formatDate } from '@/lib/utils'
import type { ItemDetailTab } from '../types'

type Linha = { label: string; value?: string | number | null; mono?: boolean }

function preenchidas(linhas: Linha[]) {
  return linhas.filter((l) => l.value !== undefined && l.value !== null && String(l.value).trim() !== '')
}

function Secao({ titulo, linhas }: { titulo: string; linhas: Linha[] }): ReactNode {
  if (linhas.length === 0) return null
  return (
    <section className="space-y-3">
      <h3 className="text-heading-sm text-foreground">{titulo}</h3>
      <dl className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-2 lg:grid-cols-3">
        {linhas.map((l) => (
          <div key={l.label} className="min-w-0">
            <dt className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{l.label}</dt>
            <dd className={`mt-0.5 break-words text-sm font-medium text-foreground ${l.mono ? 'font-mono' : ''}`}>
              {l.value}
            </dd>
          </div>
        ))}
      </dl>
    </section>
  )
}

function DadosTab({ item }: { item: Item }) {
  const geral = preenchidas([
    { label: 'Tipo', value: item.tipo },
    { label: 'Marca', value: item.brand },
    { label: 'Modelo', value: item.model },
    { label: 'Identificador', value: item.identificador, mono: true },
    { label: 'Código patrimonial', value: item.codigo_patrimonial, mono: true },
    { label: 'Status', value: item.status },
    { label: 'Unidade', value: upper(item.revenda) },
    { label: 'Setor', value: item.setor },
    { label: 'Nota fiscal', value: item.nota_fiscal, mono: true },
    { label: 'Fornecedor', value: item.fornecedor },
    { label: 'Cadastrado em', value: item.date_registered ? formatDate(item.date_registered) : null },
    { label: 'Emitido em', value: item.date_issued ? formatDate(item.date_issued) : null },
  ])
  const colaborador = preenchidas([
    { label: 'Colaborador', value: item.colaborador_nome ?? item.assigned_to },
    { label: 'Setor do colaborador', value: item.colaborador_setor },
    { label: 'Cargo', value: item.colaborador_cargo },
    { label: 'Superior imediato', value: item.superior_nome },
  ])
  const localizacao = preenchidas([
    { label: 'Local', value: item.local_nome },
    { label: 'Endereço físico', value: item.endereco_fisico },
    { label: 'Local de instalação', value: item.local_instalacao },
    { label: 'Responsável', value: item.responsavel },
  ])
  const tecnico = preenchidas([
    { label: 'Processador (CPU)', value: item.cpu },
    { label: 'Memória (RAM)', value: item.ram },
    { label: 'Armazenamento', value: item.storage },
    { label: 'Sistema operacional', value: item.sistema },
    { label: 'Licença', value: item.licenca },
    { label: 'Host', value: item.host, mono: true },
    { label: 'Domínio', value: item.dominio },
    { label: 'AnyDesk', value: item.anydesk, mono: true },
    { label: 'IP', value: item.ip, mono: true },
    { label: 'MAC', value: item.mac, mono: true },
    { label: 'IP SNMP', value: item.ip_snmp, mono: true },
    { label: 'Potência nominal', value: item.potencia_nominal },
    { label: 'Autonomia estimada', value: item.autonomia_estimada },
    { label: 'PoE', value: item.poe },
    { label: 'Quantidade de portas', value: item.quantidade_portas },
    { label: 'Periféricos vinculados', value: item.peripheral_count },
  ])

  return (
    <div className="space-y-6">
      <Secao titulo="Informações gerais" linhas={geral} />
      <Secao titulo="Colaborador" linhas={colaborador} />
      <Secao titulo="Localização" linhas={localizacao} />
      <Secao titulo="Hardware e rede" linhas={tecnico} />
    </div>
  )
}

const tab: ItemDetailTab = { id: 'dados', label: 'Dados', order: 10, Component: DadosTab }
export default tab
