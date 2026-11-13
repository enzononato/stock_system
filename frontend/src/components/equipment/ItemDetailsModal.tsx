import { createPortal } from 'react-dom'
import { useEffect, useState, type ReactNode } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import { type Item, updateItem } from '@/api/items'
import { listItemPeripherals, type Peripheral } from '@/api/peripherals'
import { downloadSignedTerm } from '@/api/loans'
import { StatusPill } from '@/components/ui/StatusPill'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { useAuth } from '@/contexts/AuthContext'
import { getErrorMessage } from '@/lib/api-error'
import { upper } from '@/lib/format'
import {
  formatDate,
  isValidNotaFiscal,
  maskNotaFiscalInput,
  isValidMac,
  maskMacInput,
  isValidIp,
} from '@/lib/utils'
import {
  X,
  Pencil,
  Save,
  FileDown,
  Cpu,
  User,
  Building2,
  HardDrive,
  Network,
  Zap,
  Info,
  Tag,
} from 'lucide-react'

interface ItemDetailsModalProps {
  item: Item | null
  onClose: () => void
}

/** Rótulo + valor somente leitura, usado fora do modo de edição (visual da FT_STC). */
function ReadOnlyField({ label, value, mono }: { label: string; value?: string | null; mono?: boolean }) {
  return (
    <div className="min-w-0">
      <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{label}</p>
      <p className={`text-sm font-medium text-foreground mt-0.5 ${mono ? 'font-mono' : ''}`}>{value || '—'}</p>
    </div>
  )
}

/**
 * Rótulo + input, usado no modo de edição — mesma grade/posição do campo
 * somente leitura equivalente. `id` liga `Label`/`Input` (acessibilidade e
 * `getByLabelText` nos testes); a FT_STC não ligava os dois.
 */
function EditableField({
  id,
  label,
  value,
  onChange,
  mono,
  placeholder,
}: {
  id: string
  label: string
  value: string
  onChange: (value: string) => void
  mono?: boolean
  placeholder?: string
}) {
  return (
    <div className="min-w-0 flex flex-col gap-1">
      <Label htmlFor={id} className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
        {label}
      </Label>
      <Input
        id={id}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className={`h-8 text-sm ${mono ? 'font-mono' : ''}`}
      />
    </div>
  )
}

/** Seção de campos agrupados (cartão com título e grade de 2/3 colunas). */
function Section({
  icon: Icon,
  title,
  children,
}: {
  icon: typeof Info
  title: string
  children: ReactNode
}) {
  return (
    <div className="rounded-md border border-border bg-surface p-4 space-y-3">
      <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-brand-text">
        <Icon className="size-3.5" aria-hidden />
        <span>{title}</span>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-4 gap-y-3">{children}</div>
    </div>
  )
}

/** Cartão de contexto rápido do topo do modal (Unidade, Usuário Atual, Identificador). */
function ContextCard({
  icon: Icon,
  label,
  value,
  mono,
}: {
  icon: typeof Info
  label: string
  value: string
  mono?: boolean
}) {
  return (
    <div className="rounded-md bg-brand-soft p-4 flex items-center gap-3">
      <div className="h-8 w-8 rounded bg-surface text-brand-text flex items-center justify-center shrink-0">
        <Icon className="size-4" aria-hidden />
      </div>
      <div className="min-w-0">
        <p className="text-[10px] font-semibold uppercase tracking-wider text-brand-soft-foreground">{label}</p>
        <p className={`text-xs font-medium text-brand-soft-foreground truncate ${mono ? 'font-mono' : ''}`}>{value}</p>
      </div>
    </div>
  )
}

/**
 * Campos específicos por tipo que a edição inline grava. Limitado de propósito
 * aos campos que já aparecem nas seções "Hardware & Sistema" (Desktop/Notebook)
 * e "Infraestrutura" (Nobreak/Switch) abaixo — a FT_STC incluía também
 * `codigo_patrimonial` nesta mesma lista, mas esse campo já tem estado próprio
 * (`codigoPatrimonial`, editado em Informações Cadastrais); misturar os dois
 * fazia o valor antigo, hidratado aqui, sobrescrever silenciosamente a edição
 * do usuário ao montar o payload de salvamento (o spread de `specificFields`
 * vinha depois no objeto enviado ao backend).
 */
const SPECIFIC_KEYS = [
  'host', 'cpu', 'ram', 'storage', 'sistema', 'dominio', 'endereco_fisico', 'ip',
  'anydesk', 'licenca', 'potencia_nominal', 'autonomia_estimada', 'ip_snmp',
  'quantidade_portas', 'poe',
] as const

export function ItemDetailsModal({ item, onClose }: ItemDetailsModalProps) {
  const navigate = useNavigate()
  const { hasRole } = useAuth()
  const queryClient = useQueryClient()
  const canEdit = hasRole('Gestor', 'Técnico')

  const [editing, setEditing] = useState(false)
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [notaFiscal, setNotaFiscal] = useState('')
  const [codigoPatrimonial, setCodigoPatrimonial] = useState('')
  const [fornecedor, setFornecedor] = useState('')
  const [specificFields, setSpecificFields] = useState<Record<string, string>>({})
  const [downloadingTerm, setDownloadingTerm] = useState(false)

  // Busca os periféricos vinculados caso existam
  const { data: peripherals = [] } = useQuery({
    queryKey: ['item-peripherals', item?.id],
    queryFn: () => listItemPeripherals(item!.id),
    enabled: Boolean(item?.id && item.peripheral_count && item.peripheral_count > 0),
  })

  const saveMutation = useMutation({
    mutationFn: (data: Record<string, unknown>) => updateItem(item!.id, data),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['items'] })
      toast.success('Equipamento atualizado com sucesso!')
      setEditing(false)
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, 'Erro ao salvar alterações.'))
    },
  })

  // Hidrata o formulário de edição sempre que o item exibido muda, e fecha
  // qualquer edição em andamento (evita misturar edição de um item com os
  // dados carregados de outro, caso o modal seja reaproveitado).
  useEffect(() => {
    if (item) {
      setBrand(item.brand ?? '')
      setModel(item.model ?? '')
      setNotaFiscal(item.nota_fiscal ?? '')
      setCodigoPatrimonial(item.codigo_patrimonial ?? '')
      setFornecedor(item.fornecedor ?? '')
      const record = item as unknown as Record<string, unknown>
      const fields: Record<string, string> = {}
      for (const key of SPECIFIC_KEYS) {
        const value = record[key]
        if (value) fields[key] = String(value)
      }
      setSpecificFields(fields)
    }
    setEditing(false)
  }, [item])

  if (!item) return null

  // Enum real de tipos (backend/app/core/config.py EQUIPMENT_TYPES) usa
  // "Desktop", nunca "Computador" — só Desktop e Notebook têm os campos de
  // hardware/rede abaixo (host, cpu, ram, storage, sistema, dominio, etc.),
  // conforme TypeSpecificFields.tsx.
  const temColaborador = item.colaborador_id != null || !!item.colaborador_nome
  const isLaptopOrPC = ['Desktop', 'Notebook'].includes(item.tipo || '')
  const isNobreak = item.tipo === 'Nobreak'
  const isSwitch = item.tipo === 'Switch'

  const sf = (key: string) => specificFields[key] ?? ''
  const setSf = (key: string, value: string) =>
    setSpecificFields((prev) => ({ ...prev, [key]: value }))

  async function handleViewSignedTerm() {
    try {
      setDownloadingTerm(true)
      const blob = await downloadSignedTerm(item!.id)
      const url = URL.createObjectURL(blob)
      window.open(url, '_blank')
      window.setTimeout(() => URL.revokeObjectURL(url), 10000)
    } catch (err) {
      toast.error(getErrorMessage(err, 'Termo assinado não encontrado.'))
    } finally {
      setDownloadingTerm(false)
    }
  }

  function cancelEdit() {
    if (item) {
      setBrand(item.brand ?? '')
      setModel(item.model ?? '')
      setNotaFiscal(item.nota_fiscal ?? '')
      setCodigoPatrimonial(item.codigo_patrimonial ?? '')
      setFornecedor(item.fornecedor ?? '')
    }
    setEditing(false)
  }

  function handleSave() {
    // Mesma validação de nota fiscal usada no cadastro (lib/utils): 9 dígitos.
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error('Nota fiscal inválida. Informe 9 dígitos.')
      return
    }
    // MAC e IP: o backend normaliza/valida (backend/app/schemas/validators.py)
    // e devolveria 422 para um valor malformado — validar aqui evita a viagem
    // ao servidor e dá um retorno imediato em português.
    if (isLaptopOrPC && sf('endereco_fisico') && !isValidMac(sf('endereco_fisico'))) {
      toast.error('Endereço MAC inválido. Use o formato AA:BB:CC:DD:EE:FF.')
      return
    }
    if ((isLaptopOrPC || isSwitch) && sf('ip') && !isValidIp(sf('ip'))) {
      toast.error('Endereço IP inválido.')
      return
    }
    if (isNobreak && sf('ip_snmp') && !isValidIp(sf('ip_snmp'))) {
      toast.error('IP da placa SNMP inválido.')
      return
    }

    const data: Record<string, unknown> = {
      brand,
      model,
      nota_fiscal: notaFiscal,
      codigo_patrimonial: codigoPatrimonial,
      fornecedor,
    }
    if (isLaptopOrPC) {
      data.host = sf('host')
      data.cpu = sf('cpu')
      data.ram = sf('ram')
      data.storage = sf('storage')
      data.sistema = sf('sistema')
      data.dominio = sf('dominio')
      data.endereco_fisico = sf('endereco_fisico')
      data.ip = sf('ip')
      data.anydesk = sf('anydesk')
      data.licenca = sf('licenca')
    }
    if (isNobreak) {
      data.potencia_nominal = sf('potencia_nominal')
      data.autonomia_estimada = sf('autonomia_estimada')
      data.ip_snmp = sf('ip_snmp')
    }
    if (isSwitch) {
      data.quantidade_portas = sf('quantidade_portas')
      data.poe = sf('poe')
      data.ip = sf('ip')
    }
    saveMutation.mutate(data)
  }

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-black/50 animate-fade-in select-none">
      <div className="relative w-full max-w-2xl max-h-[85vh] flex flex-col surface-panel shadow-overlay overflow-hidden">
        {/* Cabeçalho */}
        <div className="flex items-start justify-between px-5 py-4 border-b border-border shrink-0">
          <div className="space-y-1 min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="rounded-full bg-brand-soft px-2 py-0.5 font-mono text-xs font-semibold text-brand-soft-foreground">
                #{item.id}
              </span>
              <span className="text-xs font-medium text-muted-foreground">{item.tipo}</span>
              <StatusPill status={item.status} />
              {editing && (
                <span className="rounded-full bg-lime px-2 py-0.5 text-xs font-semibold text-lime-foreground">
                  Editando
                </span>
              )}
            </div>
            <h2 className="text-heading-sm text-foreground truncate">
              {editing ? `${brand || item.brand} ${model || item.model}` : `${item.brand} ${item.model}`}
            </h2>
            <p className="text-caption text-muted-foreground">
              {editing
                ? 'Edite os campos e clique em Salvar para confirmar.'
                : 'Detalhes do equipamento e periféricos vinculados.'}
            </p>
          </div>

          <Button
            variant="ghost"
            size="icon"
            onClick={onClose}
            aria-label="Fechar"
            className="text-muted-foreground hover:text-foreground shrink-0 ml-2"
          >
            <X size={18} />
          </Button>
        </div>

        {/* Corpo com scroll */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {/* Cartões de contexto rápido */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
            <ContextCard icon={Building2} label="Unidade" value={upper(item.revenda) || 'Não definida'} />
            <ContextCard icon={User} label="Usuário Atual" value={item.colaborador_nome || item.assigned_to || 'Nenhum'} />
            <ContextCard
              icon={Tag}
              label="Identificador / NF"
              value={item.identificador || item.nota_fiscal || '—'}
              mono
            />
          </div>

          {/* Informações Cadastrais */}
          <Section icon={Info} title="Informações Cadastrais">
            {editing ? (
              <>
                <EditableField id="edit-brand" label="Marca" value={brand} onChange={setBrand} />
                <EditableField id="edit-model" label="Modelo" value={model} onChange={setModel} />
                <EditableField
                  id="edit-nota-fiscal"
                  label="Nota Fiscal"
                  value={notaFiscal}
                  onChange={(v) => setNotaFiscal(maskNotaFiscalInput(v))}
                  placeholder="9 dígitos"
                />
                <EditableField
                  id="edit-codigo-patrimonial"
                  label="Código Patrimonial"
                  value={codigoPatrimonial}
                  onChange={setCodigoPatrimonial}
                  mono
                />
                <EditableField id="edit-fornecedor" label="Fornecedor" value={fornecedor} onChange={setFornecedor} />
                <ReadOnlyField label="Data de Cadastro" value={formatDate(item.date_registered)} />
              </>
            ) : (
              <>
                <ReadOnlyField label="Marca" value={item.brand} />
                <ReadOnlyField label="Modelo" value={item.model} />
                <ReadOnlyField label="Nota Fiscal" value={item.nota_fiscal} />
                <ReadOnlyField label="Código Patrimonial" value={item.codigo_patrimonial} mono />
                <ReadOnlyField label="Fornecedor" value={item.fornecedor} />
                <ReadOnlyField label="Data de Cadastro" value={formatDate(item.date_registered)} />
              </>
            )}
          </Section>

          {/* Dados de Alocação */}
          <Section icon={User} title="Dados de Alocação">
            <ReadOnlyField label="Funcionário Alocado" value={item.colaborador_nome || item.assigned_to} />
            <ReadOnlyField label="CPF do Colaborador" value={item.cpf} />
            <ReadOnlyField label="Setor do colaborador" value={item.colaborador_setor} />
            <ReadOnlyField label="Cargo" value={item.colaborador_cargo} />
            <ReadOnlyField
              label="Superior imediato"
              value={item.superior_nome || (temColaborador ? 'Não informado' : null)}
            />
            <ReadOnlyField label="Local" value={item.local_nome} />
            <ReadOnlyField label="Setor" value={item.setor} />
            <ReadOnlyField label="Data do Empréstimo" value={formatDate(item.date_issued)} />
            {/* Campos que a FT_STC não exibia, mantidos para não perder informação. */}
            <ReadOnlyField label="Responsável Técnico" value={item.responsavel} />
            <ReadOnlyField label="Local de Instalação" value={item.local_instalacao} />
            {item.status === 'Indisponível' && item.assigned_to && (
              <div className="sm:col-span-2 lg:col-span-3 pt-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  loading={downloadingTerm}
                  onClick={() => void handleViewSignedTerm()}
                >
                  <FileDown aria-hidden />
                  Ver Termo Assinado
                </Button>
              </div>
            )}
            {item.status?.startsWith('Pendente') && (
              <div className="sm:col-span-2 lg:col-span-3 pt-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => generateAndDownloadLoanTerm(item.id)}
                >
                  <FileDown aria-hidden />
                  Baixar Termo de Responsabilidade
                </Button>
              </div>
            )}
          </Section>

          {/* Hardware & Sistema — Desktop/Notebook */}
          {isLaptopOrPC && (
            <Section icon={HardDrive} title="Hardware & Sistema">
              {editing ? (
                <>
                  <EditableField id="edit-host" label="Host / Nome da Máquina" value={sf('host')} onChange={(v) => setSf('host', v)} mono />
                  <EditableField id="edit-cpu" label="Processador (CPU)" value={sf('cpu')} onChange={(v) => setSf('cpu', v)} />
                  <EditableField id="edit-ram" label="Memória RAM" value={sf('ram')} onChange={(v) => setSf('ram', v)} />
                  <EditableField id="edit-storage" label="Armazenamento" value={sf('storage')} onChange={(v) => setSf('storage', v)} />
                  <EditableField id="edit-sistema" label="Sistema Operacional" value={sf('sistema')} onChange={(v) => setSf('sistema', v)} />
                  <EditableField id="edit-dominio" label="Domínio Corporativo" value={sf('dominio')} onChange={(v) => setSf('dominio', v)} mono />
                  <EditableField
                    id="edit-endereco-fisico"
                    label="Endereço MAC / Físico"
                    value={sf('endereco_fisico')}
                    onChange={(v) => setSf('endereco_fisico', maskMacInput(v))}
                    mono
                  />
                  <EditableField id="edit-ip-laptop" label="Endereço IP" value={sf('ip')} onChange={(v) => setSf('ip', v)} mono />
                  <EditableField id="edit-anydesk" label="AnyDesk ID" value={sf('anydesk')} onChange={(v) => setSf('anydesk', v)} mono />
                  <div className="sm:col-span-2">
                    <EditableField id="edit-licenca" label="Licença do Windows" value={sf('licenca')} onChange={(v) => setSf('licenca', v)} mono />
                  </div>
                </>
              ) : (
                <>
                  <ReadOnlyField label="Host / Nome da Máquina" value={item.host} mono />
                  <ReadOnlyField label="Processador (CPU)" value={item.cpu} />
                  <ReadOnlyField label="Memória RAM" value={item.ram} />
                  <ReadOnlyField label="Armazenamento" value={item.storage} />
                  <ReadOnlyField label="Sistema Operacional" value={item.sistema} />
                  <ReadOnlyField label="Domínio Corporativo" value={item.dominio} mono />
                  <ReadOnlyField label="Endereço MAC / Físico" value={item.endereco_fisico || item.mac} mono />
                  <ReadOnlyField label="Endereço IP" value={item.ip} mono />
                  <ReadOnlyField label="AnyDesk ID" value={item.anydesk} mono />
                  <div className="sm:col-span-2">
                    <ReadOnlyField label="Licença do Windows" value={item.licenca} mono />
                  </div>
                </>
              )}
            </Section>
          )}

          {/* Infraestrutura — Nobreak / Switch */}
          {(isNobreak || isSwitch) && (
            <Section icon={Zap} title="Infraestrutura">
              {isNobreak && (
                editing ? (
                  <>
                    <EditableField id="edit-potencia-nominal" label="Potência Nominal" value={sf('potencia_nominal')} onChange={(v) => setSf('potencia_nominal', v)} />
                    <EditableField id="edit-autonomia-estimada" label="Autonomia Estimada" value={sf('autonomia_estimada')} onChange={(v) => setSf('autonomia_estimada', v)} />
                    <EditableField id="edit-ip-snmp" label="IP da Placa SNMP" value={sf('ip_snmp')} onChange={(v) => setSf('ip_snmp', v)} mono />
                  </>
                ) : (
                  <>
                    <ReadOnlyField label="Potência Nominal" value={item.potencia_nominal} />
                    <ReadOnlyField label="Autonomia Estimada" value={item.autonomia_estimada} />
                    <ReadOnlyField label="IP da Placa SNMP" value={item.ip_snmp} mono />
                  </>
                )
              )}
              {isSwitch && (
                editing ? (
                  <>
                    <EditableField id="edit-quantidade-portas" label="Quantidade de Portas" value={sf('quantidade_portas')} onChange={(v) => setSf('quantidade_portas', v)} />
                    <EditableField id="edit-poe" label="Suporte PoE" value={sf('poe')} onChange={(v) => setSf('poe', v)} />
                    <EditableField id="edit-ip-switch" label="Endereço IP" value={sf('ip')} onChange={(v) => setSf('ip', v)} mono />
                  </>
                ) : (
                  <>
                    <ReadOnlyField label="Quantidade de Portas" value={item.quantidade_portas} />
                    <ReadOnlyField label="Suporte PoE" value={item.poe} />
                    <ReadOnlyField label="Endereço IP" value={item.ip} mono />
                  </>
                )
              )}
            </Section>
          )}

          {/* Periféricos Vinculados */}
          {Boolean(item.peripheral_count && item.peripheral_count > 0) && (
            <div className="rounded-md border border-border bg-surface p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-brand-text">
                <Cpu className="size-3.5" aria-hidden />
                <span>Periféricos Vinculados ({item.peripheral_count})</span>
              </div>
              {peripherals.length > 0 ? (
                <div className="space-y-1.5">
                  {peripherals.map((p: Peripheral) => (
                    <div
                      key={p.id}
                      className="flex items-center justify-between p-2 rounded-md bg-brand-soft text-brand-soft-foreground text-xs"
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <Network className="size-3.5 shrink-0" aria-hidden />
                        <span className="font-medium truncate">
                          {p.tipo}: {p.brand} {p.model}
                        </span>
                      </div>
                      <span className="text-[10px] font-mono font-semibold px-1.5 py-0.5 bg-surface text-brand-text rounded-full shrink-0">
                        #{p.id}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-muted-foreground">Carregando periféricos...</p>
              )}
            </div>
          )}
        </div>

        {/* Rodapé de ações */}
        <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-4 border-t border-border shrink-0 bg-surface">
          {editing ? (
            <>
              <Button variant="outline" size="sm" onClick={cancelEdit} disabled={saveMutation.isPending}>
                <X aria-hidden />
                Cancelar
              </Button>
              <Button size="sm" onClick={handleSave} loading={saveMutation.isPending}>
                <Save aria-hidden />
                Salvar Alterações
              </Button>
            </>
          ) : (
            <>
              <Button variant="outline" size="sm" onClick={onClose}>
                Fechar
              </Button>
              <div className="flex flex-wrap items-center gap-2">
                {/* Página de detalhe com abas (dados, histórico, ...), aberta para todos os papéis. */}
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => {
                    onClose()
                    navigate(`/items/${item.id}`)
                  }}
                >
                  Ver página completa
                </Button>
                {canEdit && (
                  <>
                    {/* Atalho para o formulário completo (não existe na FT_STC), no mesmo estilo dos demais botões. */}
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => {
                        onClose()
                        navigate(`/edit/${item.id}`)
                      }}
                    >
                      Formulário Completo
                    </Button>
                    <Button size="sm" variant="outline" onClick={() => setEditing(true)}>
                      <Pencil aria-hidden />
                      Editar
                    </Button>
                  </>
                )}
              </div>
            </>
          )}
        </div>
      </div>
    </div>,
    document.body
  )
}
