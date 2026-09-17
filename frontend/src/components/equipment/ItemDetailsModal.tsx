import { createPortal } from 'react-dom'
import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { type Item, updateItem } from '@/api/items'
import { listItemPeripherals, type Peripheral } from '@/api/peripherals'
import { downloadSignedTerm, generateLoanTerm } from '@/api/loans'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { useAuth } from '@/contexts/AuthContext'
import { formatDate } from '@/lib/utils'
import { maskNotaFiscalInput, isValidNotaFiscal } from '@/lib/utils'
import { getErrorMessage } from '@/lib/api-error'
import { toast } from '@/components/ui/toast'
import {
  X,
  Pencil,
  Save,
  Loader2,
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

/** Gera o termo de responsabilidade e abre para download */
async function generateAndDownloadLoanTerm(itemId: number): Promise<void> {
  const blob = await generateLoanTerm(itemId)
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `termo_${itemId}.pdf`
  a.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 10000)
}

interface ItemDetailsModalProps {
  item: Item | null
  onClose: () => void
}

/** Campo de exibição somente leitura */
function Field({
  label,
  value,
  mono,
}: {
  label: string
  value?: string | null
  mono?: boolean
}) {
  return (
    <div className="min-w-0">
      <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
        {label}
      </p>
      <p className={`text-sm font-medium text-foreground mt-0.5 ${mono ? 'font-mono' : ''}`}>
        {value || '—'}
      </p>
    </div>
  )
}

/** Campo de edição inline */
function EditField({
  label,
  value,
  onChange,
  mono,
  placeholder,
}: {
  label: string
  value: string
  onChange: (v: string) => void
  mono?: boolean
  placeholder?: string
}) {
  return (
    <div className="min-w-0 flex flex-col gap-1">
      <Label className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
        {label}
      </Label>
      <Input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className={`h-8 text-sm ${mono ? 'font-mono' : ''}`}
        placeholder={placeholder}
      />
    </div>
  )
}

/** Seção de campos agrupados */
function Section({
  icon: Icon,
  title,
  children,
}: {
  icon: typeof Info
  title: string
  children: React.ReactNode
}) {
  return (
    <div className="rounded border border-border bg-card p-4 space-y-3">
      <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-foreground">
        <Icon className="size-3.5 text-muted-foreground" aria-hidden />
        <span>{title}</span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3">{children}</div>
    </div>
  )
}

const SPECIFIC_KEYS = [
  'identificador', 'dominio', 'host', 'endereco_fisico', 'cpu', 'ram', 'storage',
  'sistema', 'licenca', 'anydesk', 'setor', 'ip', 'mac', 'potencia_nominal',
  'autonomia_estimada', 'ip_snmp', 'codigo_patrimonial', 'responsavel',
  'local_instalacao', 'poe', 'quantidade_portas',
] as const

export function ItemDetailsModal({ item, onClose }: ItemDetailsModalProps) {
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

  // Hidrata o estado de edição quando o item muda
  useEffect(() => {
    if (item) {
      setBrand(item.brand ?? '')
      setModel(item.model ?? '')
      setNotaFiscal(item.nota_fiscal ?? '')
      setCodigoPatrimonial(item.codigo_patrimonial ?? '')
      setFornecedor(item.fornecedor ?? '')
      const fields: Record<string, string> = {}
      const record = item as unknown as Record<string, unknown>
      for (const k of SPECIFIC_KEYS) {
        const val = record[k]
        if (val) fields[k] = String(val)
      }
      setSpecificFields(fields)
    }
    setEditing(false)
  }, [item])

  async function handleViewSignedTerm(itemId: number) {
    try {
      setDownloadingTerm(true)
      const blob = await downloadSignedTerm(itemId)
      const url = URL.createObjectURL(blob)
      window.open(url, '_blank')
      window.setTimeout(() => URL.revokeObjectURL(url), 10000)
    } catch (err) {
      toast.error(getErrorMessage(err, 'Termo assinado não encontrado.'))
    } finally {
      setDownloadingTerm(false)
    }
  }

  function handleSave() {
    if (notaFiscal && !isValidNotaFiscal(notaFiscal)) {
      toast.error('Nota fiscal inválida. Informe 9 dígitos.')
      return
    }
    const data: Record<string, unknown> = {
      tipo: item!.tipo,
      brand,
      model,
      nota_fiscal: notaFiscal,
      codigo_patrimonial: codigoPatrimonial,
      fornecedor,
      revenda: item!.revenda,
      ...specificFields,
    }
    saveMutation.mutate(data)
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

  if (!item) return null

  const isLaptopOrPC = ['Computador', 'Notebook'].includes(item.tipo ?? '')
  const isNobreak = item.tipo === 'Nobreak'
  const isSwitch = item.tipo === 'Switch'

  const sf = (key: string) => specificFields[key] ?? ''
  const setSf = (key: string, value: string) =>
    setSpecificFields((prev) => ({ ...prev, [key]: value }))

  return createPortal(
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-black/50 animate-fade-in select-none"
      onClick={(e) => { if (e.target === e.currentTarget) onClose() }}
    >
      <div className="relative w-full max-w-2xl max-h-[85vh] flex flex-col surface-panel shadow-overlay overflow-hidden">
        {/* Cabeçalho */}
        <div className="flex items-start justify-between px-5 py-4 border-b border-border shrink-0">
          <div className="space-y-1 min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-mono text-[10px] font-semibold px-1.5 py-0.5 rounded border border-border bg-surface-alt text-muted-foreground">
                #{item.id}
              </span>
              <span className="text-xs font-medium text-muted-foreground">{item.tipo}</span>
              <StatusBadge status={item.status} />
              {editing && (
                <span className="text-[10px] font-semibold text-foreground bg-surface-alt border border-border-strong px-2 py-0.5 rounded uppercase tracking-wide">
                  Editando
                </span>
              )}
            </div>
            <h2 className="text-heading-sm text-foreground truncate">
              {editing
                ? `${brand || item.brand} ${model || item.model}`
                : `${item.brand} ${item.model}`}
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
            className="text-muted-foreground hover:text-foreground shrink-0 ml-2"
            aria-label="Fechar"
          >
            <X size={18} />
          </Button>
        </div>

        {/* Corpo com scroll */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {/* Cards de contexto rápido */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
            <div className="rounded border border-border bg-surface-alt p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded bg-surface border border-border text-muted-foreground flex items-center justify-center shrink-0">
                <Building2 className="size-4" aria-hidden />
              </div>
              <div className="min-w-0">
                <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Unidade
                </p>
                <p className="text-xs font-medium text-foreground truncate">
                  {item.revenda || 'Não definida'}
                </p>
              </div>
            </div>
            <div className="rounded border border-border bg-surface-alt p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded bg-surface border border-border text-muted-foreground flex items-center justify-center shrink-0">
                <User className="size-4" aria-hidden />
              </div>
              <div className="min-w-0">
                <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Usuário Atual
                </p>
                <p className="text-xs font-medium text-foreground truncate">
                  {item.assigned_to || 'Nenhum'}
                </p>
              </div>
            </div>
            <div className="rounded border border-border bg-surface-alt p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded bg-surface border border-border text-muted-foreground flex items-center justify-center shrink-0">
                <Tag className="size-4" aria-hidden />
              </div>
              <div className="min-w-0">
                <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
                  Identificador / NF
                </p>
                <p className="text-xs font-medium text-foreground font-mono truncate">
                  {item.identificador || item.nota_fiscal || '—'}
                </p>
              </div>
            </div>
          </div>

          {/* Informações Cadastrais */}
          <Section icon={Info} title="Informações Cadastrais">
            {editing ? (
              <>
                <EditField label="Marca" value={brand} onChange={setBrand} />
                <EditField label="Modelo" value={model} onChange={setModel} />
                <EditField
                  label="Nota Fiscal"
                  value={notaFiscal}
                  onChange={(v) => setNotaFiscal(maskNotaFiscalInput(v))}
                  placeholder="9 dígitos"
                />
                <EditField
                  label="Código Patrimonial"
                  value={codigoPatrimonial}
                  onChange={setCodigoPatrimonial}
                  mono
                />
                <EditField label="Fornecedor" value={fornecedor} onChange={setFornecedor} />
                <Field label="Data de Cadastro" value={formatDate(item.date_registered)} />
              </>
            ) : (
              <>
                <Field label="Marca" value={item.brand} />
                <Field label="Modelo" value={item.model} />
                <Field label="Nota Fiscal" value={item.nota_fiscal} />
                <Field label="Código Patrimonial" value={item.codigo_patrimonial} mono />
                <Field label="Fornecedor" value={item.fornecedor} />
                <Field label="Data de Cadastro" value={formatDate(item.date_registered)} />
              </>
            )}
          </Section>

          {/* Dados de Alocação */}
          <Section icon={User} title="Dados de Alocação">
            <Field label="Funcionário Alocado" value={item.assigned_to} />
            <Field label="CPF do Colaborador" value={item.cpf} />
            <Field label="Setor" value={item.setor} />
            <Field label="Data do Empréstimo" value={formatDate(item.date_issued)} />
            {Boolean(item.status === 'Indisponível' && item.assigned_to) && (
              <div className="col-span-2 sm:col-span-3 pt-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  disabled={downloadingTerm}
                  onClick={() => handleViewSignedTerm(item.id)}
                >
                  <FileDown className="size-3.5 mr-1.5" aria-hidden />
                  {downloadingTerm ? 'Baixando termo...' : 'Ver Termo Assinado'}
                </Button>
              </div>
            )}
            {item.status === 'Pendente' && (
              <div className="col-span-2 sm:col-span-3 pt-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={() => generateAndDownloadLoanTerm(item.id)}
                >
                  <FileDown className="size-3.5 mr-1.5" aria-hidden />
                  Baixar Termo de Responsabilidade
                </Button>
              </div>
            )}
          </Section>

          {/* Hardware & Sistema — Computador/Notebook */}
          {isLaptopOrPC && (
            <Section icon={HardDrive} title="Hardware & Sistema">
              {editing ? (
                <>
                  <EditField
                    label="Host / Nome da Máquina"
                    value={sf('host')}
                    onChange={(v) => setSf('host', v)}
                    mono
                  />
                  <EditField
                    label="Processador (CPU)"
                    value={sf('cpu')}
                    onChange={(v) => setSf('cpu', v)}
                  />
                  <EditField
                    label="Memória RAM"
                    value={sf('ram')}
                    onChange={(v) => setSf('ram', v)}
                  />
                  <EditField
                    label="Armazenamento"
                    value={sf('storage')}
                    onChange={(v) => setSf('storage', v)}
                  />
                  <EditField
                    label="Sistema Operacional"
                    value={sf('sistema')}
                    onChange={(v) => setSf('sistema', v)}
                  />
                  <EditField
                    label="Domínio Corporativo"
                    value={sf('dominio')}
                    onChange={(v) => setSf('dominio', v)}
                    mono
                  />
                  <EditField
                    label="Endereço MAC / Físico"
                    value={sf('endereco_fisico') || sf('mac')}
                    onChange={(v) => setSf('endereco_fisico', v)}
                    mono
                  />
                  <EditField
                    label="Endereço IP"
                    value={sf('ip')}
                    onChange={(v) => setSf('ip', v)}
                    mono
                  />
                  <EditField
                    label="AnyDesk ID"
                    value={sf('anydesk')}
                    onChange={(v) => setSf('anydesk', v)}
                    mono
                  />
                  <div className="col-span-2">
                    <EditField
                      label="Licença do Windows"
                      value={sf('licenca')}
                      onChange={(v) => setSf('licenca', v)}
                      mono
                    />
                  </div>
                </>
              ) : (
                <>
                  <Field label="Host / Nome da Máquina" value={item.host} mono />
                  <Field label="Processador (CPU)" value={item.cpu} />
                  <Field label="Memória RAM" value={item.ram} />
                  <Field label="Armazenamento" value={item.storage} />
                  <Field label="Sistema Operacional" value={item.sistema} />
                  <Field label="Domínio Corporativo" value={item.dominio} mono />
                  <Field
                    label="Endereço MAC / Físico"
                    value={item.endereco_fisico || item.mac}
                    mono
                  />
                  <Field label="Endereço IP" value={item.ip} mono />
                  <Field label="AnyDesk ID" value={item.anydesk} mono />
                  <div className="col-span-2">
                    <Field label="Licença do Windows" value={item.licenca} mono />
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
                    <EditField
                      label="Potência Nominal"
                      value={sf('potencia_nominal')}
                      onChange={(v) => setSf('potencia_nominal', v)}
                    />
                    <EditField
                      label="Autonomia Estimada"
                      value={sf('autonomia_estimada')}
                      onChange={(v) => setSf('autonomia_estimada', v)}
                    />
                    <EditField
                      label="IP da Placa SNMP"
                      value={sf('ip_snmp')}
                      onChange={(v) => setSf('ip_snmp', v)}
                      mono
                    />
                  </>
                ) : (
                  <>
                    <Field label="Potência Nominal" value={item.potencia_nominal} />
                    <Field label="Autonomia Estimada" value={item.autonomia_estimada} />
                    <Field label="IP da Placa SNMP" value={item.ip_snmp} mono />
                  </>
                )
              )}
              {isSwitch && (
                editing ? (
                  <>
                    <EditField
                      label="Quantidade de Portas"
                      value={sf('quantidade_portas')}
                      onChange={(v) => setSf('quantidade_portas', v)}
                    />
                    <EditField
                      label="Suporte PoE"
                      value={sf('poe')}
                      onChange={(v) => setSf('poe', v)}
                    />
                    <EditField
                      label="Endereço IP"
                      value={sf('ip')}
                      onChange={(v) => setSf('ip', v)}
                      mono
                    />
                  </>
                ) : (
                  <>
                    <Field label="Quantidade de Portas" value={item.quantidade_portas} />
                    <Field label="Suporte PoE" value={item.poe} />
                    <Field label="Endereço IP" value={item.ip} mono />
                  </>
                )
              )}
            </Section>
          )}

          {/* Periféricos Vinculados */}
          {Boolean(item.peripheral_count && item.peripheral_count > 0) && (
            <div className="rounded border border-border bg-card p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-foreground">
                <Cpu className="size-3.5 text-muted-foreground" aria-hidden />
                <span>Periféricos Vinculados ({item.peripheral_count})</span>
              </div>
              {peripherals.length > 0 ? (
                <div className="space-y-1.5">
                  {peripherals.map((p: Peripheral) => (
                    <div
                      key={p.id}
                      className="flex items-center justify-between p-2 rounded bg-surface-alt border border-border text-xs"
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <Network className="size-3.5 text-muted-foreground shrink-0" aria-hidden />
                        <span className="font-medium text-foreground truncate">
                          {p.tipo}: {p.brand} {p.model}
                        </span>
                      </div>
                      <span className="text-[10px] font-mono font-medium px-1.5 py-0.5 bg-surface text-muted-foreground border border-border rounded shrink-0">
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
        <div className="flex items-center justify-between gap-2 px-5 py-4 border-t border-border shrink-0 bg-surface-alt">
          {editing ? (
            <>
              <Button
                variant="outline"
                size="sm"
                onClick={cancelEdit}
                disabled={saveMutation.isPending}
              >
                <X className="size-3.5 mr-1.5" aria-hidden />
                Cancelar
              </Button>
              <Button size="sm" onClick={handleSave} disabled={saveMutation.isPending}>
                {saveMutation.isPending ? (
                  <Loader2 className="size-3.5 mr-1.5 animate-spin" aria-hidden />
                ) : (
                  <Save className="size-3.5 mr-1.5" aria-hidden />
                )}
                {saveMutation.isPending ? 'Salvando...' : 'Salvar Alterações'}
              </Button>
            </>
          ) : (
            <>
              <Button variant="outline" size="sm" onClick={onClose}>
                Fechar
              </Button>
              {canEdit && (
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => setEditing(true)}
                >
                  <Pencil className="size-3.5 mr-1.5" aria-hidden />
                  Editar
                </Button>
              )}
            </>
          )}
        </div>
      </div>
    </div>,
    document.body
  )
}
