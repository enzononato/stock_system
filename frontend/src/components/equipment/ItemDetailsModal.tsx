import { createPortal } from 'react-dom'
import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import { type Item, updateItem } from '@/api/items'
import { listItemPeripherals, type Peripheral } from '@/api/peripherals'
import { downloadSignedTerm } from '@/api/loans'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { toast } from '@/components/ui/toast'
import { generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { useAuth } from '@/contexts/AuthContext'
import { getErrorMessage } from '@/lib/api-error'
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
  Loader2,
  FileDown,
  Cpu,
  User,
  Building2,
  HardDrive,
  Monitor,
  Zap,
  Info,
  Tag,
} from 'lucide-react'

interface ItemDetailsModalProps {
  item: Item | null
  onClose: () => void
}

/** Rótulo + valor somente leitura, usado fora do modo de edição. */
function ReadOnlyField({ label, value, mono }: { label: string; value?: string | null; mono?: boolean }) {
  return (
    <div>
      <p className="text-caption text-muted-foreground">{label}</p>
      <p className={`text-body-sm text-foreground${mono ? ' num' : ''}`}>{value || '-'}</p>
    </div>
  )
}

/**
 * Rótulo + input, usado no modo de edição — mesma grade/posição do campo
 * somente leitura equivalente. `id` liga `Label`/`Input` (acessibilidade e
 * `getByLabelText` nos testes).
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
    <div className="flex flex-col gap-1">
      <Label htmlFor={id} className="text-caption text-muted-foreground">
        {label}
      </Label>
      <Input
        id={id}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className={mono ? 'num' : undefined}
      />
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
      <div className="relative w-full max-w-3xl max-h-[90vh] flex flex-col surface-panel shadow-overlay overflow-hidden">
        {/* Header Bar */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-md bg-surface-alt border border-border text-muted-foreground flex items-center justify-center">
              <Monitor size={20} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-caption num text-muted-foreground rounded-sm border border-border bg-surface-alt px-1.5 py-0.5">
                  #{item.id}
                </span>
                <span className="text-caption text-muted-foreground">{item.tipo}</span>
                <StatusBadge status={item.status} />
                {editing && (
                  <span className="text-caption font-semibold text-foreground bg-surface-alt border border-border-strong px-1.5 py-0.5 rounded-sm uppercase">
                    Editando
                  </span>
                )}
              </div>
              <h2 className="text-heading-sm text-foreground mt-0.5">
                {editing ? `${brand || item.brand} ${model || item.model}` : `${item.brand} ${item.model}`}
              </h2>
            </div>
          </div>

          <Button
            variant="ghost"
            size="icon"
            onClick={onClose}
            className="text-muted-foreground hover:text-foreground"
          >
            <X size={18} />
          </Button>
        </div>

        {/* Modal Scrollable Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-5">
          {/* Top Info Highlights Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div className="surface-panel p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-md bg-surface-alt text-muted-foreground flex items-center justify-center">
                <Building2 size={18} />
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Unidade</p>
                <p className="text-body-sm text-foreground truncate">{item.revenda || 'Não definida'}</p>
              </div>
            </div>

            <div className="surface-panel p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-md bg-surface-alt text-muted-foreground flex items-center justify-center">
                <User size={18} />
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Usuário Atual</p>
                <p className="text-body-sm text-foreground truncate">{item.assigned_to || 'Nenhum'}</p>
              </div>
            </div>

            <div className="surface-panel p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-md bg-surface-alt text-muted-foreground flex items-center justify-center">
                <Tag size={18} />
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Identificador / NF</p>
                <p className="text-body-sm num text-foreground truncate">
                  {item.identificador || item.nota_fiscal || '-'}
                </p>
              </div>
            </div>
          </div>

          {/* Section 1: Dados Gerais & Patrimônio */}
          <div className="surface-panel p-4 space-y-3">
            <div className="flex items-center gap-2 text-caption text-muted-foreground">
              <Info size={15} className="text-muted-foreground" />
              <span>Informações Cadastrais</span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
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
                    mono
                  />
                  <EditableField
                    id="edit-codigo-patrimonial"
                    label="Código Patrimonial"
                    value={codigoPatrimonial}
                    onChange={setCodigoPatrimonial}
                    mono
                  />
                  <EditableField id="edit-fornecedor" label="Fornecedor" value={fornecedor} onChange={setFornecedor} />
                  <ReadOnlyField label="Data de Cadastro" value={formatDate(item.date_registered)} mono />
                </>
              ) : (
                <>
                  <ReadOnlyField label="Marca" value={item.brand} />
                  <ReadOnlyField label="Modelo" value={item.model} />
                  <ReadOnlyField label="Nota Fiscal" value={item.nota_fiscal} mono />
                  <ReadOnlyField label="Código Patrimonial" value={item.codigo_patrimonial} mono />
                  <ReadOnlyField label="Fornecedor" value={item.fornecedor} />
                  <ReadOnlyField label="Data de Cadastro" value={formatDate(item.date_registered)} mono />
                </>
              )}
            </div>
          </div>

          {/* Section 2: Alocação & Colaborador */}
          <div className="surface-panel p-4 space-y-3">
            <div className="flex items-center gap-2 text-caption text-muted-foreground">
              <User size={15} className="text-muted-foreground" />
              <span>Dados de Alocação</span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
              <ReadOnlyField label="Funcionário Alocado" value={item.assigned_to} />
              <ReadOnlyField label="CPF do Colaborador" value={item.cpf} mono />
              <ReadOnlyField label="Setor" value={item.setor} />
              <ReadOnlyField label="Data do Empréstimo" value={formatDate(item.date_issued)} mono />
              <ReadOnlyField label="Responsável Técnico" value={item.responsavel} />
              <ReadOnlyField label="Local de Instalação" value={item.local_instalacao} />
              {item.status === 'Indisponível' && item.assigned_to && (
                <div className="col-span-2 sm:col-span-3">
                  <Button
                    variant="outline"
                    size="sm"
                    disabled={downloadingTerm}
                    onClick={() => void handleViewSignedTerm()}
                    className="rounded-md"
                  >
                    <FileDown size={14} />
                    {downloadingTerm ? 'Baixando termo...' : 'Ver Termo Assinado'}
                  </Button>
                </div>
              )}
            </div>
          </div>

          {/* Section 3: Hardware & Rede (Desktops / Notebooks) */}
          {isLaptopOrPC && (
            <div className="surface-panel p-4 space-y-3">
              <div className="flex items-center gap-2 text-caption text-muted-foreground">
                <HardDrive size={15} className="text-muted-foreground" />
                <span>Especificações de Hardware & Sistema</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
                {editing ? (
                  <>
                    <EditableField id="edit-host" label="Host / Nome da Máquina" value={sf('host')} onChange={(v) => setSf('host', v)} mono />
                    <EditableField id="edit-cpu" label="Processador (CPU)" value={sf('cpu')} onChange={(v) => setSf('cpu', v)} />
                    <EditableField id="edit-ram" label="Memória RAM" value={sf('ram')} onChange={(v) => setSf('ram', v)} />
                    <EditableField id="edit-storage" label="Armazenamento" value={sf('storage')} onChange={(v) => setSf('storage', v)} />
                    <EditableField id="edit-sistema" label="Sistema Operacional" value={sf('sistema')} onChange={(v) => setSf('sistema', v)} />
                    <EditableField id="edit-dominio" label="Domínio Corporativo" value={sf('dominio')} onChange={(v) => setSf('dominio', v)} />
                    <EditableField
                      id="edit-endereco-fisico"
                      label="Endereço MAC / Físico"
                      value={sf('endereco_fisico')}
                      onChange={(v) => setSf('endereco_fisico', maskMacInput(v))}
                      mono
                    />
                    <EditableField id="edit-ip-laptop" label="Endereço IP" value={sf('ip')} onChange={(v) => setSf('ip', v)} mono />
                    <EditableField id="edit-anydesk" label="AnyDesk ID" value={sf('anydesk')} onChange={(v) => setSf('anydesk', v)} mono />
                    <div className="col-span-2">
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
                    <ReadOnlyField label="Domínio Corporativo" value={item.dominio} />
                    <ReadOnlyField label="Endereço MAC / Físico" value={item.endereco_fisico || item.mac} mono />
                    <ReadOnlyField label="Endereço IP" value={item.ip} mono />
                    <ReadOnlyField label="AnyDesk ID" value={item.anydesk} mono />
                    <div className="col-span-2">
                      <ReadOnlyField label="Licença do Windows" value={item.licenca} mono />
                    </div>
                  </>
                )}
              </div>
            </div>
          )}

          {/* Section 4: Especificações Nobreak ou Switch */}
          {(isNobreak || isSwitch) && (
            <div className="surface-panel p-4 space-y-3">
              <div className="flex items-center gap-2 text-caption text-muted-foreground">
                <Zap size={15} className="text-muted-foreground" />
                <span>Especificações de Infraestrutura</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
                {isNobreak && (
                  editing ? (
                    <>
                      <EditableField id="edit-potencia-nominal" label="Potência Nominal" value={sf('potencia_nominal')} onChange={(v) => setSf('potencia_nominal', v)} mono />
                      <EditableField id="edit-autonomia-estimada" label="Autonomia Estimada" value={sf('autonomia_estimada')} onChange={(v) => setSf('autonomia_estimada', v)} mono />
                      <EditableField id="edit-ip-snmp" label="IP da Placa SNMP" value={sf('ip_snmp')} onChange={(v) => setSf('ip_snmp', v)} mono />
                    </>
                  ) : (
                    <>
                      <ReadOnlyField label="Potência Nominal" value={item.potencia_nominal} mono />
                      <ReadOnlyField label="Autonomia Estimada" value={item.autonomia_estimada} mono />
                      <ReadOnlyField label="IP da Placa SNMP" value={item.ip_snmp} mono />
                    </>
                  )
                )}
                {isSwitch && (
                  editing ? (
                    <>
                      <EditableField id="edit-quantidade-portas" label="Quantidade de Portas" value={sf('quantidade_portas')} onChange={(v) => setSf('quantidade_portas', v)} mono />
                      <EditableField id="edit-poe" label="Suporte PoE" value={sf('poe')} onChange={(v) => setSf('poe', v)} />
                      <EditableField id="edit-ip-switch" label="Endereço IP" value={sf('ip')} onChange={(v) => setSf('ip', v)} mono />
                    </>
                  ) : (
                    <>
                      <ReadOnlyField label="Quantidade de Portas" value={item.quantidade_portas} mono />
                      <ReadOnlyField label="Suporte PoE" value={item.poe} />
                      <ReadOnlyField label="Endereço IP" value={item.ip} mono />
                    </>
                  )
                )}
              </div>
            </div>
          )}

          {/* Section 5: Periféricos Vinculados */}
          {Boolean(item.peripheral_count && item.peripheral_count > 0) && (
            <div className="surface-panel p-4 space-y-3">
              <div className="flex items-center gap-2 text-caption text-muted-foreground">
                <Cpu size={15} className="text-muted-foreground" />
                <span>
                  Periféricos Vinculados (<span className="num">{item.peripheral_count}</span>)
                </span>
              </div>
              <div className="border-t border-border pt-3">
                {peripherals.length > 0 ? (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    {peripherals.map((p: Peripheral) => (
                      <div
                        key={p.id}
                        className="flex items-center justify-between p-2.5 rounded-md bg-surface border border-border"
                      >
                        <div className="flex items-center gap-2 min-w-0">
                          <Cpu size={14} className="text-muted-foreground flex-shrink-0" />
                          <span className="text-body-sm text-foreground truncate">
                            {p.tipo}: {p.brand} {p.model}
                          </span>
                        </div>
                        <span className="text-caption num text-muted-foreground px-1.5 py-0.5 bg-surface-alt rounded-sm">
                          #{p.id}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-body-sm text-muted-foreground">Carregando periféricos...</p>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer Actions */}
        <div className="flex items-center justify-between px-6 py-3.5 border-t border-border">
          {editing ? (
            <>
              <Button
                variant="outline"
                size="sm"
                onClick={cancelEdit}
                disabled={saveMutation.isPending}
                className="rounded-md"
              >
                Cancelar
              </Button>
              <Button
                variant="gradient"
                size="sm"
                onClick={handleSave}
                disabled={saveMutation.isPending}
                className="rounded-md"
              >
                {saveMutation.isPending ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
                {saveMutation.isPending ? 'Salvando...' : 'Salvar Alterações'}
              </Button>
            </>
          ) : (
            <>
              <Button variant="outline" size="sm" onClick={onClose} className="rounded-md">
                Fechar
              </Button>

              <div className="flex items-center gap-2">
                {item.status?.startsWith('Pendente') && (
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => generateAndDownloadLoanTerm(item.id)}
                    className="rounded-md"
                  >
                    <FileDown size={14} />
                    Baixar Termo
                  </Button>
                )}

                {canEdit && (
                  <>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => {
                        onClose()
                        navigate(`/edit/${item.id}`)
                      }}
                      className="rounded-md"
                    >
                      Formulário Completo
                    </Button>
                    <Button
                      variant="gradient"
                      size="sm"
                      onClick={() => setEditing(true)}
                      className="rounded-md"
                    >
                      <Pencil size={14} />
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
