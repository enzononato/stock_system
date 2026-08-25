import { createPortal } from 'react-dom'
import { useEffect, useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { useNavigate } from 'react-router-dom'
import { type Item } from '@/api/items'
import { listItemPeripherals, type Peripheral } from '@/api/peripherals'
import { StatusBadge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { generateAndDownloadLoanTerm } from '@/components/equipment/ConfirmacaoTermo'
import { useAuth } from '@/contexts/AuthContext'
import { formatDate } from '@/lib/utils'
import {
  X,
  Pencil,
  FileDown,
  Cpu,
  User,
  Building2,
  HardDrive,
  Monitor,
  Network,
  Zap,
  Info,
  Tag,
} from 'lucide-react'

interface ItemDetailsModalProps {
  item: Item | null
  onClose: () => void
}

function Field({ label, value, mono, accent }: { label: string; value?: string | null; mono?: boolean; accent?: boolean }) {
  return (
    <div>
      <p className="text-label mb-0.5">{label}</p>
      <p className={`text-body font-medium ${mono ? 'font-mono' : ''} ${accent ? 'text-primary' : 'text-foreground'}`}>
        {value || '-'}
      </p>
    </div>
  )
}

function Section({ icon: Icon, iconClass, title, children }: { icon: typeof Info; iconClass: string; title: string; children: React.ReactNode }) {
  return (
    <div className="surface rounded-lg p-4 space-y-3">
      <div className="flex items-center gap-2 border-b border-border pb-2 text-foreground font-semibold text-xs">
        <Icon size={14} className={iconClass} />
        <span>{title}</span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3">{children}</div>
    </div>
  )
}

/** Drawer lateral (desliza da direita) — substitui o modal centralizado
 *  anterior, seguindo o padrão do design system para telas de detalhe. */
export function ItemDetailsModal({ item, onClose }: ItemDetailsModalProps) {
  const navigate = useNavigate()
  const { hasRole } = useAuth()
  const [visible, setVisible] = useState(false)

  const { data: peripherals = [] } = useQuery({
    queryKey: ['item-peripherals', item?.id],
    queryFn: () => listItemPeripherals(item!.id),
    enabled: Boolean(item?.id && item.peripheral_count && item.peripheral_count > 0),
  })

  useEffect(() => {
    if (item) requestAnimationFrame(() => setVisible(true))
    else setVisible(false)
  }, [item])

  useEffect(() => {
    if (!item) return
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [item, onClose])

  if (!item) return null

  const isLaptopOrPC = ['Computador', 'Notebook'].includes(item.tipo || '')
  const isNobreak = item.tipo === 'Nobreak'
  const isSwitch = item.tipo === 'Switch'

  return createPortal(
    <div className="fixed inset-0 z-50 select-none" role="dialog" aria-modal="true">
      {/* Backdrop */}
      <div
        className={`absolute inset-0 bg-black/60 backdrop-blur-sm transition-opacity duration-200 ${visible ? 'opacity-100' : 'opacity-0'}`}
        onClick={onClose}
      />

      {/* Drawer panel */}
      <div
        className={`absolute right-0 top-0 h-full w-full sm:w-[480px] bg-background border-l border-border shadow-2xl flex flex-col transition-transform duration-250 ease-out ${visible ? 'translate-x-0' : 'translate-x-full'}`}
      >
        {/* Header */}
        <div className="flex items-center justify-between px-5 py-4 border-b border-border bg-card shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <div className="h-10 w-10 rounded-lg bg-primary/15 border border-primary/25 text-primary flex items-center justify-center shrink-0">
              <Monitor size={18} />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[11px] font-mono font-semibold px-1.5 py-0.5 rounded bg-primary/10 text-primary border border-primary/20">
                  #{item.id}
                </span>
                <span className="text-[11px] font-medium text-muted-foreground">{item.tipo}</span>
                <StatusBadge status={item.status} />
              </div>
              <h2 className="text-h3 text-foreground mt-0.5 truncate">
                {item.brand} {item.model}
              </h2>
            </div>
          </div>

          <button
            onClick={onClose}
            className="h-8 w-8 shrink-0 rounded-full text-muted-foreground hover:text-foreground hover:bg-accent flex items-center justify-center transition-colors"
          >
            <X size={16} />
          </button>
        </div>

        {/* Scrollable body */}
        <div className="flex-1 overflow-y-auto p-4 space-y-4">
          {/* Highlights */}
          <div className="grid grid-cols-1 gap-2">
            <div className="surface rounded-lg p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded-md bg-info/10 text-info flex items-center justify-center shrink-0">
                <Building2 size={15} />
              </div>
              <div className="min-w-0">
                <p className="text-label">Unidade</p>
                <p className="text-xs font-medium text-foreground truncate">{item.revenda || 'Não definida'}</p>
              </div>
            </div>
            <div className="surface rounded-lg p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded-md bg-success/10 text-success flex items-center justify-center shrink-0">
                <User size={15} />
              </div>
              <div className="min-w-0">
                <p className="text-label">Usuário Atual</p>
                <p className="text-xs font-medium text-foreground truncate">{item.assigned_to || 'Nenhum'}</p>
              </div>
            </div>
            <div className="surface rounded-lg p-3 flex items-center gap-3">
              <div className="h-8 w-8 rounded-md bg-purple/10 text-purple-light flex items-center justify-center shrink-0">
                <Tag size={15} />
              </div>
              <div className="min-w-0">
                <p className="text-label">Identificador / NF</p>
                <p className="text-xs font-medium text-foreground truncate">{item.identificador || item.nota_fiscal || '-'}</p>
              </div>
            </div>
          </div>

          <Section icon={Info} iconClass="text-primary" title="Informações Cadastrais">
            <Field label="Marca" value={item.brand} />
            <Field label="Modelo" value={item.model} />
            <Field label="Nota Fiscal" value={item.nota_fiscal} />
            <Field label="Código Patrimonial" value={item.codigo_patrimonial} />
            <Field label="Fornecedor" value={item.fornecedor} />
            <Field label="Data de Cadastro" value={formatDate(item.date_registered)} />
          </Section>

          <Section icon={User} iconClass="text-info" title="Dados de Alocação">
            <Field label="Funcionário Alocado" value={item.assigned_to} />
            <Field label="CPF do Colaborador" value={item.cpf} />
            <Field label="Setor" value={item.setor} />
            <Field label="Data do Empréstimo" value={formatDate(item.date_issued)} />
          </Section>

          {isLaptopOrPC && (
            <Section icon={HardDrive} iconClass="text-purple-light" title="Hardware & Sistema">
              <Field label="Host / Nome da Máquina" value={item.host} mono accent />
              <Field label="Processador (CPU)" value={item.cpu} />
              <Field label="Memória RAM" value={item.ram} />
              <Field label="Armazenamento" value={item.storage} />
              <Field label="Sistema Operacional" value={item.sistema} />
              <Field label="Domínio Corporativo" value={item.dominio} />
              <Field label="Endereço MAC / Físico" value={item.endereco_fisico || item.mac} mono />
              <Field label="Endereço IP" value={item.ip} mono />
              <Field label="AnyDesk ID" value={item.anydesk} mono />
              <div className="col-span-2">
                <Field label="Licença do Windows" value={item.licenca} mono />
              </div>
            </Section>
          )}

          {(isNobreak || isSwitch) && (
            <Section icon={Zap} iconClass="text-warning" title="Infraestrutura">
              {isNobreak && (
                <>
                  <Field label="Potência Nominal" value={item.potencia_nominal} />
                  <Field label="Autonomia Estimada" value={item.autonomia_estimada} />
                  <Field label="IP da Placa SNMP" value={item.ip_snmp} mono />
                </>
              )}
              {isSwitch && (
                <>
                  <Field label="Quantidade de Portas" value={item.quantidade_portas} />
                  <Field label="Suporte PoE" value={item.poe} />
                  <Field label="Endereço IP" value={item.ip} mono />
                </>
              )}
            </Section>
          )}

          {Boolean(item.peripheral_count && item.peripheral_count > 0) && (
            <div className="surface rounded-lg p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-border pb-2 text-foreground font-semibold text-xs">
                <Cpu size={14} className="text-success" />
                <span>Periféricos Vinculados ({item.peripheral_count})</span>
              </div>
              {peripherals.length > 0 ? (
                <div className="space-y-1.5">
                  {peripherals.map((p: Peripheral) => (
                    <div key={p.id} className="flex items-center justify-between p-2 rounded-md bg-secondary/50 border border-border text-xs">
                      <div className="flex items-center gap-2 min-w-0">
                        <Network size={13} className="text-primary shrink-0" />
                        <span className="font-medium text-foreground truncate">{p.tipo}: {p.brand} {p.model}</span>
                      </div>
                      <span className="text-[10px] font-mono font-medium px-1.5 py-0.5 bg-secondary text-muted-foreground rounded shrink-0">
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

        {/* Footer actions */}
        <div className="flex items-center justify-between px-5 py-3.5 border-t border-border bg-card shrink-0">
          <Button variant="outline" size="sm" onClick={onClose}>Fechar</Button>
          <div className="flex items-center gap-2">
            {item.status?.startsWith('Pendente') && (
              <Button variant="outline" size="sm" onClick={() => generateAndDownloadLoanTerm(item.id)}>
                <FileDown size={14} className="mr-1.5" />
                Baixar Termo
              </Button>
            )}
            {hasRole('Gestor', 'Técnico') && (
              <Button variant="default" size="sm" onClick={() => { onClose(); navigate(`/edit/${item.id}`) }}>
                <Pencil size={14} className="mr-1.5" />
                Editar
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>,
    document.body
  )
}
