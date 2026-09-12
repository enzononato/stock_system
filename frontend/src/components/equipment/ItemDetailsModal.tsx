import { createPortal } from 'react-dom'
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
  Calendar,
  Network,
  ShieldCheck,
  Zap,
  Info,
  Tag,
} from 'lucide-react'

interface ItemDetailsModalProps {
  item: Item | null
  onClose: () => void
}

export function ItemDetailsModal({ item, onClose }: ItemDetailsModalProps) {
  const navigate = useNavigate()
  const { hasRole } = useAuth()

  // Busca os periféricos vinculados caso existam
  const { data: peripherals = [] } = useQuery({
    queryKey: ['item-peripherals', item?.id],
    queryFn: () => listItemPeripherals(item!.id),
    enabled: Boolean(item?.id && item.peripheral_count && item.peripheral_count > 0),
  })

  if (!item) return null

  const isLaptopOrPC = ['Computador', 'Notebook'].includes(item.tipo || '')
  const isNobreak = item.tipo === 'Nobreak'
  const isSwitch = item.tipo === 'Switch'

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
              </div>
              <h2 className="text-heading-sm text-foreground mt-0.5">
                {item.brand} {item.model}
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
              <div>
                <p className="text-caption text-muted-foreground">Marca</p>
                <p className="text-body-sm text-foreground">{item.brand || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Modelo</p>
                <p className="text-body-sm text-foreground">{item.model || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Nota Fiscal</p>
                <p className="text-body-sm num text-foreground">{item.nota_fiscal || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Código Patrimonial</p>
                <p className="text-body-sm num text-foreground">{item.codigo_patrimonial || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Fornecedor</p>
                <p className="text-body-sm text-foreground">{item.fornecedor || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Data de Cadastro</p>
                <p className="text-body-sm num text-foreground">{formatDate(item.date_registered)}</p>
              </div>
            </div>
          </div>

          {/* Section 2: Alocação & Colaborador */}
          <div className="surface-panel p-4 space-y-3">
            <div className="flex items-center gap-2 text-caption text-muted-foreground">
              <User size={15} className="text-muted-foreground" />
              <span>Dados de Alocação</span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
              <div>
                <p className="text-caption text-muted-foreground">Funcionário Alocado</p>
                <p className="text-body-sm text-foreground">{item.assigned_to || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">CPF do Colaborador</p>
                <p className="text-body-sm num text-foreground">{item.cpf || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Setor</p>
                <p className="text-body-sm text-foreground">{item.setor || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Data do Empréstimo</p>
                <p className="text-body-sm num text-foreground">{formatDate(item.date_issued)}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Responsável Técnico</p>
                <p className="text-body-sm text-foreground">{item.responsavel || '-'}</p>
              </div>
              <div>
                <p className="text-caption text-muted-foreground">Local de Instalação</p>
                <p className="text-body-sm text-foreground">{item.local_instalacao || '-'}</p>
              </div>
            </div>
          </div>

          {/* Section 3: Hardware & Rede (Computadores / Notebooks) */}
          {isLaptopOrPC && (
            <div className="surface-panel p-4 space-y-3">
              <div className="flex items-center gap-2 text-caption text-muted-foreground">
                <HardDrive size={15} className="text-muted-foreground" />
                <span>Especificações de Hardware & Sistema</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 border-t border-border pt-3">
                <div>
                  <p className="text-caption text-muted-foreground">Host / Nome da Máquina</p>
                  <p className="text-body-sm num text-foreground">{item.host || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Processador (CPU)</p>
                  <p className="text-body-sm text-foreground">{item.cpu || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Memória RAM</p>
                  <p className="text-body-sm text-foreground">{item.ram || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Armazenamento</p>
                  <p className="text-body-sm text-foreground">{item.storage || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Sistema Operacional</p>
                  <p className="text-body-sm text-foreground">{item.sistema || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Domínio Corporativo</p>
                  <p className="text-body-sm text-foreground">{item.dominio || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Endereço MAC / Físico</p>
                  <p className="text-body-sm num text-foreground">{item.endereco_fisico || item.mac || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">Endereço IP</p>
                  <p className="text-body-sm num text-foreground">{item.ip || '-'}</p>
                </div>
                <div>
                  <p className="text-caption text-muted-foreground">AnyDesk ID</p>
                  <p className="text-body-sm num text-foreground">{item.anydesk || '-'}</p>
                </div>
                <div className="col-span-2">
                  <p className="text-caption text-muted-foreground">Licença do Windows</p>
                  <p className="text-body-sm num text-foreground truncate">{item.licenca || '-'}</p>
                </div>
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
                  <>
                    <div>
                      <p className="text-caption text-muted-foreground">Potência Nominal</p>
                      <p className="text-body-sm num text-foreground">{item.potencia_nominal || '-'}</p>
                    </div>
                    <div>
                      <p className="text-caption text-muted-foreground">Autonomia Estimada</p>
                      <p className="text-body-sm num text-foreground">{item.autonomia_estimada || '-'}</p>
                    </div>
                    <div>
                      <p className="text-caption text-muted-foreground">IP da Placa SNMP</p>
                      <p className="text-body-sm num text-foreground">{item.ip_snmp || '-'}</p>
                    </div>
                  </>
                )}
                {isSwitch && (
                  <>
                    <div>
                      <p className="text-caption text-muted-foreground">Quantidade de Portas</p>
                      <p className="text-body-sm num text-foreground">{item.quantidade_portas || '-'}</p>
                    </div>
                    <div>
                      <p className="text-caption text-muted-foreground">Suporte PoE</p>
                      <p className="text-body-sm text-foreground">{item.poe || '-'}</p>
                    </div>
                    <div>
                      <p className="text-caption text-muted-foreground">Endereço IP</p>
                      <p className="text-body-sm num text-foreground">{item.ip || '-'}</p>
                    </div>
                  </>
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

            {hasRole('Gestor', 'Técnico') && (
              <Button
                variant="gradient"
                size="sm"
                onClick={() => {
                  onClose()
                  navigate(`/edit/${item.id}`)
                }}
                className="rounded-md"
              >
                <Pencil size={14} />
                Editar Equipamento
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>,
    document.body
  )
}
