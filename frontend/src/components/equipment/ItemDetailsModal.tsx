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
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6 bg-slate-950/75 backdrop-blur-md animate-fade-in select-none">
      <div className="relative w-full max-w-3xl max-h-[90vh] flex flex-col rounded-3xl bg-white shadow-2xl border border-slate-200/80 overflow-hidden animate-scale-in">
        {/* Header Bar */}
        <div className="flex items-center justify-between px-6 py-4 bg-slate-950 text-white border-b border-slate-800">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-xl bg-indigo-600/30 border border-indigo-500/40 text-indigo-300 flex items-center justify-center font-bold">
              <Monitor size={20} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-mono font-bold px-2 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                  #{item.id}
                </span>
                <span className="text-xs font-bold text-slate-400">{item.tipo}</span>
                <StatusBadge status={item.status} />
              </div>
              <h2 className="text-lg font-bold text-white tracking-tight mt-0.5 font-heading">
                {item.brand} {item.model}
              </h2>
            </div>
          </div>

          <button
            onClick={onClose}
            className="h-9 w-9 rounded-full bg-slate-900 text-slate-400 hover:text-white hover:bg-slate-800 flex items-center justify-center transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* Modal Scrollable Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-5 bg-slate-50/50">
          {/* Top Info Highlights Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div className="glass-card rounded-2xl p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center">
                <Building2 size={18} />
              </div>
              <div>
                <p className="text-[10px] font-bold text-slate-400 uppercase">Unidade</p>
                <p className="text-xs font-bold text-slate-800 truncate">{item.revenda || 'Não definida'}</p>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-xl bg-sky-50 text-sky-600 flex items-center justify-center">
                <User size={18} />
              </div>
              <div>
                <p className="text-[10px] font-bold text-slate-400 uppercase">Usuário Atual</p>
                <p className="text-xs font-bold text-slate-800 truncate">{item.assigned_to || 'Nenhum'}</p>
              </div>
            </div>

            <div className="glass-card rounded-2xl p-3.5 flex items-center gap-3">
              <div className="h-9 w-9 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center">
                <Tag size={18} />
              </div>
              <div>
                <p className="text-[10px] font-bold text-slate-400 uppercase">Identificador / NF</p>
                <p className="text-xs font-bold text-slate-800 truncate">
                  {item.identificador || item.nota_fiscal || '-'}
                </p>
              </div>
            </div>
          </div>

          {/* Section 1: Dados Gerais & Patrimônio */}
          <div className="glass-card rounded-2xl p-4 space-y-3">
            <div className="flex items-center gap-2 border-b border-slate-100 pb-2 text-slate-800 font-bold text-xs">
              <Info size={15} className="text-indigo-600" />
              <span>Informações Cadastrais</span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 text-xs">
              <div>
                <p className="text-slate-400 font-medium">Marca</p>
                <p className="font-semibold text-slate-800">{item.brand || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Modelo</p>
                <p className="font-semibold text-slate-800">{item.model || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Nota Fiscal</p>
                <p className="font-semibold text-slate-800">{item.nota_fiscal || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Código Patrimonial</p>
                <p className="font-semibold text-slate-800">{item.codigo_patrimonial || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Fornecedor</p>
                <p className="font-semibold text-slate-800">{item.fornecedor || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Data de Cadastro</p>
                <p className="font-semibold text-slate-800">{formatDate(item.date_registered)}</p>
              </div>
            </div>
          </div>

          {/* Section 2: Alocação & Colaborador */}
          <div className="glass-card rounded-2xl p-4 space-y-3">
            <div className="flex items-center gap-2 border-b border-slate-100 pb-2 text-slate-800 font-bold text-xs">
              <User size={15} className="text-sky-600" />
              <span>Dados de Alocação</span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 text-xs">
              <div>
                <p className="text-slate-400 font-medium">Funcionário Alocado</p>
                <p className="font-semibold text-slate-800">{item.assigned_to || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">CPF do Colaborador</p>
                <p className="font-semibold text-slate-800">{item.cpf || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Setor</p>
                <p className="font-semibold text-slate-800">{item.setor || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Data do Empréstimo</p>
                <p className="font-semibold text-slate-800">{formatDate(item.date_issued)}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Responsável Técnico</p>
                <p className="font-semibold text-slate-800">{item.responsavel || '-'}</p>
              </div>
              <div>
                <p className="text-slate-400 font-medium">Local de Instalação</p>
                <p className="font-semibold text-slate-800">{item.local_instalacao || '-'}</p>
              </div>
            </div>
          </div>

          {/* Section 3: Hardware & Rede (Computadores / Notebooks) */}
          {isLaptopOrPC && (
            <div className="glass-card rounded-2xl p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2 text-slate-800 font-bold text-xs">
                <HardDrive size={15} className="text-violet-600" />
                <span>Especificações de Hardware & Sistema</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 text-xs">
                <div>
                  <p className="text-slate-400 font-medium">Host / Nome da Máquina</p>
                  <p className="font-mono font-bold text-indigo-600">{item.host || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Processador (CPU)</p>
                  <p className="font-semibold text-slate-800">{item.cpu || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Memória RAM</p>
                  <p className="font-semibold text-slate-800">{item.ram || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Armazenamento</p>
                  <p className="font-semibold text-slate-800">{item.storage || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Sistema Operacional</p>
                  <p className="font-semibold text-slate-800">{item.sistema || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Domínio Corporativo</p>
                  <p className="font-semibold text-slate-800">{item.dominio || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Endereço MAC / Físico</p>
                  <p className="font-mono text-slate-700">{item.endereco_fisico || item.mac || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">Endereço IP</p>
                  <p className="font-mono text-slate-700">{item.ip || '-'}</p>
                </div>
                <div>
                  <p className="text-slate-400 font-medium">AnyDesk ID</p>
                  <p className="font-mono font-bold text-slate-800">{item.anydesk || '-'}</p>
                </div>
                <div className="col-span-2">
                  <p className="text-slate-400 font-medium">Licença do Windows</p>
                  <p className="font-mono text-xs text-slate-600 truncate">{item.licenca || '-'}</p>
                </div>
              </div>
            </div>
          )}

          {/* Section 4: Especificações Nobreak ou Switch */}
          {(isNobreak || isSwitch) && (
            <div className="glass-card rounded-2xl p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2 text-slate-800 font-bold text-xs">
                <Zap size={15} className="text-amber-600" />
                <span>Especificações de Infraestrutura</span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3 text-xs">
                {isNobreak && (
                  <>
                    <div>
                      <p className="text-slate-400 font-medium">Potência Nominal</p>
                      <p className="font-semibold text-slate-800">{item.potencia_nominal || '-'}</p>
                    </div>
                    <div>
                      <p className="text-slate-400 font-medium">Autonomia Estimada</p>
                      <p className="font-semibold text-slate-800">{item.autonomia_estimada || '-'}</p>
                    </div>
                    <div>
                      <p className="text-slate-400 font-medium">IP da Placa SNMP</p>
                      <p className="font-mono text-slate-700">{item.ip_snmp || '-'}</p>
                    </div>
                  </>
                )}
                {isSwitch && (
                  <>
                    <div>
                      <p className="text-slate-400 font-medium">Quantidade de Portas</p>
                      <p className="font-semibold text-slate-800">{item.quantidade_portas || '-'}</p>
                    </div>
                    <div>
                      <p className="text-slate-400 font-medium">Suporte PoE</p>
                      <p className="font-semibold text-slate-800">{item.poe || '-'}</p>
                    </div>
                    <div>
                      <p className="text-slate-400 font-medium">Endereço IP</p>
                      <p className="font-mono text-slate-700">{item.ip || '-'}</p>
                    </div>
                  </>
                )}
              </div>
            </div>
          )}

          {/* Section 5: Periféricos Vinculados */}
          {Boolean(item.peripheral_count && item.peripheral_count > 0) && (
            <div className="glass-card rounded-2xl p-4 space-y-3">
              <div className="flex items-center gap-2 border-b border-slate-100 pb-2 text-slate-800 font-bold text-xs">
                <Cpu size={15} className="text-emerald-600" />
                <span>Periféricos Vinculados ({item.peripheral_count})</span>
              </div>
              {peripherals.length > 0 ? (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {peripherals.map((p: Peripheral) => (
                    <div
                      key={p.id}
                      className="flex items-center justify-between p-2.5 rounded-xl bg-white border border-slate-200/80 text-xs"
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <Cpu size={14} className="text-indigo-500 flex-shrink-0" />
                        <span className="font-semibold text-slate-700 truncate">
                          {p.tipo}: {p.brand} {p.model}
                        </span>
                      </div>
                      <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 bg-slate-100 text-slate-600 rounded">
                        #{p.id}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-slate-400">Carregando periféricos...</p>
              )}
            </div>
          )}
        </div>

        {/* Modal Footer Actions */}
        <div className="flex items-center justify-between px-6 py-3.5 bg-white border-t border-slate-200/80">
          <Button variant="outline" size="sm" onClick={onClose} className="rounded-xl">
            Fechar
          </Button>

          <div className="flex items-center gap-2">
            {item.status?.startsWith('Pendente') && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => generateAndDownloadLoanTerm(item.id)}
                className="rounded-xl border-amber-300 text-amber-800 bg-amber-50 hover:bg-amber-100"
              >
                <FileDown size={14} className="mr-1.5" />
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
                className="rounded-xl"
              >
                <Pencil size={14} className="mr-1.5" />
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

