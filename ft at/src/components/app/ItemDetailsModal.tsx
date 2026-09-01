import { useQuery } from "@tanstack/react-query";
import { Link } from "@tanstack/react-router";
import { Cpu, User, Building2, HardDrive, Network, Zap, Info, Tag, Pencil } from "lucide-react";

import type { Item } from "@/api/items";
import { listItemPeripherals, type Peripheral } from "@/api/peripherals";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { StatusBadge } from "@/components/app/StatusBadge";
import { useAuth } from "@/lib/auth";
import { formatDate } from "@/lib/utils";

interface ItemDetailsModalProps {
  item: Item | null;
  onClose: () => void;
}

function Field({ label, value, mono }: { label: string; value?: string | null | undefined; mono?: boolean | undefined }) {
  return (
    <div className="min-w-0">
      <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">{label}</p>
      <p className={`text-sm font-medium text-foreground ${mono ? "font-mono" : ""}`}>{value || "-"}</p>
    </div>
  );
}

function Section({
  icon: Icon,
  title,
  children,
}: {
  icon: typeof Info;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div className="rounded-lg border border-border bg-card p-4 space-y-3">
      <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-foreground">
        <Icon className="size-3.5 text-primary" aria-hidden />
        <span>{title}</span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-3">{children}</div>
    </div>
  );
}

/** Porta o antigo drawer lateral como Dialog, mantendo os mesmos campos por tipo de equipamento e periféricos vinculados. */
export function ItemDetailsModal({ item, onClose }: ItemDetailsModalProps) {
  const { hasRole } = useAuth();

  const { data: peripherals = [] } = useQuery({
    queryKey: ["item-peripherals", item?.id],
    queryFn: () => listItemPeripherals(item!.id),
    enabled: Boolean(item?.id && item.peripheral_count && item.peripheral_count > 0),
  });

  const isLaptopOrPC = ["Computador", "Notebook"].includes(item?.tipo ?? "");
  const isNobreak = item?.tipo === "Nobreak";
  const isSwitch = item?.tipo === "Switch";

  return (
    <Dialog open={Boolean(item)} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="max-w-2xl max-h-[85vh] overflow-y-auto">
        {item ? (
          <>
            <DialogHeader>
              <div className="flex items-center gap-2 flex-wrap">
                <span className="font-mono text-xs font-semibold px-1.5 py-0.5 rounded bg-primary/10 text-primary border border-primary/20">
                  #{item.id}
                </span>
                <span className="text-xs font-medium text-muted-foreground">{item.tipo}</span>
                <StatusBadge status={item.status} />
              </div>
              <DialogTitle>
                {item.brand} {item.model}
              </DialogTitle>
              <DialogDescription>Detalhes do equipamento e periféricos vinculados.</DialogDescription>
            </DialogHeader>

            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                <div className="rounded-lg border border-border bg-muted/30 p-3 flex items-center gap-3">
                  <div className="h-8 w-8 rounded-md bg-blue-500/10 text-blue-600 dark:text-blue-400 flex items-center justify-center shrink-0">
                    <Building2 className="size-4" aria-hidden />
                  </div>
                  <div className="min-w-0">
                    <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">Unidade</p>
                    <p className="text-xs font-medium text-foreground truncate">{item.revenda || "Não definida"}</p>
                  </div>
                </div>
                <div className="rounded-lg border border-border bg-muted/30 p-3 flex items-center gap-3">
                  <div className="h-8 w-8 rounded-md bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 flex items-center justify-center shrink-0">
                    <User className="size-4" aria-hidden />
                  </div>
                  <div className="min-w-0">
                    <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">Usuário Atual</p>
                    <p className="text-xs font-medium text-foreground truncate">{item.assigned_to || "Nenhum"}</p>
                  </div>
                </div>
                <div className="rounded-lg border border-border bg-muted/30 p-3 flex items-center gap-3">
                  <div className="h-8 w-8 rounded-md bg-purple-500/10 text-purple-600 dark:text-purple-400 flex items-center justify-center shrink-0">
                    <Tag className="size-4" aria-hidden />
                  </div>
                  <div className="min-w-0">
                    <p className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">Identificador / NF</p>
                    <p className="text-xs font-medium text-foreground truncate">{item.identificador || item.nota_fiscal || "-"}</p>
                  </div>
                </div>
              </div>

              <Section icon={Info} title="Informações Cadastrais">
                <Field label="Marca" value={item.brand} />
                <Field label="Modelo" value={item.model} />
                <Field label="Nota Fiscal" value={item.nota_fiscal} />
                <Field label="Código Patrimonial" value={item.codigo_patrimonial} />
                <Field label="Fornecedor" value={item.fornecedor} />
                <Field label="Data de Cadastro" value={formatDate(item.date_registered)} />
              </Section>

              <Section icon={User} title="Dados de Alocação">
                <Field label="Funcionário Alocado" value={item.assigned_to} />
                <Field label="CPF do Colaborador" value={item.cpf} />
                <Field label="Setor" value={item.setor} />
                <Field label="Data do Empréstimo" value={formatDate(item.date_issued)} />
              </Section>

              {isLaptopOrPC && (
                <Section icon={HardDrive} title="Hardware & Sistema">
                  <Field label="Host / Nome da Máquina" value={item.host} mono />
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
                <Section icon={Zap} title="Infraestrutura">
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
                <div className="rounded-lg border border-border bg-card p-4 space-y-3">
                  <div className="flex items-center gap-2 border-b border-border pb-2 text-xs font-semibold text-foreground">
                    <Cpu className="size-3.5 text-emerald-600 dark:text-emerald-400" aria-hidden />
                    <span>Periféricos Vinculados ({item.peripheral_count})</span>
                  </div>
                  {peripherals.length > 0 ? (
                    <div className="space-y-1.5">
                      {peripherals.map((p: Peripheral) => (
                        <div
                          key={p.id}
                          className="flex items-center justify-between p-2 rounded-md bg-muted/40 border border-border text-xs"
                        >
                          <div className="flex items-center gap-2 min-w-0">
                            <Network className="size-3.5 text-primary shrink-0" aria-hidden />
                            <span className="font-medium text-foreground truncate">
                              {p.tipo}: {p.brand} {p.model}
                            </span>
                          </div>
                          <span className="text-[10px] font-mono font-medium px-1.5 py-0.5 bg-muted text-muted-foreground rounded shrink-0">
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

            <DialogFooter>
              <Button variant="outline" size="sm" onClick={onClose}>
                Fechar
              </Button>
              {hasRole("Gestor", "Técnico") && (
                <Button asChild size="sm" onClick={onClose}>
                  <Link to="/edit/$id" params={{ id: String(item.id) }}>
                    <Pencil className="size-3.5 mr-1.5" aria-hidden />
                    Editar
                  </Link>
                </Button>
              )}
            </DialogFooter>
          </>
        ) : null}
      </DialogContent>
    </Dialog>
  );
}
