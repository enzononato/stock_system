import { useEffect, useState } from "react";
import { useMutation, useQuery, useQueryClient, keepPreviousData } from "@tanstack/react-query";
import {
  ArrowDownLeft,
  ArrowUpRight,
  Download,
  History as HistoryIcon,
  Paperclip,
  RotateCcw,
  Search,
  SlidersHorizontal,
  AlertTriangle,
  ChevronDown,
  ChevronsUpDown,
} from "lucide-react";

import { listHistoryPaginated, reverseEntryWithPassword, type HistoryEntry } from "@/api/history";
import { downloadAuthenticated } from "@/api/client";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { useAuth } from "@/lib/auth";
import { getErrorMessage } from "@/lib/api-error";
import { cn, formatCpf, formatDateTime, exportToCsv } from "@/lib/utils";
import { toast } from "sonner";

const REVERSIBLE_OPS = [
  "Cadastro",
  "Empréstimo",
  "Confirmação Empréstimo",
  "Devolução",
  "Confirmação Devolução",
];
const PAGE_SIZE = 7;

function operationSymbol(op?: string) {
  if (!op) return "·";
  if (op.includes("Empréstimo")) return "↑";
  if (op.includes("Devolução")) return "↓";
  if (op === "Cadastro") return "+";
  if (op === "Exclusão") return "×";
  if (op === "Estorno") return "↺";
  return "·";
}

function operationColor(op?: string) {
  if (!op) return "text-muted-foreground";
  if (op.includes("Empréstimo")) return "text-foreground";
  if (op.includes("Devolução")) return "text-foreground";
  if (op === "Exclusão") return "text-foreground";
  if (op === "Estorno") return "text-muted-foreground";
  return "text-muted-foreground";
}

function attachmentFilename(key: string): string {
  const idx = key.indexOf("/");
  return idx >= 0 ? key.slice(idx + 1) : key;
}

export function HistoryPage() {
  const { hasRole } = useAuth();
  const queryClient = useQueryClient();

  const [pageIndex, setPageIndex] = useState(0);
  const [search, setSearch] = useState("");
  const [searchAplicado, setSearchAplicado] = useState("");
  const [downloadingKey, setDownloadingKey] = useState<string | null>(null);
  const [reversingEntry, setReversingEntry] = useState<HistoryEntry | null>(null);
  const [password, setPassword] = useState("");
  const [reverseError, setReverseError] = useState<string | null>(null);
  const [expandedIds, setExpandedIds] = useState<Record<number, boolean>>({});

  const { data, isLoading, error, refetch } = useQuery({
    queryKey: ["history", pageIndex, searchAplicado],
    queryFn: () =>
      listHistoryPaginated({
        search: searchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: pageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  });

  const history = data?.items ?? [];
  const total = data?.total ?? 0;
  const totalPages = Math.ceil(total / PAGE_SIZE);

  const allExpanded = history.length > 0 && history.every((e) => expandedIds[e.id]);

  function toggleExpand(id: number) {
    setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  }

  function toggleAll() {
    if (allExpanded) {
      setExpandedIds({});
    } else {
      const next: Record<number, boolean> = {};
      history.forEach((e) => {
        next[e.id] = true;
      });
      setExpandedIds(next);
    }
  }

  useEffect(() => {
    if (!data || data.total <= 0) return;
    const maxPageIndex = Math.max(0, Math.ceil(data.total / PAGE_SIZE) - 1);
    if (pageIndex > maxPageIndex) setPageIndex(maxPageIndex);
  }, [data, pageIndex]);

  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchAplicado(search.trim());
      setPageIndex(0);
    }, 400);
    return () => clearTimeout(timer);
  }, [search]);

  function openReverseConfirm(entry: HistoryEntry) {
    setReversingEntry(entry);
    setPassword("");
    setReverseError(null);
  }

  function closeReverseConfirm() {
    setReversingEntry(null);
    setPassword("");
    setReverseError(null);
  }

  const reverseMutation = useMutation({
    mutationFn: ({ id, password }: { id: number; password: string }) =>
      reverseEntryWithPassword(id, password),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["history"] });
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success("Operação estornada com sucesso!");
      closeReverseConfirm();
    },
    onError: (err: unknown) => {
      const status = (err as { response?: { status?: number } })?.response?.status;
      const msg =
        status === 403
          ? "Senha incorreta. Ação não autorizada."
          : getErrorMessage(err, "Erro ao estornar.");
      setReverseError(msg);
      toast.error(msg);
    },
  });

  function handleConfirmReverse() {
    if (!reversingEntry || !password) return;
    reverseMutation.mutate({ id: reversingEntry.id, password });
  }

  async function handleDownloadAttachment(key: string) {
    setDownloadingKey(key);
    try {
      await downloadAuthenticated(`/api/documents/files/${key}`, attachmentFilename(key));
    } catch (err) {
      toast.error(getErrorMessage(err, "Erro ao baixar anexo."));
    } finally {
      setDownloadingKey(null);
    }
  }

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Auditoria e Rastreabilidade
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Linha Temporal</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Histórico de Operações
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Registro cronológico e imutável de todas as ações realizadas no sistema patrimonial.
          </p>
        </div>
        <Button
          variant="outline"
          size="sm"
          disabled={history.length === 0}
          onClick={() =>
            exportToCsv("historico", history as unknown as Record<string, unknown>[], [
              { key: "id", label: "ID" },
              { key: "operador", label: "Operador" },
              { key: "operation", label: "Operação" },
              { key: "tipo", label: "Tipo" },
              { key: "marca", label: "Marca" },
              { key: "modelo", label: "Modelo" },
              { key: "identificador", label: "Identificador" },
              { key: "usuario", label: "Usuário" },
              { key: "setor", label: "Setor" },
              { key: "revenda", label: "Revenda" },
            ])
          }
          className="rounded-[4px] border-border text-foreground hover:bg-muted text-xs h-8"
        >
          <Download className="mr-1.5 size-3.5" />
          Exportar CSV
        </Button>
      </div>

      {/* Grid: 30% filtros / 70% timeline */}
      <div className="grid grid-cols-1 lg:grid-cols-[280px_1fr] gap-6 items-start">
        {/* Painel de Filtros */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-5">
          <div className="border-b border-border pb-2.5">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Filtros de Consulta
            </span>
          </div>

          <div className="space-y-2">
            <Label className="text-caption font-medium">Busca Geral</Label>
            <div className="relative">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
              <Input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Operador, usuário, operação…"
                className="pl-8 h-8 rounded-[4px] border-border bg-background text-body-sm"
              />
            </div>
          </div>

          {/* Sumário de Página */}
          <div className="space-y-2 pt-2 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              Sumário — Página Atual
            </span>
            <div className="space-y-2 text-caption">
              {[
                { label: "Total de Registros", value: total, symbol: "#" },
                {
                  label: "Empréstimos",
                  value: history.filter((h) => h.operation?.includes("Empréstimo")).length,
                  symbol: "↑",
                },
                {
                  label: "Devoluções",
                  value: history.filter((h) => h.operation?.includes("Devolução")).length,
                  symbol: "↓",
                },
                {
                  label: "Estornos",
                  value: history.filter((h) => h.operation === "Estorno").length,
                  symbol: "↺",
                },
              ].map((stat) => (
                <div key={stat.label} className="flex justify-between py-1 border-b border-border/50">
                  <span className="text-muted-foreground">{stat.label}</span>
                  <span className="font-mono font-semibold text-foreground">
                    {stat.symbol} {stat.value}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* Paginação */}
          {totalPages > 1 && (
            <div className="space-y-2 pt-2 border-t border-border">
              <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
                Paginação
              </span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pageIndex === 0}
                  onClick={() => setPageIndex((p) => p - 1)}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5 flex-1"
                >
                  ← Anterior
                </Button>
                <span className="text-caption font-mono text-foreground whitespace-nowrap">
                  {pageIndex + 1}/{totalPages}
                </span>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pageIndex >= totalPages - 1}
                  onClick={() => setPageIndex((p) => p + 1)}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5 flex-1"
                >
                  Próxima →
                </Button>
              </div>
            </div>
          )}
        </div>

        {/* Timeline Editorial */}
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-1">
          <div className="border-b border-border pb-3 mb-4 flex items-center justify-between">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Linha Temporal — {total} evento{total !== 1 ? "s" : ""} auditados
            </span>
            {history.length > 0 && (
              <Button
                type="button"
                size="sm"
                variant="ghost"
                onClick={toggleAll}
                className="h-6 px-2 text-[11px] text-muted-foreground hover:text-foreground rounded-[2px]"
              >
                <ChevronsUpDown className="mr-1 size-3" />
                {allExpanded ? "Recolher todos" : "Expandir todos"}
              </Button>
            )}
          </div>

          {isLoading ? (
            <div className="py-12 text-center text-caption text-muted-foreground">
              Carregando registros…
            </div>
          ) : history.length === 0 ? (
            <div className="py-12 text-center text-caption text-muted-foreground">
              Nenhum registro encontrado para os filtros aplicados.
            </div>
          ) : (
            <div className="relative">
              {/* Linha vertical da timeline */}
              <div className="absolute left-[13.5px] top-2 bottom-2 w-px bg-border" />

              <div className="space-y-0">
                {history.map((entry) => {
                  const symbol = operationSymbol(entry.operation);
                  const colorClass = operationColor(entry.operation);
                  const isReversible = REVERSIBLE_OPS.includes(entry.operation ?? "") && hasRole("Gestor");
                  const hasAttachments =
                    Boolean(entry.operacao_anexo) || Boolean(entry.termo_assinado_anexo);
                  const isExpanded = Boolean(expandedIds[entry.id]);

                  return (
                    <div key={entry.id} className="relative flex gap-3 pb-2.5 last:pb-0">
                      {/* Nó da Timeline reduzido */}
                      <div className="relative z-10 flex-shrink-0 mt-0.5">
                        <div
                          className={`size-7 rounded-[4px] border border-border bg-surface flex items-center justify-center font-mono text-caption font-bold shadow-2xs ${colorClass}`}
                        >
                          {symbol}
                        </div>
                      </div>

                      {/* Card da Linha Recolhível */}
                      <div className="flex-1 rounded-[4px] border border-border bg-surface-alt transition-colors duration-micro overflow-hidden min-w-0">
                        {/* Linha Resumo Clicável */}
                        <button
                          type="button"
                          onClick={() => toggleExpand(entry.id)}
                          className="w-full px-3 py-2 flex items-center justify-between gap-2.5 text-left cursor-pointer hover:bg-muted/40 transition-colors"
                          aria-expanded={isExpanded}
                        >
                          <div className="flex items-center gap-2 flex-wrap min-w-0">
                            <span className="font-semibold text-body-sm text-foreground">
                              {entry.operation ?? "—"}
                            </span>
                            {(entry.item_id || entry.peripheral_id) && (
                              <span className="font-mono text-caption text-muted-foreground">
                                {entry.item_id ? `Item #${entry.item_id}` : ""}
                                {entry.peripheral_id ? ` Periférico #${entry.peripheral_id}` : ""}
                              </span>
                            )}
                            {!isExpanded && (entry.usuario || entry.tipo || entry.operador) && (
                              <span className="text-caption text-muted-foreground truncate hidden sm:inline">
                                • {[entry.usuario, entry.tipo, entry.operador].filter(Boolean).slice(0, 2).join(" • ")}
                              </span>
                            )}
                          </div>

                          <div className="flex items-center gap-2.5 shrink-0">
                            <span className="font-mono text-caption text-muted-foreground whitespace-nowrap">
                              {formatDateTime(entry.data_operacao)}
                            </span>
                            <ChevronDown
                              className={cn(
                                "size-3.5 text-muted-foreground transition-transform duration-micro",
                                isExpanded && "rotate-180 text-foreground"
                              )}
                            />
                          </div>
                        </button>

                        {/* Detalhes Recolhíveis */}
                        {isExpanded && (
                          <div className="px-3 pb-3 pt-1 border-t border-border/50 space-y-2 bg-surface/50">
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-caption pt-1">
                              <div className="p-2 rounded-[4px] border border-border/60 bg-surface">
                                <span className="text-muted-foreground block text-[11px] uppercase tracking-wider">Operador</span>
                                <span className="font-medium text-foreground">{entry.operador || "—"}</span>
                              </div>
                              {entry.usuario && (
                                <div className="p-2 rounded-[4px] border border-border/60 bg-surface">
                                  <span className="text-muted-foreground block text-[11px] uppercase tracking-wider">Colaborador</span>
                                  <span className="font-medium text-foreground">{entry.usuario}</span>
                                </div>
                              )}
                              {(entry.tipo || entry.marca) && (
                                <div className="p-2 rounded-[4px] border border-border/60 bg-surface">
                                  <span className="text-muted-foreground block text-[11px] uppercase tracking-wider">Equipamento</span>
                                  <span className="font-medium text-foreground">
                                    {entry.tipo} {entry.marca} {entry.modelo}
                                  </span>
                                </div>
                              )}
                              {entry.revenda && (
                                <div className="p-2 rounded-[4px] border border-border/60 bg-surface">
                                  <span className="text-muted-foreground block text-[11px] uppercase tracking-wider">Revenda</span>
                                  <span className="text-foreground">{entry.revenda}</span>
                                </div>
                              )}
                            </div>

                            {entry.details && (
                              <p className="text-caption text-muted-foreground border-t border-border/50 pt-1.5">
                                {entry.details}
                              </p>
                            )}

                            {/* Ações da Entrada */}
                            {(hasAttachments || isReversible) && (
                              <div className="flex flex-wrap items-center gap-2 pt-1 border-t border-border/50">
                                {entry.operacao_anexo && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] rounded-[2px] border-border text-muted-foreground hover:text-foreground"
                                    disabled={downloadingKey === entry.operacao_anexo}
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleDownloadAttachment(entry.operacao_anexo!);
                                    }}
                                  >
                                    <Paperclip className="mr-1 size-2.5" />
                                    {downloadingKey === entry.operacao_anexo ? "Baixando…" : "Comprovante"}
                                  </Button>
                                )}
                                {entry.termo_assinado_anexo && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] rounded-[2px] border-border text-muted-foreground hover:text-foreground"
                                    disabled={downloadingKey === entry.termo_assinado_anexo}
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleDownloadAttachment(entry.termo_assinado_anexo!);
                                    }}
                                  >
                                    <Paperclip className="mr-1 size-2.5" />
                                    {downloadingKey === entry.termo_assinado_anexo ? "Baixando…" : "Termo"}
                                  </Button>
                                )}
                                {isReversible && (
                                  <Button
                                    type="button"
                                    size="sm"
                                    variant="outline"
                                    className="h-6 px-2 text-[11px] rounded-[2px] border-border text-muted-foreground hover:text-foreground ml-auto"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      openReverseConfirm(entry);
                                    }}
                                  >
                                    <RotateCcw className="mr-1 size-2.5" />
                                    Estornar
                                  </Button>
                                )}
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Dialog de Estorno */}
      <Dialog open={reversingEntry !== null} onOpenChange={(open) => !open && closeReverseConfirm()}>
        <DialogContent className="max-w-md rounded-[6px] border border-border bg-surface p-6">
          <DialogHeader>
            <div className="flex items-start gap-3 mb-2">
              <div className="p-2 rounded-[4px] bg-foreground text-background shrink-0">
                <AlertTriangle className="size-4" />
              </div>
              <div>
                <DialogTitle className="text-body-lg font-semibold text-foreground">
                  Confirmar Estorno — Operação #{reversingEntry?.id}
                </DialogTitle>
                <DialogDescription className="text-caption text-muted-foreground mt-0.5">
                  Desfaz a operação "{reversingEntry?.operation ?? "—"}"
                  {reversingEntry?.item_id != null && ` do item #${reversingEntry.item_id}`}. Esta ação é irreversível e será registrada em auditoria.
                </DialogDescription>
              </div>
            </div>
          </DialogHeader>

          <div className="space-y-2">
            <Label className="text-caption font-medium text-foreground">
              Confirme sua senha de acesso para autorizar *
            </Label>
            <Input
              type="password"
              value={password}
              onChange={(e) => {
                setPassword(e.target.value);
                setReverseError(null);
              }}
              placeholder="Senha de acesso ao sistema"
              autoFocus
              className="rounded-[4px] border-border bg-background text-body-sm h-9"
            />
            {reverseError && (
              <p className="text-caption font-medium" style={{ color: "var(--foreground)" }}>
                ✕ {reverseError}
              </p>
            )}
          </div>

          <DialogFooter className="pt-4 border-t border-border flex justify-end gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={closeReverseConfirm}
              className="rounded-[4px] border-border text-xs"
            >
              Cancelar
            </Button>
            <Button
              size="sm"
              disabled={!password || reverseMutation.isPending}
              onClick={handleConfirmReverse}
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs font-medium"
            >
              <RotateCcw className="mr-1.5 size-3.5" />
              {reverseMutation.isPending ? "Estornando…" : "Confirmar Estorno"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
