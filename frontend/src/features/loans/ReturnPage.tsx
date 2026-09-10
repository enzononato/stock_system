import { useMemo, useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import {
  CheckCircle2,
  FileDown,
  Search,
  ClipboardCheck,
  CheckSquare,
  ShieldCheck,
  PackageCheck,
} from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { confirmReturn, downloadReturnTerm } from "@/api/loans";
import { getErrorMessage } from "@/lib/api-error";
import { formatDate } from "@/lib/utils";
import { FileUpload } from "@/components/app/FileUpload";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

const FETCH_ALL_LIMIT = 500;

export function ReturnPage() {
  const queryClient = useQueryClient();
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null);
  const [signedPdf, setSignedPdf] = useState<File | null>(null);
  const [searchActive, setSearchActive] = useState("");
  const [pendentesPage, setPendentesPage] = useState(0);
  const [indisponivelPage, setIndisponivelPage] = useState(0);
  const PAGE_SIZE = 7;

  // Physical checklist state during return confirmation
  const [checkChassis, setCheckChassis] = useState(false);
  const [checkScreenPower, setCheckScreenPower] = useState(false);
  const [checkPeripherals, setCheckPeripherals] = useState(false);

  const { data, isLoading, refetch } = useQuery({
    queryKey: ["items"],
    queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }),
  });
  const items = data?.items ?? [];
  const indisponivel = items.filter((i) => i.status === "Indisponível" && Boolean(i.assigned_to));
  const pendenteDevolucao = items.filter((i) => i.status === "Pendente Devolução");

  const filteredIndisponivel = useMemo(() => {
    if (!searchActive.trim()) return indisponivel;
    const q = searchActive.toLowerCase();
    return indisponivel.filter(
      (i) =>
        String(i.id).includes(q) ||
        (i.tipo ?? "").toLowerCase().includes(q) ||
        (i.brand ?? "").toLowerCase().includes(q) ||
        (i.assigned_to ?? "").toLowerCase().includes(q) ||
        (i.cpf ?? "").includes(q) ||
        (i.revenda ?? "").toLowerCase().includes(q),
    );
  }, [indisponivel, searchActive]);

  const initiateMutation = useMutation({
    mutationFn: (itemId: number) => downloadReturnTerm(itemId),
    onSuccess: (_, itemId) => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      setPendingReturnId(itemId);
      setCheckChassis(false);
      setCheckScreenPower(false);
      setCheckPeripherals(false);
      toast.success("Termo de devolução gerado! Realize a inspeção física e anexe o documento.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao gerar termo de devolução.")),
  });

  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmReturn(itemId, pdf),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["items"] });
      setPendingReturnId(null);
      setSignedPdf(null);
      setCheckChassis(false);
      setCheckScreenPower(false);
      setCheckPeripherals(false);
      toast.success("Devolução concluída com sucesso! Item retornado ao estoque disponível.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao confirmar devolução.")),
  });

  const activeSelectedItem = items.find((i) => i.id === pendingReturnId);

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Operações de Campo
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Conferência Física</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Devolução de Equipamentos
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Recepção patrimonial com checklist de integridade, laudo de avarias e arquivamento de termo.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-[4px] border border-border bg-surface text-caption">
            <PackageCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Em Campo:</span>
            <span className="font-mono font-semibold text-foreground">{indisponivel.length}</span>
          </div>
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-[4px] border border-border bg-surface text-caption">
            <ClipboardCheck className="size-3.5 text-muted-foreground" />
            <span className="text-muted-foreground">Aguardando Confirmação:</span>
            <span className="font-mono font-semibold text-foreground">{pendenteDevolucao.length}</span>
          </div>
        </div>
      </div>

      {/* Modal / Workflow de Conferência Física */}
      <Dialog
        open={pendingReturnId !== null}
        onOpenChange={(open) => {
          if (!open) {
            setPendingReturnId(null);
            setSignedPdf(null);
          }
        }}
      >
        <DialogContent className="max-w-lg rounded-[6px] border border-border bg-surface p-6">
          <DialogHeader>
            <DialogTitle className="text-body-lg font-semibold text-foreground">
              Conferência Física & Termo de Devolução
            </DialogTitle>
            <DialogDescription className="text-caption text-muted-foreground">
              Verifique o estado físico do equipamento #{pendingReturnId} antes de reintegrar ao estoque disponível.
            </DialogDescription>
          </DialogHeader>

          {activeSelectedItem && (
            <div className="rounded-[4px] border border-border bg-surface-alt p-3 text-caption space-y-1">
              <div className="flex justify-between">
                <span className="text-muted-foreground">Equipamento:</span>
                <span className="font-medium text-foreground">
                  {activeSelectedItem.tipo} {activeSelectedItem.brand} {activeSelectedItem.model}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Responsável atual:</span>
                <span className="font-medium text-foreground">{activeSelectedItem.assigned_to}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Revenda:</span>
                <span className="text-foreground">{activeSelectedItem.revenda || "—"}</span>
              </div>
            </div>
          )}

          {/* Checklist Físico */}
          <div className="space-y-3 pt-2 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              Checklist de Inspeção Física
            </span>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox
                checked={checkChassis}
                onCheckedChange={(c) => setCheckChassis(Boolean(c))}
                className="rounded-[2px] border-border"
              />
              <span>Integridade do chassi / gabinete (sem trincas, amassados ou danos estruturais)</span>
            </label>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox
                checked={checkScreenPower}
                onCheckedChange={(c) => setCheckScreenPower(Boolean(c))}
                className="rounded-[2px] border-border"
              />
              <span>Tela, teclado e fonte de alimentação testados e em pleno funcionamento</span>
            </label>

            <label className="flex items-center gap-2.5 cursor-pointer text-caption text-foreground select-none">
              <Checkbox
                checked={checkPeripherals}
                onCheckedChange={(c) => setCheckPeripherals(Boolean(c))}
                className="rounded-[2px] border-border"
              />
              <span>Periféricos e cabos complementares entregues e conferidos</span>
            </label>
          </div>

          {/* Upload de Termo Assinado */}
          <div className="space-y-2 pt-3 border-t border-border">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground block">
              Termo de Devolução Assinado (PDF) *
            </span>
            <FileUpload
              accept="application/pdf"
              onFile={setSignedPdf}
              label="Arraste ou clique para anexar o termo assinado (PDF)"
            />
          </div>

          <DialogFooter className="pt-3 border-t border-border flex justify-end gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => setPendingReturnId(null)}
              className="rounded-[4px] border-border"
            >
              Cancelar
            </Button>
            <Button
              size="sm"
              disabled={
                confirmMutation.isPending ||
                !signedPdf ||
                !checkChassis ||
                !checkScreenPower ||
                !checkPeripherals
              }
              onClick={() => {
                if (pendingReturnId && signedPdf) {
                  confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf });
                }
              }}
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 font-medium"
            >
              <CheckCircle2 className="mr-1.5 size-3.5" />
              {confirmMutation.isPending ? "Concluindo…" : "Finalizar Devolução"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Seção 1: Pendentes de Confirmação de Devolução */}
      {pendenteDevolucao.length > 0 && (
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-3">
          <div className="flex items-center justify-between border-b border-border pb-2.5">
            <div>
              <h2 className="text-body font-semibold text-foreground">
                Aguardando Confirmação e Termo ({pendenteDevolucao.length})
              </h2>
              <p className="text-caption text-muted-foreground">
                Equipamentos cujo termo de devolução já foi impresso e aguardam upload da via assinada.
              </p>
            </div>
            <Badge variant="outline" className="rounded-[2px] font-mono text-[11px] border-border">
              ! RECOLHIMENTO
            </Badge>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3">ID</th>
                  <th className="py-2 px-3">Equipamento</th>
                  <th className="py-2 px-3">Colaborador</th>
                  <th className="py-2 px-3">CPF</th>
                  <th className="py-2 px-3">Revenda</th>
                  <th className="py-2 px-3 text-right">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {pendenteDevolucao
                  .slice(pendentesPage * PAGE_SIZE, (pendentesPage + 1) * PAGE_SIZE)
                  .map((item) => (
                    <tr key={item.id} className="hover:bg-muted/30">
                      <td className="py-2 px-3 font-mono font-medium">#{item.id}</td>
                      <td className="py-2 px-3">{item.tipo} {item.brand} {item.model}</td>
                      <td className="py-2 px-3 font-medium">{item.assigned_to}</td>
                      <td className="py-2 px-3 font-mono text-muted-foreground">{item.cpf || "—"}</td>
                      <td className="py-2 px-3 text-muted-foreground">{item.revenda || "—"}</td>
                      <td className="py-2 px-3 text-right">
                        <Button
                          size="sm"
                          onClick={() => setPendingReturnId(item.id)}
                          className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-7 px-2.5"
                        >
                          Confirmar Devolução
                        </Button>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>

          {Math.ceil(pendenteDevolucao.length / PAGE_SIZE) > 1 && (
            <div className="flex items-center justify-between border-t border-border pt-3">
              <span className="text-caption text-muted-foreground">
                Página {pendentesPage + 1} de {Math.ceil(pendenteDevolucao.length / PAGE_SIZE)} ({pendenteDevolucao.length} pendentes)
              </span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage === 0}
                  onClick={() => setPendentesPage((p) => Math.max(0, p - 1))}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  ← Anterior
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  disabled={pendentesPage >= Math.ceil(pendenteDevolucao.length / PAGE_SIZE) - 1}
                  onClick={() => setPendentesPage((p) => p + 1)}
                  className="rounded-[4px] border-border text-xs h-7 px-2.5"
                >
                  Próxima →
                </Button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Seção 2: Lista Operacional de Equipamentos em Campo */}
      <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-border pb-3">
          <div>
            <h2 className="text-body font-semibold text-foreground">
              Equipamentos Alocados em Campo ({filteredIndisponivel.length})
            </h2>
            <p className="text-caption text-muted-foreground">
              Selecione um ativo para iniciar o procedimento de devolução e emitir o documento formal.
            </p>
          </div>

          <div className="relative w-full sm:w-72">
            <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 size-3.5 text-muted-foreground" />
            <Input
              value={searchActive}
              onChange={(e) => {
                setSearchActive(e.target.value);
                setIndisponivelPage(0);
              }}
              placeholder="Buscar por ID, modelo, CPF, colaborador…"
              className="pl-8 h-8 rounded-[4px] border-border bg-background text-body-sm"
            />
          </div>
        </div>

        {filteredIndisponivel.length === 0 ? (
          <div className="py-8 text-center text-caption text-muted-foreground">
            Nenhum equipamento em campo encontrado para os critérios de busca.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-body-sm">
              <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <tr>
                  <th className="py-2 px-3">ID</th>
                  <th className="py-2 px-3">Tipo / Modelo</th>
                  <th className="py-2 px-3">Colaborador</th>
                  <th className="py-2 px-3">CPF</th>
                  <th className="py-2 px-3">Revenda</th>
                  <th className="py-2 px-3">Data Empréstimo</th>
                  <th className="py-2 px-3 text-right">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-border">
                {filteredIndisponivel
                  .slice(indisponivelPage * PAGE_SIZE, (indisponivelPage + 1) * PAGE_SIZE)
                  .map((item) => (
                    <tr key={item.id} className="hover:bg-muted/30">
                      <td className="py-2 px-3 font-mono font-medium">#{item.id}</td>
                      <td className="py-2 px-3">
                        <span className="font-medium text-foreground">{item.tipo}</span>{" "}
                        <span className="text-muted-foreground">{item.brand} {item.model}</span>
                      </td>
                      <td className="py-2 px-3 font-medium">{item.assigned_to}</td>
                      <td className="py-2 px-3 font-mono text-muted-foreground">{item.cpf || "—"}</td>
                      <td className="py-2 px-3 text-muted-foreground">{item.revenda || "—"}</td>
                      <td className="py-2 px-3 font-mono text-muted-foreground">{formatDate(item.date_issued)}</td>
                      <td className="py-2 px-3 text-right">
                        <Button
                          size="sm"
                          variant="outline"
                          disabled={initiateMutation.isPending}
                          onClick={() => initiateMutation.mutate(item.id)}
                          className="rounded-[4px] border-border text-xs h-7 px-2.5 hover:bg-muted"
                        >
                          <FileDown className="mr-1 size-3" />
                          Gerar Termo
                        </Button>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        )}

        {Math.ceil(filteredIndisponivel.length / PAGE_SIZE) > 1 && (
          <div className="flex items-center justify-between border-t border-border pt-3">
            <span className="text-caption text-muted-foreground">
              Página {indisponivelPage + 1} de {Math.ceil(filteredIndisponivel.length / PAGE_SIZE)} ({filteredIndisponivel.length} em campo)
            </span>
            <div className="flex items-center gap-2">
              <Button
                size="sm"
                variant="outline"
                disabled={indisponivelPage === 0}
                onClick={() => setIndisponivelPage((p) => Math.max(0, p - 1))}
                className="rounded-[4px] border-border text-xs h-7 px-2.5"
              >
                ← Anterior
              </Button>
              <Button
                size="sm"
                variant="outline"
                disabled={indisponivelPage >= Math.ceil(filteredIndisponivel.length / PAGE_SIZE) - 1}
                onClick={() => setIndisponivelPage((p) => p + 1)}
                className="rounded-[4px] border-border text-xs h-7 px-2.5"
              >
                Próxima →
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
