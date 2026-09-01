import { useEffect, useState } from "react";
import { useMutation, useQuery, useQueryClient, keepPreviousData } from "@tanstack/react-query";
import { Download, Paperclip, RotateCcw } from "lucide-react";

import { listHistoryPaginated, reverseEntryWithPassword, type HistoryEntry } from "@/api/history";
import { downloadAuthenticated } from "@/api/client";
import { DataTable, type Column } from "@/components/app/DataTable";
import { KpiCard } from "@/components/app/KpiCard";
import { PageHeader, Section } from "@/components/app/PageHeader";
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
import { formatCpf, formatDateTime, exportToCsv } from "@/lib/utils";
import { toast } from "sonner";

const REVERSIBLE_OPS = ["Cadastro", "Empréstimo", "Confirmação Empréstimo", "Devolução", "Confirmação Devolução"];
const PAGE_SIZE = 20;

function OperationBadge({ op }: { op?: string | undefined }) {
  if (!op) return <Badge variant="outline">-</Badge>;
  if (op.includes("Empréstimo")) return <Badge className="bg-blue-500/12 text-blue-700 dark:text-blue-400 border-blue-500/25" variant="outline">{op}</Badge>;
  if (op.includes("Devolução")) return <Badge className="bg-emerald-500/12 text-emerald-700 dark:text-emerald-400 border-emerald-500/25" variant="outline">{op}</Badge>;
  if (op === "Exclusão") return <Badge variant="destructive">{op}</Badge>;
  if (op === "Estorno") return <Badge className="bg-amber-500/14 text-amber-700 dark:text-amber-400 border-amber-500/30" variant="outline">{op}</Badge>;
  return <Badge variant="outline">{op}</Badge>;
}

/**
 * Deriva um nome de arquivo legível para download a partir da chave de
 * storage ("categoria/arquivo.ext"): o trecho após a primeira barra já é o
 * nome original enviado pelo usuário.
 */
function attachmentFilename(key: string): string {
  const idx = key.indexOf("/");
  return idx >= 0 ? key.slice(idx + 1) : key;
}

interface AttachmentDescriptor {
  key: string;
  label: string;
}

function AttachmentCell({
  entry,
  downloadingKey,
  onDownload,
}: {
  entry: HistoryEntry;
  downloadingKey: string | null;
  onDownload: (key: string) => void;
}) {
  const attachments: AttachmentDescriptor[] = [];
  if (entry.operacao_anexo) attachments.push({ key: entry.operacao_anexo, label: "Comprovante" });
  if (entry.termo_assinado_anexo) attachments.push({ key: entry.termo_assinado_anexo, label: "Termo" });

  if (attachments.length === 0) return <span className="text-muted-foreground">-</span>;

  return (
    <div className="flex flex-col items-start gap-1">
      {attachments.map(({ key, label }) => (
        <Button
          key={key}
          type="button"
          size="sm"
          variant="ghost"
          className="h-7 px-2 text-muted-foreground hover:text-foreground"
          disabled={downloadingKey === key}
          onClick={() => onDownload(key)}
        >
          <Paperclip className="mr-1 size-3" aria-hidden />
          {downloadingKey === key ? "Baixando..." : label}
        </Button>
      ))}
    </div>
  );
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

  const emprestimos = history.filter((h) => h.operation?.includes("Empréstimo")).length;
  const devolucoes = history.filter((h) => h.operation?.includes("Devolução")).length;
  const estornos = history.filter((h) => h.operation === "Estorno").length;

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
    mutationFn: ({ id, password }: { id: number; password: string }) => reverseEntryWithPassword(id, password),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["history"] });
      queryClient.invalidateQueries({ queryKey: ["items"] });
      toast.success("Operação estornada com sucesso!");
      closeReverseConfirm();
    },
    onError: (err: unknown) => {
      const status = (err as { response?: { status?: number } })?.response?.status;
      const msg = status === 403 ? "Senha incorreta. Ação não autorizada." : getErrorMessage(err, "Erro ao estornar.");
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

  const columns: Column<HistoryEntry>[] = [
    { key: "id", header: "ID", cell: (r) => r.id },
    { key: "item_id", header: "Item", cell: (r) => r.item_id ?? "-" },
    { key: "peripheral_id", header: "Periférico", cell: (r) => r.peripheral_id ?? "-", hideBelow: "lg" },
    { key: "operador", header: "Operador", cell: (r) => r.operador ?? "-", primary: true },
    { key: "operation", header: "Operação", cell: (r) => <OperationBadge op={r.operation} /> },
    { key: "tipo", header: "Tipo", cell: (r) => r.tipo || "-", hideBelow: "md" },
    { key: "marca", header: "Marca", cell: (r) => r.marca || "-", hideBelow: "lg" },
    { key: "modelo", header: "Modelo", cell: (r) => r.modelo || "-", hideBelow: "lg" },
    { key: "identificador", header: "Identificador", cell: (r) => r.identificador || "-", hideBelow: "lg" },
    { key: "nota_fiscal", header: "Nota Fiscal", cell: (r) => r.nota_fiscal || "-", hideBelow: "xl" },
    { key: "usuario", header: "Usuário", cell: (r) => r.usuario || "-", hideBelow: "md" },
    { key: "cpf", header: "CPF", cell: (r) => formatCpf(r.cpf), hideBelow: "xl" },
    { key: "cargo", header: "Cargo", cell: (r) => r.cargo || "-", hideBelow: "xl" },
    { key: "setor", header: "Setor", cell: (r) => r.setor || "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (r) => r.revenda || "-", hideBelow: "md" },
    { key: "data_operacao", header: "Data", cell: (r) => formatDateTime(r.data_operacao) },
    { key: "details", header: "Detalhes", cell: (r) => r.details || "-", hideBelow: "xl" },
    {
      key: "anexo",
      header: "Anexo",
      cell: (r) => <AttachmentCell entry={r} downloadingKey={downloadingKey} onDownload={handleDownloadAttachment} />,
    },
    ...(hasRole("Gestor")
      ? [
          {
            key: "actions",
            header: "",
            hideOnMobile: false,
            cell: (r: HistoryEntry) =>
              REVERSIBLE_OPS.includes(r.operation ?? "") ? (
                <Button
                  size="sm"
                  variant="ghost"
                  className="text-destructive hover:text-destructive/80"
                  onClick={() => openReverseConfirm(r)}
                >
                  <RotateCcw className="mr-1 size-3.5" aria-hidden />
                  Estornar
                </Button>
              ) : null,
          } satisfies Column<HistoryEntry>,
        ]
      : []),
  ];

  return (
    <div className="space-y-6">
      <PageHeader
        eyebrow="Auditoria"
        title="Histórico de Ações"
        description="Registro completo de todas as operações do sistema"
        actions={
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
          >
            <Download className="mr-2 size-4" aria-hidden />
            Exportar CSV
          </Button>
        }
      />

      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
        <KpiCard label="Total de Registros" value={total} />
        <KpiCard label="Empréstimos (página)" value={emprestimos} />
        <KpiCard label="Devoluções (página)" value={devolucoes} />
        <KpiCard label="Estornos (página)" value={estornos} />
      </div>

      <Section
        title="Registros"
        actions={
          <Input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Buscar por operador, usuário, operação..."
            className="w-64"
          />
        }
      >
        <DataTable
          data={history}
          columns={columns}
          rowKey={(r) => r.id}
          isLoading={isLoading}
          error={error}
          onRetry={() => void refetch()}
          emptyTitle="Nenhum registro encontrado"
          emptyDescription="Ajuste a busca para ver outros resultados."
          pagination={{
            page: pageIndex,
            pageSize: PAGE_SIZE,
            total,
            onPageChange: setPageIndex,
          }}
        />
      </Section>

      <Dialog open={reversingEntry !== null} onOpenChange={(open) => !open && closeReverseConfirm()}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Confirmar Estorno — Operação #{reversingEntry?.id}</DialogTitle>
            <DialogDescription>
              Isso desfará a operação &quot;{reversingEntry?.operation ?? "-"}&quot;
              {reversingEntry?.item_id != null && ` do item #${reversingEntry.item_id}`}
              {reversingEntry?.peripheral_id != null && ` do periférico #${reversingEntry.peripheral_id}`}
              {reversingEntry?.operador && ` (operador: ${reversingEntry.operador})`}. Esta ação não pode ser
              desfeita.
            </DialogDescription>
          </DialogHeader>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="senha-estorno">Confirme sua senha</Label>
            <Input
              id="senha-estorno"
              type="password"
              value={password}
              onChange={(e) => {
                setPassword(e.target.value);
                setReverseError(null);
              }}
              placeholder="Sua senha de acesso"
              autoFocus
            />
            {reverseError && <p className="text-sm text-destructive">{reverseError}</p>}
          </div>
          <DialogFooter>
            <Button variant="ghost" onClick={closeReverseConfirm}>
              Cancelar
            </Button>
            <Button variant="destructive" disabled={!password || reverseMutation.isPending} onClick={handleConfirmReverse}>
              {reverseMutation.isPending ? "Estornando..." : "Confirmar"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
