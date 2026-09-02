import { useMemo, useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { CheckCircle2, FileDown, Search } from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { confirmReturn, downloadReturnTerm } from "@/api/loans";
import { getErrorMessage } from "@/lib/api-error";
import { formatDate } from "@/lib/utils";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { DataTable, type Column } from "@/components/app/DataTable";
import { FileUpload } from "@/components/app/FileUpload";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

const FETCH_ALL_LIMIT = 500;

export function ReturnPage() {
  const queryClient = useQueryClient();
  const [pendingReturnId, setPendingReturnId] = useState<number | null>(null);
  const [signedPdf, setSignedPdf] = useState<File | null>(null);
  const [searchActive, setSearchActive] = useState("");

  const { data, isLoading, error, refetch } = useQuery({ queryKey: ["items"], queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }) });
  const items = data?.items ?? [];
  const indisponivel = items.filter((i) => i.status === "Indisponível" && Boolean(i.assigned_to));
  const pendenteDevolucao = items.filter((i) => i.status === "Pendente Devolução");

  const filteredIndisponivel = useMemo(() => {
    if (!searchActive.trim()) return indisponivel;
    const q = searchActive.toLowerCase();
    return indisponivel.filter((i) =>
      String(i.id).includes(q) ||
      (i.tipo ?? "").toLowerCase().includes(q) ||
      (i.brand ?? "").toLowerCase().includes(q) ||
      (i.assigned_to ?? "").toLowerCase().includes(q) ||
      (i.cpf ?? "").includes(q) ||
      (i.revenda ?? "").toLowerCase().includes(q)
    );
  }, [indisponivel, searchActive]);

  const initiateMutation = useMutation({
    mutationFn: (itemId: number) => downloadReturnTerm(itemId),
    onSuccess: (_, itemId) => { queryClient.invalidateQueries({ queryKey: ["items"] }); setPendingReturnId(itemId); toast.success("Termo de devolução gerado! Faça a assinatura e confirme."); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao gerar termo.")),
  });
  const confirmMutation = useMutation({
    mutationFn: ({ itemId, pdf }: { itemId: number; pdf: File }) => confirmReturn(itemId, pdf),
    onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["items"] }); setPendingReturnId(null); setSignedPdf(null); toast.success("Devolução confirmada com sucesso!"); },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao confirmar devolução.")),
  });

  const activeColumns: Column<Item>[] = [
    { key: "id", header: "ID", cell: (i) => `#${i.id}`, primary: true },
    { key: "tipo", header: "Tipo", cell: (i) => i.tipo ?? "-" },
    { key: "brand", header: "Marca", cell: (i) => i.brand ?? "-" },
    { key: "assigned_to", header: "Usuário", cell: (i) => i.assigned_to ?? "-" },
    { key: "cpf", header: "CPF", cell: (i) => i.cpf ?? "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (i) => i.revenda ?? "-", hideBelow: "md" },
    { key: "date_issued", header: "Empréstimo", cell: (i) => formatDate(i.date_issued), hideBelow: "lg" },
    { key: "actions", header: "", hideOnMobile: true, cell: (i) => <Button size="sm" variant="outline" disabled={initiateMutation.isPending} onClick={() => initiateMutation.mutate(i.id)}><FileDown className="mr-1 size-3.5" aria-hidden />Gerar termo</Button> },
  ];

  const pendingColumns: Column<Item>[] = [
    { key: "id", header: "ID", cell: (i) => `#${i.id}`, primary: true },
    { key: "tipo", header: "Tipo", cell: (i) => i.tipo ?? "-" },
    { key: "brand", header: "Marca", cell: (i) => i.brand ?? "-" },
    { key: "assigned_to", header: "Usuário", cell: (i) => i.assigned_to ?? "-" },
    { key: "cpf", header: "CPF", cell: (i) => i.cpf ?? "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (i) => i.revenda ?? "-", hideBelow: "md" },
    { key: "date_issued", header: "Empréstimo", cell: (i) => formatDate(i.date_issued), hideBelow: "lg" },
    { key: "actions", header: "", hideOnMobile: true, cell: (i) => <Button size="sm" variant="outline" className="text-primary hover:text-primary" onClick={() => setPendingReturnId(i.id)}>Confirmar</Button> },
  ];

  return (
    <div className="space-y-8">
      <PageHeader title="Devolver equipamento" description="Gerencie as devoluções de equipamentos emprestados." />

      <div className="grid items-stretch gap-6 xl:grid-cols-2 xl:auto-rows-fr">
        <Section
          title={`Pendentes de confirmação (${pendenteDevolucao.length})`}
          description="Termos gerados que aguardam o PDF assinado."
          className="flex flex-col h-full min-h-[460px]"
        >
          <DataTable
            data={pendenteDevolucao}
            columns={pendingColumns}
            rowKey={(i) => i.id}
            isLoading={isLoading}
            error={error}
            clientPageSize={7}
            onRetry={() => void refetch()}
            emptyTitle="Nenhuma devolução pendente"
            emptyDescription="As devoluções aguardando assinatura aparecerão aqui."
          />
        </Section>

        <Section
          title={`Empréstimos ativos (${indisponivel.length})`}
          description="Selecione um item para gerar o termo de devolução."
          className="flex flex-col h-full min-h-[460px]"
          actions={
            <div className="relative w-44 sm:w-56">
              <Search className="absolute left-2.5 top-2.5 size-3.5 text-muted-foreground" aria-hidden />
              <Input
                placeholder="Buscar usuário/patrimônio…"
                value={searchActive}
                onChange={(e) => setSearchActive(e.target.value)}
                className="h-8 pl-8 text-xs"
              />
            </div>
          }
        >
          <DataTable
            data={filteredIndisponivel}
            columns={activeColumns}
            rowKey={(i) => i.id}
            isLoading={isLoading}
            error={error}
            clientPageSize={7}
            onRetry={() => void refetch()}
            emptyTitle="Nenhum empréstimo ativo"
            emptyDescription={searchActive ? "Nenhum registro corresponde à busca." : undefined}
          />
        </Section>
      </div>

      {pendingReturnId && (
        <Section title={`Confirmar devolução — item #${pendingReturnId}`} className="border-warning/30 bg-warning/5">
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground">O termo de devolução foi gerado. Faça o upload do PDF assinado para confirmar.</p>
            <FileUpload onFile={setSignedPdf} label="Upload do termo de devolução assinado (PDF)" />
            <div className="flex gap-3">
              <Button disabled={!signedPdf || confirmMutation.isPending} onClick={() => signedPdf && confirmMutation.mutate({ itemId: pendingReturnId, pdf: signedPdf })}>
                <CheckCircle2 className="mr-2 size-3.5" aria-hidden />
                {confirmMutation.isPending ? "Confirmando…" : "Confirmar devolução"}
              </Button>
              <Button variant="ghost" onClick={() => setPendingReturnId(null)}>
                Cancelar
              </Button>
            </div>
          </div>
        </Section>
      )}
    </div>
  );
}

