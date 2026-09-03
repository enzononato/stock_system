import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { toast } from "sonner";
import { CheckCircle2, FileDown } from "lucide-react";

import { listItemsPaginated, type Item } from "@/api/items";
import { downloadSignedTerm } from "@/api/loans";
import { getErrorMessage } from "@/lib/api-error";
import { formatDate } from "@/lib/utils";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { DataTable, type Column } from "@/components/app/DataTable";
import { ConfirmacaoTermo, generateAndDownloadLoanTerm } from "@/components/app/ConfirmacaoTermo";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";

const FETCH_ALL_LIMIT = 500;

export function TermsPage() {
  const [confirmingId, setConfirmingId] = useState<number | null>(null);
  const { data, isLoading, error, refetch } = useQuery({ queryKey: ["items"], queryFn: () => listItemsPaginated({ limit: FETCH_ALL_LIMIT }) });
  const items = data?.items ?? [];
  const pendentes = items.filter((i) => i.status === "Pendente");
  const ativos = items.filter((i) => i.status === "Indisponível" && Boolean(i.assigned_to));

  async function handleViewSignedTerm(itemId: number) {
    try {
      const blob = await downloadSignedTerm(itemId);
      const url = URL.createObjectURL(blob);
      window.open(url, "_blank");
      window.setTimeout(() => URL.revokeObjectURL(url), 10000);
    } catch (err) {
      toast.error(getErrorMessage(err, "Termo assinado não encontrado."));
    }
  }

  const pendingColumns: Column<Item>[] = [
    { key: "id", header: "ID", cell: (i) => `#${i.id}`, primary: true },
    { key: "tipo", header: "Tipo", cell: (i) => i.tipo ?? "-" },
    { key: "brand", header: "Marca", cell: (i) => i.brand ?? "-", hideBelow: "md" },
    { key: "model", header: "Modelo", cell: (i) => i.model ?? "-", hideBelow: "lg" },
    { key: "assigned_to", header: "Usuário", cell: (i) => i.assigned_to ?? "-" },
    { key: "cpf", header: "CPF", cell: (i) => i.cpf ?? "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (i) => i.revenda ?? "-", hideBelow: "md" },
    { key: "date_issued", header: "Data", cell: (i) => formatDate(i.date_issued), hideBelow: "lg" },
    { key: "actions", header: "Ações", hideOnMobile: true, cell: (i) => <div className="flex gap-2"><Button size="sm" variant="outline" onClick={() => generateAndDownloadLoanTerm(i.id)}><FileDown className="mr-1 size-3.5" aria-hidden />Termo</Button><Button size="sm" onClick={() => setConfirmingId(i.id)}><CheckCircle2 className="mr-1 size-3.5" aria-hidden />Confirmar</Button></div> },
  ];
  const activeColumns: Column<Item>[] = [
    { key: "id", header: "ID", cell: (i) => `#${i.id}`, primary: true },
    { key: "tipo", header: "Tipo", cell: (i) => i.tipo ?? "-" },
    { key: "assigned_to", header: "Usuário", cell: (i) => i.assigned_to ?? "-" },
    { key: "cpf", header: "CPF", cell: (i) => i.cpf ?? "-", hideBelow: "lg" },
    { key: "revenda", header: "Revenda", cell: (i) => i.revenda ?? "-", hideBelow: "md" },
    { key: "date_issued", header: "Data empréstimo", cell: (i) => formatDate(i.date_issued), hideBelow: "lg" },
    { key: "actions", header: "Termo", hideOnMobile: true, cell: (i) => <Button size="sm" variant="outline" onClick={() => handleViewSignedTerm(i.id)}><FileDown className="mr-1 size-3.5" aria-hidden />Ver termo</Button> },
  ];

  return <div className="space-y-8">
    <PageHeader title="Termos de responsabilidade" description="Gerencie os termos de empréstimo pendentes e confirmados." />
    {confirmingId && <ConfirmacaoTermo itemId={confirmingId} description="Faça o upload do termo de responsabilidade assinado (PDF)." uploadLabel="Arraste ou clique para enviar o PDF assinado" errorMessage="Erro ao confirmar." onConfirmed={() => setConfirmingId(null)} onCancel={() => setConfirmingId(null)} />}
    <div className="grid grid-cols-1 items-stretch gap-5 xl:grid-cols-2 xl:auto-rows-fr">
      <Section title="Pendentes de confirmação" actions={pendentes.length > 0 ? <Badge variant="outline">{pendentes.length}</Badge> : undefined} description="Gere o termo, colete a assinatura e confirme o empréstimo com o PDF assinado." className="h-full min-h-[430px]">
        <DataTable data={pendentes} columns={pendingColumns} rowKey={(i) => i.id} isLoading={isLoading} error={error} onRetry={() => void refetch()} clientPageSize={7} emptyTitle="Nenhum empréstimo pendente de confirmação" />
      </Section>
      <Section title={`Empréstimos ativos (${ativos.length})`} description="Consulte os empréstimos que já foram confirmados e acesse o termo assinado." className="h-full min-h-[430px]">
        <DataTable data={ativos} columns={activeColumns} rowKey={(i) => i.id} isLoading={isLoading} error={error} onRetry={() => void refetch()} clientPageSize={7} emptyTitle="Nenhum empréstimo ativo no momento" />
      </Section>
    </div>
  </div>;
}
