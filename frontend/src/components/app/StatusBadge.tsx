import { Badge } from "@/components/ui/badge";
import { cn } from "@/lib/utils";

/**
 * Mapeia os status vindos do backend (`disponivel`, `emprestado`, `manutencao`,
 * `aguardando_assinatura`, `removido`…) para cores consistentes. Status
 * desconhecido cai no visual neutro em vez de sumir da tela.
 */
const STATUS_STYLE: Record<string, string> = {
  disponivel: "badge-success",
  disponível: "badge-success",
  emprestado: "badge-info",
  aguardando_assinatura: "badge-warning",
  aguardando_devolucao: "badge-warning",
  manutencao: "badge-warning",
  inativo: "badge-muted",
  removido: "badge-destructive",
};

export function statusLabel(status?: string | null): string {
  if (!status) return "—";
  return status.replace(/_/g, " ").replace(/^\w/, (c) => c.toUpperCase());
}

export function StatusBadge({
  status,
  className,
}: {
  status?: string | null | undefined;
  className?: string | undefined;
}) {
  const key = (status ?? "").toLowerCase();
  return (
    <Badge
      variant="outline"
      className={cn(
        "font-medium text-xs tracking-tight",
        STATUS_STYLE[key] ?? "badge-muted",
        className,
      )}
    >
      {statusLabel(status)}
    </Badge>
  );
}
