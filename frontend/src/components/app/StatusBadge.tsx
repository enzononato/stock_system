import { Badge } from "@/components/ui/badge";
import { cn } from "@/lib/utils";

/**
 * Mapeia os status vindos do backend (`disponivel`, `emprestado`, `manutencao`,
 * `aguardando_assinatura`, `removido`…) para cores consistentes. Status
 * desconhecido cai no visual neutro em vez de sumir da tela.
 */
const STATUS_STYLE: Record<string, string> = {
  disponivel: "bg-emerald-500/12 text-emerald-700 dark:text-emerald-400 border-emerald-500/25",
  disponível: "bg-emerald-500/12 text-emerald-700 dark:text-emerald-400 border-emerald-500/25",
  emprestado: "bg-blue-500/12 text-blue-700 dark:text-blue-400 border-blue-500/25",
  aguardando_assinatura: "bg-amber-500/14 text-amber-700 dark:text-amber-400 border-amber-500/30",
  aguardando_devolucao: "bg-amber-500/14 text-amber-700 dark:text-amber-400 border-amber-500/30",
  manutencao: "bg-purple-500/12 text-purple-700 dark:text-purple-400 border-purple-500/25",
  inativo: "bg-muted text-muted-foreground border-border",
  removido: "bg-destructive/12 text-destructive border-destructive/25",
};

export function statusLabel(status?: string | null): string {
  if (!status) return "—";
  return status
    .replace(/_/g, " ")
    .replace(/^\w/, (c) => c.toUpperCase());
}

export function StatusBadge({ status, className }: { status?: string | null | undefined; className?: string | undefined }) {
  const key = (status ?? "").toLowerCase();
  return (
    <Badge
      variant="outline"
      className={cn("font-medium", STATUS_STYLE[key] ?? "bg-muted text-muted-foreground border-border", className)}
    >
      {statusLabel(status)}
    </Badge>
  );
}
