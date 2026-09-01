import type { ReactNode } from "react";
import { Link } from "@tanstack/react-router";
import { ShieldAlert } from "lucide-react";

import { useAuth } from "@/lib/auth";
import { Button } from "@/components/ui/button";
import { EmptyState } from "@/components/app/StateBlocks";

/**
 * Reflete no frontend as permissões já aplicadas pelo backend. A segurança real
 * continua sendo do servidor: aqui apenas evitamos mostrar telas que o usuário
 * não pode operar.
 */
export function RequireRole({ roles, children }: { roles: string[]; children: ReactNode }) {
  const { user } = useAuth();

  if (!user || !roles.includes(user.role)) {
    return (
      <EmptyState
        icon={<ShieldAlert className="size-5" />}
        title="Acesso restrito"
        description={`Esta área é exclusiva para: ${roles.join(", ")}. Seu perfil atual é ${user?.role ?? "desconhecido"}.`}
        action={
          <Button asChild variant="outline" size="sm">
            <Link to="/">Voltar ao estoque</Link>
          </Button>
        }
      />
    );
  }

  return <>{children}</>;
}
