import { createFileRoute, Outlet, useNavigate } from "@tanstack/react-router";
import { useEffect } from "react";

import { AppShell } from "@/components/app/AppShell";
import { LoadingState } from "@/components/app/StateBlocks";
import { useAuth } from "@/lib/auth";

// A sessão vive apenas no cliente (access token em memória + cookie httpOnly de
// refresh), então o servidor nunca consegue decidir o gate — renderizamos só no
// navegador para evitar flash de conteúdo e loops de redirecionamento.
export const Route = createFileRoute("/_shell")({
  ssr: false,
  component: ShellLayout,
});

function ShellLayout() {
  const { user, isLoading } = useAuth();
  const navigate = useNavigate();

  useEffect(() => {
    if (!isLoading && !user) void navigate({ to: "/login", replace: true });
  }, [isLoading, user, navigate]);

  if (isLoading || !user) {
    return (
      <div className="grid min-h-screen place-items-center bg-background">
        <LoadingState label="Verificando sessão…" />
      </div>
    );
  }

  return (
    <AppShell>
      <Outlet />
    </AppShell>
  );
}
