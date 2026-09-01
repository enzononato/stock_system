import { createFileRoute } from "@tanstack/react-router";

import { UnidadesPage } from "@/features/admin/UnidadesPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/unidades")({
  head: () => ({
    meta: [
      { title: "Unidades · Controle de Estoque de TI" },
      { name: "description", content: "Unidades no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Unidades · Controle de Estoque de TI" },
      { property: "og:description", content: "Unidades no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: RouteComponent,
});

function RouteComponent() {
  return (
    <>
      <RequireRole roles={["Gestor"]}>
        <UnidadesPage />
      </RequireRole>
    </>
  );
}
