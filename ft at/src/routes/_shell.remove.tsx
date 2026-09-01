import { createFileRoute } from "@tanstack/react-router";

import { RemovePage } from "@/features/items/RemovePage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/remove")({
  head: () => ({
    meta: [
      { title: "Remover item · Controle de Estoque de TI" },
      { name: "description", content: "Remover item no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Remover item · Controle de Estoque de TI" },
      { property: "og:description", content: "Remover item no sistema corporativo de controle de estoque e patrimônio de TI." },
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
        <RemovePage />
      </RequireRole>
    </>
  );
}
