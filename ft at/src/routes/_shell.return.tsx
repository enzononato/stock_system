import { createFileRoute } from "@tanstack/react-router";

import { ReturnPage } from "@/features/loans/ReturnPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/return")({
  head: () => ({
    meta: [
      { title: "Devolução · Controle de Estoque de TI" },
      { name: "description", content: "Devolução no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Devolução · Controle de Estoque de TI" },
      { property: "og:description", content: "Devolução no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: RouteComponent,
});

function RouteComponent() {
  return (
    <>
      <RequireRole roles={["Gestor", "Técnico"]}>
        <ReturnPage />
      </RequireRole>
    </>
  );
}
