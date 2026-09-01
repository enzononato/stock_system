import { createFileRoute } from "@tanstack/react-router";

import { HistoryPage } from "@/features/history/HistoryPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/history")({
  head: () => ({
    meta: [
      { title: "Histórico · Controle de Estoque de TI" },
      { name: "description", content: "Histórico no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Histórico · Controle de Estoque de TI" },
      { property: "og:description", content: "Histórico no sistema corporativo de controle de estoque e patrimônio de TI." },
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
        <HistoryPage />
      </RequireRole>
    </>
  );
}
