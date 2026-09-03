import { createFileRoute } from "@tanstack/react-router";

import { ReportPage } from "@/features/reports/ReportPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/report")({
  head: () => ({
    meta: [
      { title: "Relatório mensal · Controle de Estoque de TI" },
      {
        name: "description",
        content:
          "Relatório mensal no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Relatório mensal · Controle de Estoque de TI" },
      {
        property: "og:description",
        content:
          "Relatório mensal no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
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
        <ReportPage />
      </RequireRole>
    </>
  );
}
