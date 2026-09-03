import { createFileRoute } from "@tanstack/react-router";

import { TermsPage } from "@/features/loans/TermsPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/terms")({
  head: () => ({
    meta: [
      { title: "Termos · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Termos no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Termos · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Termos no sistema corporativo de controle de estoque e patrimônio de TI.",
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
        <TermsPage />
      </RequireRole>
    </>
  );
}
