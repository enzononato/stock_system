import { createFileRoute } from "@tanstack/react-router";

import { LoanPage } from "@/features/loans/LoanPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/loan")({
  head: () => ({
    meta: [
      { title: "Empréstimo · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Empréstimo no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Empréstimo · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Empréstimo no sistema corporativo de controle de estoque e patrimônio de TI.",
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
        <LoanPage />
      </RequireRole>
    </>
  );
}
