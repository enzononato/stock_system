import { createFileRoute } from "@tanstack/react-router";

import { PeripheralsPage } from "@/features/peripherals/PeripheralsPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/peripherals")({
  head: () => ({
    meta: [
      { title: "Periféricos · Controle de Estoque de TI" },
      { name: "description", content: "Periféricos no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Periféricos · Controle de Estoque de TI" },
      { property: "og:description", content: "Periféricos no sistema corporativo de controle de estoque e patrimônio de TI." },
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
        <PeripheralsPage />
      </RequireRole>
    </>
  );
}
