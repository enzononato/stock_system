import { createFileRoute } from "@tanstack/react-router";

import { LinkPeripheralPage } from "@/features/peripherals/LinkPeripheralPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/link")({
  head: () => ({
    meta: [
      { title: "Vincular periférico · Controle de Estoque de TI" },
      { name: "description", content: "Vincular periférico no sistema corporativo de controle de estoque e patrimônio de TI." },
      { property: "og:title", content: "Vincular periférico · Controle de Estoque de TI" },
      { property: "og:description", content: "Vincular periférico no sistema corporativo de controle de estoque e patrimônio de TI." },
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
        <LinkPeripheralPage />
      </RequireRole>
    </>
  );
}
