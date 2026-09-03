import { createFileRoute } from "@tanstack/react-router";

import { EditItemPage } from "@/features/items/EditItemPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/edit/$id")({
  head: () => ({
    meta: [
      { title: "Editar item · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Editar item no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Editar item · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Editar item no sistema corporativo de controle de estoque e patrimônio de TI.",
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
        <EditItemPage />
      </RequireRole>
    </>
  );
}
