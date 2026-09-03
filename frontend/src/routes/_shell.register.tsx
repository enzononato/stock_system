import { createFileRoute } from "@tanstack/react-router";

import { RegisterItemPage } from "@/features/items/RegisterItemPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/register")({
  head: () => ({
    meta: [
      { title: "Cadastrar item · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Cadastrar item no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Cadastrar item · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Cadastrar item no sistema corporativo de controle de estoque e patrimônio de TI.",
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
        <RegisterItemPage />
      </RequireRole>
    </>
  );
}
