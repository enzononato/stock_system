import { createFileRoute } from "@tanstack/react-router";

import { UsersPage } from "@/features/admin/UsersPage";
import { RequireRole } from "@/components/app/RequireRole";

export const Route = createFileRoute("/_shell/users")({
  head: () => ({
    meta: [
      { title: "Usuários · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Usuários no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Usuários · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Usuários no sistema corporativo de controle de estoque e patrimônio de TI.",
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
      <RequireRole roles={["Gestor"]}>
        <UsersPage />
      </RequireRole>
    </>
  );
}
