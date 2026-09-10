import { createFileRoute } from "@tanstack/react-router";

import { StockPage } from "@/features/stock/StockPage";

export const Route = createFileRoute("/_shell/stock")({
  head: () => ({
    meta: [
      { title: "Estoque de Equipamentos · Controle de Patrimônio" },
      {
        name: "description",
        content: "Data grid operacional e gestão do inventário de patrimônio de TI.",
      },
      { property: "og:title", content: "Estoque de Equipamentos · Controle de Patrimônio" },
      {
        property: "og:description",
        content: "Data grid operacional e gestão do inventário de patrimônio de TI.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: RouteComponent,
});

function RouteComponent() {
  return <StockPage />;
}
