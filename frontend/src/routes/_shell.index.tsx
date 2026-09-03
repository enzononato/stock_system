import { createFileRoute } from "@tanstack/react-router";

import { StockPage } from "@/features/stock/StockPage";

export const Route = createFileRoute("/_shell/")({
  head: () => ({
    meta: [
      { title: "Estoque de equipamentos · Controle de Patrimônio" },
      {
        name: "description",
        content: "Estoque de equipamentos no sistema corporativo de controle de patrimônio.",
      },
      { property: "og:title", content: "Estoque de equipamentos · Controle de Patrimônio" },
      {
        property: "og:description",
        content: "Estoque de equipamentos no sistema corporativo de controle de patrimônio.",
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
