import { createFileRoute } from "@tanstack/react-router";

import { ChartsPage } from "@/features/charts/ChartsPage";

export const Route = createFileRoute("/_shell/charts")({
  head: () => ({
    meta: [
      { title: "Indicadores · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Indicadores no sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Indicadores · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Indicadores no sistema corporativo de controle de estoque e patrimônio de TI.",
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
      <ChartsPage />
    </>
  );
}
