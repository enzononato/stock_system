import { createFileRoute } from "@tanstack/react-router";

import { DashboardPage } from "@/features/dashboard/DashboardPage";

export const Route = createFileRoute("/_shell/")({
  head: () => ({
    meta: [
      { title: "Dashboard Executivo · Controle de Patrimônio" },
      {
        name: "description",
        content: "Visão executiva e panorama operacional do patrimônio e estoque de TI.",
      },
      { property: "og:title", content: "Dashboard Executivo · Controle de Patrimônio" },
      {
        property: "og:description",
        content: "Visão executiva e panorama operacional do patrimônio e estoque de TI.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: RouteComponent,
});

function RouteComponent() {
  return <DashboardPage />;
}
