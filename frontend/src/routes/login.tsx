import { createFileRoute } from "@tanstack/react-router";

import { LoginPage } from "@/features/auth/LoginPage";

export const Route = createFileRoute("/login")({
  ssr: false,
  head: () => ({
    meta: [
      { title: "Entrar · Controle de Estoque de TI" },
      {
        name: "description",
        content: "Acesse o sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:title", content: "Entrar · Controle de Estoque de TI" },
      {
        property: "og:description",
        content: "Acesse o sistema corporativo de controle de estoque e patrimônio de TI.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
      { name: "robots", content: "noindex" },
    ],
  }),
  component: LoginPage,
});
