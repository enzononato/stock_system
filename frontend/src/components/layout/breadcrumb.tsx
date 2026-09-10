import { Link, useRouterState } from "@tanstack/react-router";
import { ChevronRight } from "lucide-react";
import { NAV_ITEMS } from "@/components/app/nav";

export function BreadcrumbTrail() {
  const pathname = useRouterState({ select: (s) => s.location.pathname });

  // Mapeamento dinâmico
  const activeItem = NAV_ITEMS.find((item) =>
    item.to === "/" ? pathname === "/" : pathname.startsWith(item.to)
  );

  let group = activeItem?.group || "Sistema";
  let current = activeItem?.label || "Visão Geral";

  // Rotas complementares/subfluxos
  if (pathname.startsWith("/edit/")) {
    group = "Inventário";
    current = "Edição de Ativo";
  } else if (pathname === "/remove") {
    group = "Operações";
    current = "Remoção de Ativo";
  } else if (pathname === "/terms") {
    group = "Operações";
    current = "Termos de Responsabilidade";
  } else if (pathname === "/link") {
    group = "Operações";
    current = "Vincular Periférico";
  }

  return (
    <nav aria-label="Trilha de navegação" className="flex items-center gap-1.5 text-body-sm">
      <span className="text-muted-foreground">{group}</span>
      <ChevronRight className="size-3.5 text-muted-foreground shrink-0" aria-hidden />
      <span className="font-semibold text-foreground tracking-tight">{current}</span>
    </nav>
  );
}
