import { useState } from "react";
import { Link, useRouterState } from "@tanstack/react-router";
import { ChevronRight } from "lucide-react";

import { cn } from "@/lib/utils";
import { NAV_GROUPS, NAV_ITEMS, type NavItem } from "@/components/app/nav";
import { useAuth } from "@/lib/auth";

interface SidebarProps {
  onNavigate?: () => void;
  className?: string;
}

export function Sidebar({ onNavigate, className }: SidebarProps) {
  const { user } = useAuth();
  const pathname = useRouterState({ select: (s) => s.location.pathname });

  // Controle de colapso de grupos
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({
    Principal: true,
    Inventário: true,
    Operações: true,
    Gestão: true,
  });

  function toggleGroup(group: string) {
    setOpenGroups((prev) => ({ ...prev, [group]: !prev[group] }));
  }

  // Filtragem por permissão
  const visibleItems = NAV_ITEMS.filter((item) => !item.roles || (user && item.roles.includes(user.role)));

  return (
    <aside
      className={cn(
        "flex flex-col h-full w-64 bg-surface border-r border-border select-none",
        className
      )}
      aria-label="Menu principal"
    >
      {/* Topo institucional com logo e subtítulo */}
      <div className="h-14 px-4 flex items-center justify-between border-b border-border shrink-0">
        <div className="flex items-center gap-2.5 min-w-0">
          <img
            src="/logo-revalle.png"
            alt="Revalle"
            className="size-8 rounded-[4px] object-cover shrink-0 border border-border"
          />
          <div className="flex flex-col min-w-0">
            <span className="font-semibold text-sm tracking-tight text-foreground leading-none">
              REVALLE
            </span>
            <span className="text-[11px] text-muted-foreground truncate mt-0.5">
              Controle de Patrimônio
            </span>
          </div>
        </div>
      </div>

      {/* Lista de navegação em grupos */}
      <nav className="flex-1 overflow-y-auto px-3 py-4 space-y-4">
        {NAV_GROUPS.map((group) => {
          const items = visibleItems.filter((i) => i.group === group);
          if (items.length === 0) return null;
          const isOpen = openGroups[group] ?? true;

          return (
            <div key={group} className="space-y-1">
              <button
                type="button"
                onClick={() => toggleGroup(group)}
                className="flex w-full items-center justify-between px-2 py-1 text-caption text-muted-foreground hover:text-foreground transition-colors cursor-pointer"
              >
                <span>{group}</span>
                <ChevronRight
                  className={cn(
                    "size-3 text-muted-foreground transition-transform duration-micro",
                    isOpen && "rotate-90"
                  )}
                  aria-hidden
                />
              </button>

              {isOpen && (
                <ul className="space-y-0.5 mt-0.5">
                  {items.map((item) => {
                    const active = item.to === "/" ? pathname === "/" : pathname.startsWith(item.to);
                    const Icon = item.icon;

                    return (
                      <li key={item.to}>
                        <Link
                          to={item.to}
                          onClick={onNavigate}
                          className={cn(
                            "flex items-center gap-2.5 px-2.5 py-1.5 rounded text-body-sm transition-colors duration-micro",
                            active
                              ? "bg-surface-alt text-foreground font-semibold border-l-2 border-foreground"
                              : "text-secondary hover:text-foreground hover:bg-surface-alt/60"
                          )}
                        >
                          <Icon className={cn("size-4 shrink-0", active ? "text-foreground" : "text-muted-foreground")} />
                          <span className="truncate">{item.label}</span>
                        </Link>
                      </li>
                    );
                  })}
                </ul>
              )}
            </div>
          );
        })}
      </nav>

      {/* Rodapé institucional com operador */}
      {user && (
        <div className="p-3 border-t border-border shrink-0 bg-surface-alt/20">
          <div className="flex items-center justify-between px-1">
            <div className="flex flex-col truncate">
              <span className="text-body-sm font-medium text-foreground truncate">{user.username}</span>
              <span className="text-[11px] text-muted-foreground font-mono">{user.role}</span>
            </div>
            <span className="text-[10px] uppercase font-mono px-1.5 py-0.5 rounded-sm border border-border text-muted-foreground">
              {user.role}
            </span>
          </div>
        </div>
      )}
    </aside>
  );
}
