import { useEffect, useState, type ReactNode } from "react";
import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import { LogOut, Menu, Moon, PanelLeftClose, PanelLeftOpen, Sun, X } from "lucide-react";

import logoAsset from "@/assets/revalle-logo.png.asset.json";
import { cn } from "@/lib/utils";
import { useAuth } from "@/lib/auth";
import { useTheme } from "@/lib/theme";
import { NAV_GROUPS, NAV_ITEMS, type NavItem } from "@/components/app/nav";
import { Button } from "@/components/ui/button";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";

function useVisibleNav(): NavItem[] {
  const { user } = useAuth();
  if (!user) return [];
  return NAV_ITEMS.filter((item) => !item.roles || item.roles.includes(user.role));
}

function NavLinks({ collapsed, onNavigate }: { collapsed: boolean; onNavigate?: (() => void) | undefined }) {
  const items = useVisibleNav();
  const pathname = useRouterState({ select: (s) => s.location.pathname });

  return (
    <nav className="flex-1 space-y-5 overflow-y-auto px-3 py-4" aria-label="Navegação principal">
      {NAV_GROUPS.map((group) => {
        const groupItems = items.filter((i) => i.group === group);
        if (groupItems.length === 0) return null;
        return (
          <div key={group}>
            {!collapsed && <p className="text-eyebrow px-2 pb-2">{group}</p>}
            <ul className="space-y-1">
              {groupItems.map((item) => {
                const active = item.to === "/" ? pathname === "/" : pathname.startsWith(item.to);
                const childActive = item.children?.some((child) => pathname.startsWith(child.to)) ?? false;
                const link = (
                  <Link to={item.to} onClick={onNavigate} aria-current={active ? "page" : undefined} className={cn("flex items-center gap-3 rounded-md px-2.5 py-2 text-sm font-medium transition-colors", "text-sidebar-foreground hover:bg-sidebar-accent hover:text-sidebar-accent-foreground", active && "bg-sidebar-accent text-sidebar-accent-foreground", collapsed && "justify-center px-0")}>
                    <item.icon className="size-4 shrink-0" aria-hidden />
                    {!collapsed && <span className="truncate">{item.label}</span>}
                    {item.children && !collapsed && <span className={cn("ml-auto text-muted-foreground transition-transform", childActive && "rotate-90")} aria-hidden>›</span>}
                    {collapsed && <span className="sr-only">{item.label}</span>}
                  </Link>
                );
                return (
                  <li key={item.to}>
                    {collapsed ? <Tooltip><TooltipTrigger asChild>{link}</TooltipTrigger><TooltipContent side="right">{item.label}</TooltipContent></Tooltip> : link}
                    {item.children && !collapsed && (
                      <ul className="ml-5 mt-1 space-y-0.5 border-l border-sidebar-border pl-2">
                        {item.children.filter((child) => !child.roles || child.roles.some((role) => items.some((parent) => parent.roles?.includes(role)))).map((child) => {
                          const childIsActive = pathname.startsWith(child.to);
                          return <li key={child.to}><Link to={child.to} onClick={onNavigate} aria-current={childIsActive ? "page" : undefined} className={cn("flex items-center gap-2 rounded-md px-2.5 py-1.5 text-xs font-medium text-sidebar-foreground/80 hover:bg-sidebar-accent hover:text-sidebar-accent-foreground", childIsActive && "bg-sidebar-accent text-sidebar-accent-foreground")}><child.icon className="size-3.5 shrink-0" aria-hidden /><span className="truncate">{child.label}</span></Link></li>;
                        })}
                      </ul>
                    )}
                  </li>
                );
              })}
            </ul>
          </div>
        );
      })}
    </nav>
  );
}

function BrandMark({ collapsed }: { collapsed: boolean }) {
  return <div className={cn("flex items-center gap-2.5 px-4 py-4", collapsed && "justify-center px-0")}><img src={logoAsset.url} alt="Revalle" className="size-8 shrink-0 rounded-md object-cover" />{!collapsed && <div className="min-w-0"><p className="truncate text-sm font-bold leading-tight">Controle de Patrimônio</p><p className="truncate text-[11px] text-muted-foreground">Gestão de ativos</p></div>}</div>;
}

function SidebarAccount({ collapsed, user, theme, toggle, onLogout }: { collapsed: boolean; user: ReturnType<typeof useAuth>["user"]; theme: string; toggle: () => void; onLogout: () => void }) {
  return <div className="border-t border-sidebar-border p-3">{!collapsed && user && <div className="mb-2 rounded-md border border-sidebar-border bg-sidebar-accent/40 px-3 py-2.5"><p className="truncate text-sm font-semibold">{user.username}</p><p className="mt-0.5 truncate text-[11px] text-muted-foreground">{user.role}</p></div>}<div className={cn("flex items-center gap-1", collapsed && "flex-col")}><Button variant="ghost" size="sm" className={cn("min-h-9", collapsed ? "w-full justify-center px-0" : "flex-1 justify-start")} onClick={toggle} aria-label={theme === "dark" ? "Ativar tema claro" : "Ativar tema escuro"}>{theme === "dark" ? <Sun className="size-4" /> : <Moon className="size-4" />}{!collapsed && <span className="ml-2">{theme === "dark" ? "Tema claro" : "Tema escuro"}</span>}</Button><Button variant="ghost" size="sm" className={cn("min-h-9 text-muted-foreground hover:text-destructive", collapsed ? "w-full justify-center px-0" : "px-2")} onClick={onLogout} aria-label="Sair da conta"><LogOut className="size-4" />{!collapsed && <span className="ml-2">Sair</span>}</Button></div></div>;
}

export function AppShell({ children }: { children: ReactNode }) {
  const { user, logout } = useAuth();
  const { theme, toggle } = useTheme();
  const navigate = useNavigate();
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  useEffect(() => { setMobileOpen(false); }, [pathname]);
  async function handleLogout() { await logout(); void navigate({ to: "/login" }); }
  const account = <SidebarAccount collapsed={collapsed} user={user} theme={theme} toggle={toggle} onLogout={() => void handleLogout()} />;
  const collapseButton = <div className="border-t border-sidebar-border p-3"><Button variant="ghost" size="sm" className="w-full justify-center" onClick={() => setCollapsed((c) => !c)} aria-label={collapsed ? "Expandir menu" : "Recolher menu"}>{collapsed ? <PanelLeftOpen className="size-4" /> : <PanelLeftClose className="size-4" />}{!collapsed && <span className="ml-2">Recolher</span>}</Button></div>;
  return <div className="flex min-h-screen bg-background"><aside className={cn("sticky top-0 hidden h-screen shrink-0 flex-col border-r border-sidebar-border bg-sidebar lg:flex", collapsed ? "w-16" : "w-64")}><BrandMark collapsed={collapsed} /><NavLinks collapsed={collapsed} />{account}{collapseButton}</aside>{mobileOpen && <div className="fixed inset-0 z-50 lg:hidden"><button className="absolute inset-0 bg-foreground/40 backdrop-blur-sm" aria-label="Fechar menu" onClick={() => setMobileOpen(false)} /><div className="absolute inset-y-0 left-0 flex w-72 max-w-[85vw] flex-col border-r border-sidebar-border bg-sidebar"><div className="flex items-center justify-between"><BrandMark collapsed={false} /><Button variant="ghost" size="icon" className="mr-2 min-h-11 min-w-11" aria-label="Fechar menu" onClick={() => setMobileOpen(false)}><X className="size-5" /></Button></div><NavLinks collapsed={false} onNavigate={() => setMobileOpen(false)} /><div className="mt-auto">{account}{collapseButton}</div></div></div>}<div className="flex min-w-0 flex-1 flex-col"><header className="sticky top-0 z-40 border-b border-border bg-background/85 backdrop-blur"><div className="flex items-center px-4 py-2.5 sm:px-6"><Button variant="ghost" size="icon" className="min-h-11 min-w-11 lg:hidden" aria-label="Abrir menu" onClick={() => setMobileOpen(true)}><Menu className="size-5" /></Button><div className="flex-1" /></div></header><main className="mx-auto w-full max-w-[1600px] flex-1 px-4 py-5 sm:px-6 sm:py-7">{children}</main></div></div>;
}
