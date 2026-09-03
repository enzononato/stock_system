import { useEffect, useState, useMemo, type ReactNode } from "react";
import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import {
  ChevronRight,
  Command as CommandIcon,
  LogOut,
  Menu,
  Moon,
  Search,
  ShieldCheck,
  Sun,
  X,
} from "lucide-react";

const logoUrl = "/logo-revalle.jpg";
import { cn } from "@/lib/utils";
import { useAuth } from "@/lib/auth";
import { useTheme } from "@/lib/theme";
import { NAV_GROUPS, NAV_ITEMS, type NavItem } from "@/components/app/nav";
import { Button } from "@/components/ui/button";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import {
  CommandDialog,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from "@/components/ui/command";

function useVisibleNav(): NavItem[] {
  const { user } = useAuth();
  if (!user) return [];
  return NAV_ITEMS.filter((item) => !item.roles || item.roles.includes(user.role));
}

/** Obtém a trilha de navegação (breadcrumb) com base na rota atual. */
function useRouteBreadcrumbs() {
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  const items = useVisibleNav();

  return useMemo(() => {
    if (pathname === "/") {
      return { group: "Inventário", current: "Estoque" };
    }
    if (pathname.startsWith("/edit/")) {
      return { group: "Inventário", parent: "Estoque", current: "Editar equipamento" };
    }

    // Procura em submenus primeiro, pulando a rota raiz "/" para não capturar tudo
    for (const item of items) {
      if (item.children) {
        for (const child of item.children) {
          if (pathname.startsWith(child.to)) {
            return { group: item.group, parent: item.label, current: child.label };
          }
        }
      }
      if (item.to !== "/" && pathname.startsWith(item.to)) {
        return { group: item.group, current: item.label };
      }
    }

    return { group: "Sistema", current: "Visão Geral" };
  }, [pathname, items]);
}

function NavLinks({
  collapsed,
  onNavigate,
}: {
  collapsed: boolean;
  onNavigate?: (() => void) | undefined;
}) {
  const items = useVisibleNav();
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({
    Inventário: true,
    Operação: true,
    Gestão: true,
  });

  function toggleGroup(group: string) {
    setOpenGroups((prev) => ({ ...prev, [group]: !prev[group] }));
  }

  return (
    <nav className="flex-1 space-y-3 overflow-y-auto px-2 py-3" aria-label="Navegação principal">
      {NAV_GROUPS.map((group) => {
        const groupItems = items.filter((i) => i.group === group);
        if (groupItems.length === 0) return null;
        const isOpen = openGroups[group] ?? true;

        return (
          <div key={group} className="space-y-1">
            {!collapsed ? (
              <button
                type="button"
                onClick={() => toggleGroup(group)}
                className="flex w-full items-center justify-between rounded px-2.5 py-1.5 text-eyebrow text-sidebar-foreground/70 hover:bg-sidebar-accent/50 hover:text-sidebar-foreground transition-colors group cursor-pointer"
                title={`Alternar ${group}`}
              >
                <span className="font-semibold text-[11px] uppercase tracking-wider">{group}</span>
                <ChevronRight
                  className={cn(
                    "size-3 text-muted-foreground transition-transform duration-200 group-hover:text-sidebar-foreground",
                    isOpen && "rotate-90",
                  )}
                  aria-hidden
                />
              </button>
            ) : null}

            {(isOpen || collapsed) && (
              <ul className="space-y-0.5">
                {groupItems.map((item) => {
                  const active = item.to === "/" ? pathname === "/" : pathname.startsWith(item.to);
                  const childActive =
                    item.children?.some((child) => pathname.startsWith(child.to)) ?? false;
                  const link = (
                    <Link
                      to={item.to}
                      onClick={onNavigate}
                      aria-current={active ? "page" : undefined}
                      className={cn(
                        "group relative flex items-center gap-3 rounded-md px-2.5 py-2 text-sm font-medium transition-all duration-150",
                        "text-sidebar-foreground/80 hover:bg-sidebar-accent/80 hover:text-foreground",
                        active && [
                          "bg-sidebar-accent text-foreground font-semibold shadow-xs",
                          "before:absolute before:left-0 before:top-1.5 before:bottom-1.5 before:w-1 before:rounded-r before:bg-primary",
                        ],
                        collapsed && "justify-center px-0 before:hidden",
                      )}
                    >
                      <item.icon
                        className={cn(
                          "size-4 shrink-0 transition-colors",
                          active
                            ? "text-primary"
                            : "text-muted-foreground group-hover:text-foreground",
                        )}
                        aria-hidden
                      />
                      {!collapsed && <span className="truncate">{item.label}</span>}
                      {item.children && !collapsed && (
                        <span
                          className={cn(
                            "ml-auto text-muted-foreground transition-transform text-xs",
                            childActive && "rotate-90 text-primary",
                          )}
                          aria-hidden
                        >
                          ›
                        </span>
                      )}
                      {collapsed && <span className="sr-only">{item.label}</span>}
                    </Link>
                  );

                  return (
                    <li key={item.to}>
                      {collapsed ? (
                        <Tooltip>
                          <TooltipTrigger asChild>{link}</TooltipTrigger>
                          <TooltipContent side="right">{item.label}</TooltipContent>
                        </Tooltip>
                      ) : (
                        link
                      )}
                      {item.children && !collapsed && (
                        <ul className="ml-5 mt-1 space-y-0.5 border-l border-sidebar-border pl-2">
                          {item.children
                            .filter(
                              (child) =>
                                !child.roles ||
                                child.roles.some((role) =>
                                  items.some((parent) => parent.roles?.includes(role)),
                                ),
                            )
                            .map((child) => {
                              const childIsActive = pathname.startsWith(child.to);
                              return (
                                <li key={child.to}>
                                  <Link
                                    to={child.to}
                                    onClick={onNavigate}
                                    aria-current={childIsActive ? "page" : undefined}
                                    className={cn(
                                      "group relative flex items-center gap-2 rounded-md px-2.5 py-1.5 text-xs font-medium transition-all duration-150",
                                      "text-sidebar-foreground/75 hover:bg-sidebar-accent hover:text-foreground",
                                      childIsActive && [
                                        "bg-sidebar-accent text-foreground font-semibold",
                                        "before:absolute before:left-0 before:top-1 before:bottom-1 before:w-0.5 before:rounded-r before:bg-primary",
                                      ],
                                    )}
                                  >
                                    <child.icon
                                      className={cn(
                                        "size-3.5 shrink-0 transition-colors",
                                        childIsActive
                                          ? "text-primary"
                                          : "text-muted-foreground group-hover:text-foreground",
                                      )}
                                      aria-hidden
                                    />
                                    <span className="truncate">{child.label}</span>
                                  </Link>
                                </li>
                              );
                            })}
                        </ul>
                      )}
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        );
      })}
    </nav>
  );
}

function BrandHeader({
  collapsed,
  onToggleCollapse,
}: {
  collapsed: boolean;
  onToggleCollapse: () => void;
}) {
  return (
    <div
      className={cn(
        "flex items-center gap-2.5 px-3 py-3.5 border-b border-sidebar-border bg-sidebar/50",
        collapsed && "flex-col justify-center px-1 py-3",
      )}
    >
      <Button
        variant="ghost"
        size="icon"
        className="size-8 shrink-0 text-sidebar-foreground/80 hover:text-foreground hover:bg-sidebar-accent"
        onClick={onToggleCollapse}
        aria-label={collapsed ? "Expandir menu lateral" : "Recolher menu lateral"}
        title={collapsed ? "Expandir menu lateral" : "Recolher menu lateral"}
      >
        <Menu className="size-4" />
      </Button>
      {!collapsed && (
        <div className="flex items-center gap-2.5 min-w-0 flex-1">
          <div className="relative flex size-7 shrink-0 items-center justify-center rounded-md border border-primary/30 bg-primary/10 overflow-hidden shadow-xs">
            <img src={logoUrl} alt="Revalle" className="size-full object-cover" />
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-1.5">
              <p className="truncate text-xs font-bold leading-tight text-sidebar-foreground">
                Revalle
              </p>
              <span className="rounded bg-primary/15 px-1 py-0.2 text-[9px] font-semibold text-primary uppercase">
                Core
              </span>
            </div>
            <p className="truncate text-[10px] text-muted-foreground">Controle de Patrimônio</p>
          </div>
        </div>
      )}
      {collapsed && (
        <img src={logoUrl} alt="Revalle" className="size-6 shrink-0 rounded-md object-cover mt-1" />
      )}
    </div>
  );
}

function SidebarAccount({
  collapsed,
  user,
  theme,
  toggle,
  onLogout,
}: {
  collapsed: boolean;
  user: ReturnType<typeof useAuth>["user"];
  theme: string;
  toggle: () => void;
  onLogout: () => void;
}) {
  return (
    <div className="border-t border-sidebar-border p-3 bg-sidebar/30">
      {!collapsed && user && (
        <div className="mb-2 rounded-md border border-sidebar-border/80 bg-sidebar-accent/50 px-3 py-2.5">
          <div className="flex items-center justify-between gap-1.5">
            <p className="truncate text-xs font-semibold text-foreground">{user.username}</p>
            <span className="shrink-0 rounded bg-primary/15 px-1.5 py-0.5 text-[10px] font-medium text-primary">
              {user.role}
            </span>
          </div>
        </div>
      )}
      <div className={cn("flex items-center gap-1", collapsed && "flex-col")}>
        <Button
          variant="ghost"
          size="sm"
          className={cn(
            "min-h-9 text-muted-foreground hover:text-foreground",
            collapsed ? "w-full justify-center px-0" : "flex-1 justify-start px-2.5",
          )}
          onClick={toggle}
          aria-label={theme === "dark" ? "Ativar tema claro" : "Ativar tema escuro"}
        >
          {theme === "dark" ? (
            <Sun className="size-4 text-amber-400" />
          ) : (
            <Moon className="size-4 text-primary" />
          )}
          {!collapsed && (
            <span className="ml-2 text-xs">{theme === "dark" ? "Modo Claro" : "Modo Escuro"}</span>
          )}
        </Button>
        <Button
          variant="ghost"
          size="sm"
          className={cn(
            "min-h-9 text-muted-foreground hover:text-destructive hover:bg-destructive/10",
            collapsed ? "w-full justify-center px-0" : "px-2.5",
          )}
          onClick={onLogout}
          aria-label="Sair da conta"
          title="Encerrar sessão"
        >
          <LogOut className="size-4" />
          {!collapsed && <span className="ml-2 text-xs">Sair</span>}
        </Button>
      </div>
    </div>
  );
}

export function AppShell({ children }: { children: ReactNode }) {
  const { user, logout } = useAuth();
  const { theme, toggle } = useTheme();
  const navigate = useNavigate();
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [commandOpen, setCommandOpen] = useState(false);
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  const breadcrumbs = useRouteBreadcrumbs();
  const visibleNav = useVisibleNav();

  useEffect(() => {
    setMobileOpen(false);
  }, [pathname]);

  // Atalho global Ctrl+K / Cmd+K para abrir a paleta de comandos
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.key === "k" || e.key === "K") && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        setCommandOpen((prev) => !prev);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  async function handleLogout() {
    await logout();
    void navigate({ to: "/login" });
  }

  function handleCommandSelect(to: string) {
    setCommandOpen(false);
    void navigate({ to });
  }

  const account = (
    <SidebarAccount
      collapsed={collapsed}
      user={user}
      theme={theme}
      toggle={toggle}
      onLogout={() => void handleLogout()}
    />
  );

  return (
    <div className="flex min-h-screen bg-background text-foreground selection:bg-primary/20 selection:text-primary">
      {/* Sidebar Desktop */}
      <aside
        className={cn(
          "sticky top-0 hidden h-screen shrink-0 flex-col border-r border-sidebar-border bg-sidebar lg:flex transition-[width] duration-200 z-30",
          collapsed ? "w-16" : "w-64",
        )}
      >
        <BrandHeader collapsed={collapsed} onToggleCollapse={() => setCollapsed((c) => !c)} />
        <NavLinks collapsed={collapsed} />
        {account}
      </aside>

      {/* Drawer Mobile */}
      {mobileOpen && (
        <div className="fixed inset-0 z-50 lg:hidden">
          <button
            className="absolute inset-0 bg-background/80 backdrop-blur-sm"
            aria-label="Fechar menu"
            onClick={() => setMobileOpen(false)}
          />
          <div className="absolute inset-y-0 left-0 flex w-72 max-w-[85vw] flex-col border-r border-sidebar-border bg-sidebar shadow-2xl animate-in slide-in-from-left duration-200">
            <div className="flex items-center justify-between pr-2">
              <BrandHeader collapsed={false} onToggleCollapse={() => setMobileOpen(false)} />
              <Button
                variant="ghost"
                size="icon"
                className="size-8 text-muted-foreground hover:text-foreground"
                aria-label="Fechar menu"
                onClick={() => setMobileOpen(false)}
              >
                <X className="size-5" />
              </Button>
            </div>
            <NavLinks collapsed={false} onNavigate={() => setMobileOpen(false)} />
            <div className="mt-auto">{account}</div>
          </div>
        </div>
      )}

      {/* Conteúdo Principal + Header Enterprise */}
      <div className="flex min-w-0 flex-1 flex-col">
        <header className="sticky top-0 z-40 border-b border-border bg-background/90 backdrop-blur-md">
          <div className="flex h-14 items-center px-4 sm:px-6 gap-4">
            {/* Esquerda: Botão mobile + Breadcrumb corporativo */}
            <div className="flex items-center gap-3 min-w-0 flex-1">
              <Button
                variant="ghost"
                size="icon"
                className="size-8 lg:hidden text-foreground hover:bg-muted"
                aria-label="Abrir menu lateral"
                onClick={() => setMobileOpen(true)}
              >
                <Menu className="size-5" />
              </Button>

              <nav
                aria-label="Trilha de navegação"
                className="flex items-center gap-1.5 text-xs text-muted-foreground min-w-0"
              >
                <span className="text-eyebrow text-muted-foreground/80">{breadcrumbs.group}</span>
                <ChevronRight className="size-3 shrink-0 text-muted-foreground/50" aria-hidden />
                {breadcrumbs.parent && (
                  <>
                    <span className="truncate hidden sm:inline">{breadcrumbs.parent}</span>
                    <ChevronRight
                      className="size-3 shrink-0 text-muted-foreground/50 hidden sm:inline"
                      aria-hidden
                    />
                  </>
                )}
                <span className="truncate font-semibold text-foreground">
                  {breadcrumbs.current}
                </span>
              </nav>
            </div>

            {/* Centro: Busca rápida (Command Palette) */}
            <button
              type="button"
              onClick={() => setCommandOpen(true)}
              className="flex items-center gap-2 rounded-lg border border-border/70 bg-muted/30 hover:bg-muted/60 hover:border-primary/40 px-3 py-1.5 text-xs text-muted-foreground transition-all duration-150 cursor-pointer shadow-2xs"
              title="Abrir busca rápida (Ctrl+K)"
            >
              <Search className="size-3.5" aria-hidden />
              <span className="hidden sm:inline">Buscar rota ou ação…</span>
              <kbd className="hidden sm:inline-flex items-center gap-0.5 rounded border border-border bg-background px-1.5 py-0.5 text-[10px] font-mono text-muted-foreground">
                <CommandIcon className="size-2.5" /> K
              </kbd>
            </button>

            {/* Direita: espaçador para manter a busca centralizada */}
            <div className="hidden lg:block flex-1" />
          </div>
        </header>

        {/* Paleta de Comandos (Modal cmdk) */}
        <CommandDialog open={commandOpen} onOpenChange={setCommandOpen}>
          <CommandInput placeholder="Digite o nome de uma tela ou funcionalidade…" />
          <CommandList>
            <CommandEmpty>Nenhum resultado encontrado.</CommandEmpty>
            {NAV_GROUPS.map((group) => {
              const groupItems = visibleNav.filter((i) => i.group === group);
              if (groupItems.length === 0) return null;

              return (
                <CommandGroup key={group} heading={group}>
                  {groupItems.map((item) => (
                    <CommandItem
                      key={item.to}
                      onSelect={() => handleCommandSelect(item.to)}
                      className="flex items-center gap-2.5 py-2 cursor-pointer"
                    >
                      <item.icon className="size-4 text-primary shrink-0" />
                      <span className="font-medium text-sm">{item.label}</span>
                    </CommandItem>
                  ))}
                  {groupItems.map((item) =>
                    item.children?.map((child) => (
                      <CommandItem
                        key={child.to}
                        onSelect={() => handleCommandSelect(child.to)}
                        className="flex items-center gap-2.5 py-2 pl-6 cursor-pointer"
                      >
                        <child.icon className="size-3.5 text-muted-foreground shrink-0" />
                        <span className="text-sm">{child.label}</span>
                        <span className="ml-auto text-[10px] text-muted-foreground uppercase">
                          {item.label}
                        </span>
                      </CommandItem>
                    )),
                  )}
                </CommandGroup>
              );
            })}
          </CommandList>
        </CommandDialog>

        {/* Viewport Principal */}
        <main className="mx-auto w-full max-w-[1600px] flex-1 px-4 py-5 sm:px-6 sm:py-7">
          {children}
        </main>
      </div>
    </div>
  );
}
