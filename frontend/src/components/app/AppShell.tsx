import { useState, useEffect, type ReactNode } from "react";
import { useRouterState } from "@tanstack/react-router";

import { Sidebar } from "@/components/layout/sidebar";
import { Topbar } from "@/components/layout/topbar";
import { CommandPalette } from "@/components/layout/command-palette";
import { AssetDetailsPanel } from "@/components/app/AssetDetailsPanel";
import { Sheet, SheetContent } from "@/components/ui/sheet";
import type { Item } from "@/api/items";

interface AppShellProps {
  children: ReactNode;
}

export function AppShell({ children }: AppShellProps) {
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false);
  const [selectedAsset, setSelectedAsset] = useState<Item | null>(null);

  const pathname = useRouterState({ select: (s) => s.location.pathname });

  // Atalho global Ctrl+K / ⌘K
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setCommandPaletteOpen((prev) => !prev);
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  // Determina se a tela exige container de dados densos (1440px) ou leitura/formulário (1120px)
  const isReadingRoute =
    pathname.startsWith("/register") ||
    pathname.startsWith("/edit/") ||
    pathname === "/remove" ||
    pathname === "/loan" ||
    pathname === "/terms" ||
    pathname === "/link";

  return (
    <div className="flex h-screen h-[100dvh] w-full overflow-hidden bg-background text-foreground">
      {/* Sidebar Desktop Fixa */}
      <Sidebar className="hidden lg:flex" />

      {/* Sidebar Mobile via Sheet Drawer */}
      <Sheet open={mobileNavOpen} onOpenChange={setMobileNavOpen}>
        <SheetContent side="left" className="p-0 w-64 bg-surface border-r border-border">
          <Sidebar onNavigate={() => setMobileNavOpen(false)} />
        </SheetContent>
      </Sheet>

      {/* Área Principal */}
      <div className="flex flex-1 flex-col min-w-0 overflow-hidden">
        <Topbar
          onOpenMobileNav={() => setMobileNavOpen(true)}
          onOpenCommandPalette={() => setCommandPaletteOpen(true)}
        />

        <main className="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8">
          <div className={isReadingRoute ? "page-container-reading" : "page-container-dense"}>
            {children}
          </div>
        </main>
      </div>

      {/* Command Palette Global */}
      <CommandPalette
        open={commandPaletteOpen}
        onOpenChange={setCommandPaletteOpen}
        onSelectAsset={(item) => setSelectedAsset(item)}
      />

      {/* Painel Lateral de Detalhes do Ativo Global */}
      <AssetDetailsPanel
        item={selectedAsset}
        open={selectedAsset !== null}
        onOpenChange={(open) => !open && setSelectedAsset(null)}
      />
    </div>
  );
}
