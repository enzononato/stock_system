import { Menu, Moon, Sun, LogOut } from "lucide-react";
import { BreadcrumbTrail } from "./breadcrumb";
import { CommandPaletteTrigger } from "./command-palette";
import { useTheme } from "@/lib/theme";
import { useAuth } from "@/lib/auth";
import { Button } from "@/components/ui/button";

interface TopbarProps {
  onOpenMobileNav: () => void;
  onOpenCommandPalette: () => void;
}

export function Topbar({ onOpenMobileNav, onOpenCommandPalette }: TopbarProps) {
  const { theme, toggle } = useTheme();
  const { logout } = useAuth();

  return (
    <header className="h-14 bg-surface border-b border-border px-4 sm:px-6 flex items-center justify-between gap-3 shrink-0 z-30">
      <div className="flex items-center gap-3 min-w-0">
        <Button
          variant="ghost"
          size="sm"
          className="lg:hidden size-8 p-0"
          onClick={onOpenMobileNav}
          aria-label="Abrir menu de navegação"
        >
          <Menu className="size-4" />
        </Button>
        <BreadcrumbTrail />
      </div>

      <div className="flex items-center gap-2 sm:gap-3 shrink-0">
        <CommandPaletteTrigger onClick={onOpenCommandPalette} />

        <Button
          variant="ghost"
          size="sm"
          className="size-8 p-0 text-muted-foreground hover:text-foreground"
          onClick={toggle}
          aria-label={theme === "dark" ? "Ativar modo claro" : "Ativar modo escuro"}
          title={theme === "dark" ? "Modo claro" : "Modo escuro"}
        >
          {theme === "dark" ? <Sun className="size-4" /> : <Moon className="size-4" />}
        </Button>

        <Button
          variant="ghost"
          size="sm"
          className="size-8 p-0 text-muted-foreground hover:text-foreground"
          onClick={logout}
          aria-label="Sair da sessão"
          title="Sair da sessão"
        >
          <LogOut className="size-4" />
        </Button>
      </div>
    </header>
  );
}
