import { useState, useEffect } from "react";
import { useNavigate } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import {
  Boxes,
  Command as CommandIcon,
  FileSignature,
  FileSpreadsheet,
  History,
  Keyboard,
  PackagePlus,
  Search,
  Trash2,
  Undo2,
  ArrowRightLeft,
  UserRoundCog,
  Building2,
  BarChart3,
  UserX,
  Link2,
} from "lucide-react";

import {
  CommandDialog,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
  CommandSeparator,
} from "@/components/ui/command";
import { listItemsPaginated, type Item } from "@/api/items";

interface CommandPaletteProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSelectAsset?: (item: Item) => void;
}

export function CommandPalette({ open, onOpenChange, onSelectAsset }: CommandPaletteProps) {
  const navigate = useNavigate();

  // Busca rápida de patrimônios
  const { data } = useQuery({
    queryKey: ["items-command-palette"],
    queryFn: () => listItemsPaginated({ limit: 100 }),
    enabled: open,
  });

  const items = data?.items ?? [];

  function handleSelect(to: string) {
    onOpenChange(false);
    void navigate({ to });
  }

  return (
    <CommandDialog open={open} onOpenChange={onOpenChange}>
      <CommandInput placeholder="Buscar patrimônio, navegação ou ação rápida (Ctrl+K)..." />
      <CommandList className="max-h-[380px] overflow-y-auto">
        <CommandEmpty className="py-6 text-center text-body-sm text-muted-foreground">
          Nenhum resultado encontrado.
        </CommandEmpty>

        <CommandGroup heading="Módulos Principais">
          <CommandItem onSelect={() => handleSelect("/")}>
            <Boxes className="mr-2 size-4 text-foreground" />
            <span>Dashboard Executivo</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/stock")}>
            <Boxes className="mr-2 size-4 text-foreground" />
            <span>Estoque de Equipamentos</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/register")}>
            <PackagePlus className="mr-2 size-4 text-foreground" />
            <span>Cadastrar Novo Ativo</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/loan")}>
            <ArrowRightLeft className="mr-2 size-4 text-foreground" />
            <span>Novo Empréstimo</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/return")}>
            <Undo2 className="mr-2 size-4 text-foreground" />
            <span>Devolução de Equipamento</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/peripherals")}>
            <Keyboard className="mr-2 size-4 text-foreground" />
            <span>Gestão de Periféricos</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/offboarding-ti")}>
            <UserX className="mr-2 size-4 text-foreground" />
            <span>Esteira de Desligamentos (Offboarding)</span>
          </CommandItem>
        </CommandGroup>

        <CommandSeparator />

        <CommandGroup heading="Ações Operacionais & Documentos">
          <CommandItem onSelect={() => handleSelect("/terms")}>
            <FileSignature className="mr-2 size-4 text-foreground" />
            <span>Termos de Responsabilidade</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/link")}>
            <Link2 className="mr-2 size-4 text-foreground" />
            <span>Vincular Periférico a Ativo</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/remove")}>
            <Trash2 className="mr-2 size-4 text-foreground" />
            <span>Remoção / Baixa de Ativo</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/history")}>
            <History className="mr-2 size-4 text-foreground" />
            <span>Histórico de Operações</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/report")}>
            <FileSpreadsheet className="mr-2 size-4 text-foreground" />
            <span>Relatórios Mensais</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/charts")}>
            <BarChart3 className="mr-2 size-4 text-foreground" />
            <span>Indicadores & Análise</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/unidades")}>
            <Building2 className="mr-2 size-4 text-foreground" />
            <span>Gestão de Unidades / Filiais</span>
          </CommandItem>
          <CommandItem onSelect={() => handleSelect("/users")}>
            <UserRoundCog className="mr-2 size-4 text-foreground" />
            <span>Gestão de Operadores & Usuários</span>
          </CommandItem>
        </CommandGroup>

        {items.length > 0 && (
          <>
            <CommandSeparator />
            <CommandGroup heading="Ativos no Estoque (Acesso Rápido)">
              {items.slice(0, 15).map((item) => (
                <CommandItem
                  key={item.id}
                  onSelect={() => {
                    onOpenChange(false);
                    if (onSelectAsset) {
                      onSelectAsset(item);
                    } else {
                      void navigate({ to: "/stock" });
                    }
                  }}
                  className="flex items-center justify-between"
                >
                  <div className="flex items-center gap-2">
                    <span className="font-mono text-xs num font-semibold">#{item.id}</span>
                    <span>
                      {item.brand} {item.model} ({item.tipo})
                    </span>
                  </div>
                  <span className="text-[11px] text-muted-foreground font-mono">
                    {item.identificador || item.revenda || item.status}
                  </span>
                </CommandItem>
              ))}
            </CommandGroup>
          </>
        )}
      </CommandList>
    </CommandDialog>
  );
}

export function CommandPaletteTrigger({ onClick }: { onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="flex items-center justify-between gap-3 h-8 w-44 sm:w-60 px-2.5 rounded border border-border bg-surface-alt/60 hover:bg-surface-alt hover:border-border-strong text-muted-foreground hover:text-foreground transition-colors text-xs cursor-pointer"
      title="Busca rápida (Ctrl+K)"
    >
      <div className="flex items-center gap-1.5 truncate">
        <Search className="size-3.5 shrink-0" />
        <span className="truncate">Buscar no sistema...</span>
      </div>
      <kbd className="hidden sm:inline-flex items-center gap-0.5 rounded border border-border bg-surface px-1.5 py-0.5 text-[10px] font-mono text-muted-foreground">
        ⌘K
      </kbd>
    </button>
  );
}
