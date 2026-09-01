import { useEffect, useRef, useState } from "react";
import { Check, ChevronDown, Search, X } from "lucide-react";

import { cn } from "@/lib/utils";

export interface SearchableOption {
  value: string;
  label: string;
  subtitle?: string | undefined;
}

interface SearchableSelectProps {
  options: SearchableOption[];
  value: string;
  onValueChange: (val: string) => void;
  placeholder?: string | undefined;
  searchPlaceholder?: string | undefined;
  emptyMessage?: string | undefined;
  className?: string | undefined;
  disabled?: boolean | undefined;
  id?: string | undefined;
}

/**
 * Select com busca client-side. Mantido como componente próprio (em vez do
 * Select do Radix) porque as listas de itens/periféricos chegam com centenas de
 * registros e precisam ser filtráveis por identificador, marca e modelo.
 */
export function SearchableSelect({
  options,
  value,
  onValueChange,
  placeholder = "Selecione uma opção…",
  searchPlaceholder = "Digite para buscar…",
  emptyMessage = "Nenhum resultado encontrado",
  className,
  disabled = false,
  id,
}: SearchableSelectProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState("");
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const selectedOption = options.find((o) => o.value === value);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) setIsOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    if (isOpen) {
      const t = setTimeout(() => inputRef.current?.focus(), 40);
      return () => clearTimeout(t);
    }
    setSearch("");
    return undefined;
  }, [isOpen]);

  const term = search.trim().toLowerCase();
  const filtered = term
    ? options.filter(
        (o) =>
          o.label.toLowerCase().includes(term) ||
          o.value.toLowerCase().includes(term) ||
          (o.subtitle ?? "").toLowerCase().includes(term),
      )
    : options;

  return (
    <div ref={containerRef} className={cn("relative w-full", className)}>
      <button
        id={id}
        type="button"
        disabled={disabled}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
        onClick={() => !disabled && setIsOpen((v) => !v)}
        onKeyDown={(e) => {
          if (e.key === "Escape") setIsOpen(false);
        }}
        className={cn(
          "flex h-10 w-full items-center justify-between gap-2 rounded-md border border-input bg-background px-3 py-2 text-sm shadow-xs transition-colors",
          "focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-50",
          isOpen && "ring-2 ring-ring",
        )}
      >
        <span className="min-w-0 truncate text-left">
          {selectedOption ? (
            <>
              <span className="font-medium text-foreground">{selectedOption.label}</span>
              {selectedOption.subtitle ? (
                <span className="ml-1.5 text-xs text-muted-foreground">{selectedOption.subtitle}</span>
              ) : null}
            </>
          ) : (
            <span className="text-muted-foreground">{placeholder}</span>
          )}
        </span>
        <ChevronDown
          className={cn("size-4 shrink-0 text-muted-foreground transition-transform", isOpen && "rotate-180")}
          aria-hidden
        />
      </button>

      {isOpen ? (
        <div className="absolute left-0 top-full z-50 mt-1 flex max-h-80 w-full flex-col gap-2 rounded-md border border-border bg-popover p-2 shadow-lg">
          <div className="relative shrink-0">
            <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden />
            <input
              ref={inputRef}
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder={searchPlaceholder}
              aria-label={searchPlaceholder}
              className="h-9 w-full rounded-md border border-input bg-background pl-8 pr-8 text-sm focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
            />
            {search ? (
              <button
                type="button"
                onClick={() => setSearch("")}
                aria-label="Limpar busca"
                className="absolute right-2 top-1/2 -translate-y-1/2 rounded-sm p-0.5 text-muted-foreground hover:bg-accent"
              >
                <X className="size-3.5" aria-hidden />
              </button>
            ) : null}
          </div>

          <div className="min-h-0 flex-1 overflow-y-auto" role="listbox">
            {filtered.length === 0 ? (
              <p className="py-6 text-center text-sm text-muted-foreground">{emptyMessage}</p>
            ) : (
              filtered.map((opt) => {
                const isSelected = opt.value === value;
                return (
                  <button
                    key={opt.value}
                    type="button"
                    role="option"
                    aria-selected={isSelected}
                    onClick={() => {
                      onValueChange(opt.value);
                      setIsOpen(false);
                    }}
                    className={cn(
                      "flex w-full items-center justify-between gap-2 rounded-md px-2.5 py-2 text-left text-sm transition-colors",
                      isSelected ? "bg-primary text-primary-foreground" : "hover:bg-accent",
                    )}
                  >
                    <span className="min-w-0">
                      <span className="block truncate font-medium">{opt.label}</span>
                      {opt.subtitle ? (
                        <span
                          className={cn(
                            "block truncate text-xs",
                            isSelected ? "text-primary-foreground/80" : "text-muted-foreground",
                          )}
                        >
                          {opt.subtitle}
                        </span>
                      ) : null}
                    </span>
                    {isSelected ? <Check className="size-4 shrink-0" aria-hidden /> : null}
                  </button>
                );
              })
            )}
          </div>
        </div>
      ) : null}
    </div>
  );
}
