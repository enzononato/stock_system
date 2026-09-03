import { useRef, useState } from "react";
import { CheckCircle2, FileText, Upload, X } from "lucide-react";

import { cn } from "@/lib/utils";

interface FileUploadProps {
  /** Lista de extensões/mime aceitos (atributo `accept` do input). */
  accept?: string | undefined;
  onFile: (file: File | null) => void;
  label?: string | undefined;
  hint?: string | undefined;
  className?: string | undefined;
  id?: string | undefined;
}

/**
 * Upload por clique ou arrastar-e-soltar, sem dependência externa. Usado nos
 * fluxos que exigem anexo (termo assinado, comprovante de remoção).
 */
export function FileUpload({
  accept = "application/pdf",
  onFile,
  label = "Arraste ou clique para selecionar o PDF assinado",
  hint,
  className,
  id,
}: FileUploadProps) {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  function select(next: File | null) {
    setFile(next);
    onFile(next);
  }

  return (
    <div
      className={cn(
        "relative flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed border-border bg-muted/30 p-6 text-center transition-colors",
        isDragging && "border-primary bg-primary/5",
        className,
      )}
      onDragOver={(e) => {
        e.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setIsDragging(false);
        const dropped = e.dataTransfer.files?.[0] ?? null;
        if (dropped) select(dropped);
      }}
    >
      <input
        ref={inputRef}
        id={id}
        type="file"
        accept={accept}
        className="sr-only"
        onChange={(e) => select(e.target.files?.[0] ?? null)}
      />

      {file ? (
        <div className="flex w-full items-center gap-3 text-left">
          <CheckCircle2 className="size-5 shrink-0 text-primary" aria-hidden />
          <div className="min-w-0 flex-1">
            <p className="truncate text-sm font-medium text-foreground">{file.name}</p>
            <p className="text-xs text-muted-foreground">{(file.size / 1024).toFixed(0)} KB</p>
          </div>
          <button
            type="button"
            onClick={() => {
              select(null);
              if (inputRef.current) inputRef.current.value = "";
            }}
            aria-label="Remover arquivo selecionado"
            className="rounded-md p-1.5 text-muted-foreground hover:bg-accent hover:text-foreground"
          >
            <X className="size-4" aria-hidden />
          </button>
        </div>
      ) : (
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="flex flex-col items-center gap-2 rounded-md px-2 py-1 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring"
        >
          <span
            className="grid size-10 place-items-center rounded-full bg-muted text-muted-foreground"
            aria-hidden
          >
            {isDragging ? <Upload className="size-5" /> : <FileText className="size-5" />}
          </span>
          <span className="text-sm font-medium text-foreground">{label}</span>
          {hint ? <span className="text-xs text-muted-foreground">{hint}</span> : null}
        </button>
      )}
    </div>
  );
}
