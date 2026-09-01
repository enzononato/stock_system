import { useState, type ReactNode } from "react";
import { useMutation, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { CheckCircle2, FileDown, Loader2, ShieldAlert } from "lucide-react";

import { confirmLoan, generateLoanTerm } from "@/api/loans";
import { getErrorMessage } from "@/lib/api-error";
import { Button } from "@/components/ui/button";
import { FileUpload } from "@/components/app/FileUpload";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

/**
 * Gera o termo de empréstimo (.docx) e dispara o download no navegador.
 * Compartilhada por LoanPage (botão "Gerar Termo" por linha + painel de
 * confirmação) e TermsPage (mesma ação nas duas telas).
 */
export async function generateAndDownloadLoanTerm(itemId: number): Promise<void> {
  try {
    const blob = await generateLoanTerm(itemId);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `termo_emprestimo_${itemId}.docx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err: unknown) {
    let msg = "Erro ao gerar termo.";
    const axiosErr = err as { response?: { data?: unknown } };
    if (axiosErr.response?.data instanceof Blob) {
      try {
        const text = await axiosErr.response.data.text();
        const json = JSON.parse(text) as { detail?: string };
        if (json.detail) msg = json.detail;
      } catch {
        // fallback para mensagem padrão
      }
    } else {
      msg = getErrorMessage(err, msg);
    }
    toast.error(msg);
  }
}

interface ConfirmacaoTermoProps {
  /** ID do item cujo empréstimo está sendo confirmado. */
  itemId: number;
  /** Texto/instruções exibidos entre o título e a área de upload. */
  description: ReactNode;
  /** Mostra, dentro do próprio painel, um botão para gerar o termo antes do upload. */
  showGenerateButton?: boolean;
  /** Rótulo da área de upload. */
  uploadLabel?: string;
  /** Mensagem de erro usada quando o backend não retorna `detail`. */
  errorMessage?: string;
  /** Chamado após a confirmação ter sucesso. */
  onConfirmed: () => void;
  /** Quando informado, exibe um botão "Cancelar" que o chama. */
  onCancel?: (() => void) | undefined;
}

/**
 * Modal de confirmação de empréstimo: upload do termo assinado + confirmação.
 * Painel compartilhado por LoanPage e TermsPage.
 */
export function ConfirmacaoTermo({
  itemId,
  description,
  showGenerateButton = false,
  uploadLabel = "Upload do termo assinado (PDF)",
  errorMessage = "Erro ao confirmar empréstimo.",
  onConfirmed,
  onCancel,
}: ConfirmacaoTermoProps) {
  const queryClient = useQueryClient();
  const [signedPdf, setSignedPdf] = useState<File | null>(null);
  const close = onCancel ?? onConfirmed;

  const confirmMutation = useMutation({
    mutationFn: (pdf: File) => confirmLoan(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["items"] });
      setSignedPdf(null);
      toast.success("Empréstimo confirmado com sucesso!");
      onConfirmed();
    },
    onError: (err: unknown) => {
      toast.error(getErrorMessage(err, errorMessage));
    },
  });

  return (
    <Dialog open onOpenChange={(open) => !open && close()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <ShieldAlert className="size-4 text-warning" aria-hidden />
            Confirmar empréstimo — item #{itemId}
          </DialogTitle>
        </DialogHeader>
        <div className="space-y-4">
          <div className="text-sm text-muted-foreground">{description}</div>
          {showGenerateButton && (
            <Button variant="outline" size="sm" onClick={() => generateAndDownloadLoanTerm(itemId)}>
              <FileDown className="mr-2 size-3.5" aria-hidden />
              Gerar termo
            </Button>
          )}
          <FileUpload onFile={setSignedPdf} label={uploadLabel} />
          <div className="flex gap-3 pt-1">
            <Button
              disabled={!signedPdf || confirmMutation.isPending}
              onClick={() => signedPdf && confirmMutation.mutate(signedPdf)}
              className="flex-1"
            >
              {confirmMutation.isPending ? (
                <Loader2 className="mr-2 size-3.5 animate-spin" aria-hidden />
              ) : (
                <CheckCircle2 className="mr-2 size-3.5" aria-hidden />
              )}
              {confirmMutation.isPending ? "Confirmando…" : "Confirmar empréstimo"}
            </Button>
            <Button variant="ghost" onClick={close} disabled={confirmMutation.isPending}>
              Cancelar
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
