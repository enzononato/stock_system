import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { confirmLoan, generateLoanTerm } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { cn } from '@/lib/utils'
import { FileDown, CheckCircle } from 'lucide-react'

/**
 * Gera o termo de empréstimo (.docx) e dispara o download no navegador.
 *
 * Extraída (T5) de LoanPage.tsx (`handleGenerateTerm`) e TermsPage.tsx
 * (`handleDownloadTerm`), que tinham exatamente a mesma implementação sob
 * nomes diferentes. Fica aqui, junto de `ConfirmacaoTermo`, porque ambas as
 * telas também usam essa função fora do painel de confirmação (botão
 * "Gerar Termo" por linha da tabela de pendentes).
 */
export async function generateAndDownloadLoanTerm(itemId: number): Promise<void> {
  try {
    const blob = await generateLoanTerm(itemId)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `termo_emprestimo_${itemId}.docx`
    a.click()
    URL.revokeObjectURL(url)
  } catch {
    toast('Erro ao gerar termo.', 'error')
  }
}

interface ConfirmacaoTermoProps {
  /** ID do item cujo empréstimo está sendo confirmado. */
  itemId: number
  /** Texto/instruções exibidos entre o título e a área de upload. */
  description: React.ReactNode
  /** Esquema de cor do painel — cada tela preserva a identidade visual que já tinha. */
  variant?: 'amber' | 'blue'
  /** Rótulo da área de upload (FileUpload). */
  uploadLabel?: string
  /** Mostra, dentro do próprio painel, um botão para gerar o termo antes do upload. */
  showGenerateButton?: boolean
  /** Mensagem de erro usada quando o backend não retorna `detail`. */
  errorMessage?: string
  /** Chamado após a confirmação ter sucesso — cada tela zera seu próprio estado de "item pendente". */
  onConfirmed: () => void
  /** Quando informado, exibe um botão "Cancelar" que o chama. */
  onCancel?: () => void
}

const VARIANT_CLASSES: Record<'amber' | 'blue', { container: string; heading: string; text: string }> = {
  amber: { container: 'bg-amber-50 border-amber-200', heading: 'text-amber-800', text: 'text-amber-700' },
  blue: { container: 'bg-blue-50 border-blue-200', heading: 'text-blue-800', text: 'text-blue-700' },
}

/**
 * Painel de confirmação de empréstimo: upload do termo assinado + confirmação.
 *
 * Unifica (T5) o fluxo que LoanPage.tsx e TermsPage.tsx implementavam em
 * duplicata — dois estados (`signedPdf`), duas mutations idênticas
 * (`confirmLoan`) e dois blocos de JSX quase iguais que precisavam ser
 * mantidos manualmente em sincronia. O texto, o rótulo de upload, a cor do
 * painel, a mensagem de erro e a presença dos botões "Gerar Termo"/"Cancelar"
 * continuam configuráveis por tela para preservar o comportamento atual de
 * cada uma.
 */
export function ConfirmacaoTermo({
  itemId,
  description,
  variant = 'amber',
  uploadLabel = 'Upload do Termo Assinado (PDF)',
  showGenerateButton = false,
  errorMessage = 'Erro ao confirmar empréstimo.',
  onConfirmed,
  onCancel,
}: ConfirmacaoTermoProps) {
  const queryClient = useQueryClient()
  const [signedPdf, setSignedPdf] = useState<File | null>(null)
  const colors = VARIANT_CLASSES[variant]

  const confirmMutation = useMutation({
    mutationFn: (pdf: File) => confirmLoan(itemId, pdf),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] })
      setSignedPdf(null)
      toast('Empréstimo confirmado com sucesso!')
      onConfirmed()
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? errorMessage
      toast(msg, 'error')
    },
  })

  return (
    <div className={cn('rounded-xl border p-6 space-y-4', colors.container)}>
      <h3 className={cn('font-medium', colors.heading)}>Confirmar Empréstimo — Item #{itemId}</h3>
      <div className={cn('text-sm', colors.text)}>{description}</div>
      {showGenerateButton && (
        <Button variant="outline" onClick={() => generateAndDownloadLoanTerm(itemId)}>
          <FileDown size={14} className="mr-2" />Gerar Termo
        </Button>
      )}
      <FileUpload onFile={setSignedPdf} label={uploadLabel} />
      <div className="flex gap-3">
        <Button
          disabled={!signedPdf || confirmMutation.isPending}
          onClick={() => signedPdf && confirmMutation.mutate(signedPdf)}
        >
          <CheckCircle size={14} className="mr-2" />
          {confirmMutation.isPending ? 'Confirmando...' : 'Confirmar Empréstimo'}
        </Button>
        {onCancel && (
          <Button variant="ghost" onClick={onCancel}>
            Cancelar
          </Button>
        )}
      </div>
    </div>
  )
}
