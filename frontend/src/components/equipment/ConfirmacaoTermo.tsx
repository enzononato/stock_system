import { useState } from 'react'
import { useMutation, useQueryClient } from '@tanstack/react-query'
import { confirmLoan, generateLoanTerm } from '@/api/loans'
import { Button } from '@/components/ui/button'
import { FileUpload } from '@/components/ui/FileUpload'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
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
  } catch (err) {
    // A resposta de erro chega como Blob (não como JSON) porque a requisição
    // usa responseType 'blob' para o caminho de sucesso (o .docx). Precisa
    // ser desembrulhada manualmente antes de extrair o `detail` — é o único
    // caso que getErrorMessage não cobre, então o resultado dela só entra
    // como mensagem padrão quando esse desembrulho não encontra nada.
    let msg: string | undefined
    const data = (err as { response?: { data?: unknown } })?.response?.data
    if (data instanceof Blob) {
      try {
        const text = await data.text()
        const json = JSON.parse(text)
        if (json.detail) msg = json.detail
      } catch {
        // fallback para mensagem padrão
      }
    }
    toast(msg ?? getErrorMessage(err, 'Erro ao gerar termo.'), 'error')
  }
}

interface ConfirmacaoTermoProps {
  /** ID do item cujo empréstimo está sendo confirmado. */
  itemId: number
  /** Texto/instruções exibidos entre o título e a área de upload. */
  description: React.ReactNode
  /**
   * Antigo esquema de cor do painel (âmbar para empréstimo, azul para
   * devolução). Interface monocromática: as duas variantes hoje renderizam
   * o mesmo tratamento neutro (ver `VARIANT_CLASSES`). Prop mantida apenas
   * para não quebrar as chamadas existentes em LoanPage/TermsPage.
   */
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

// Interface monocromática: os dois esquemas de cor (âmbar para empréstimo,
// azul para devolução) colapsam no mesmo tratamento neutro — como o variant
// "gradient" do Button, que também renderiza como default hoje. O prop
// `variant` continua aceito para não quebrar quem já o passa (LoanPage,
// TermsPage), mas não produz mais diferença visual entre as duas telas.
const VARIANT_CLASSES: Record<'amber' | 'blue', { container: string; heading: string; text: string }> = {
  amber: { container: 'figure-ground-panel', heading: 'text-heading-sm text-foreground', text: 'text-body-sm text-muted-foreground' },
  blue: { container: 'figure-ground-panel', heading: 'text-heading-sm text-foreground', text: 'text-body-sm text-muted-foreground' },
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
      toast(getErrorMessage(err, errorMessage), 'error')
    },
  })

  return (
    <div className={cn('space-y-4', colors.container)}>
      <h3 className={colors.heading}>
        Confirmar Empréstimo — Item #<span className="num">{itemId}</span>
      </h3>
      <div className={colors.text}>{description}</div>
      {showGenerateButton && (
        <Button variant="outline" onClick={() => generateAndDownloadLoanTerm(itemId)}>
          <FileDown size={14} />Gerar Termo
        </Button>
      )}
      <FileUpload onFile={setSignedPdf} label={uploadLabel} />
      <div className="flex gap-3">
        <Button
          disabled={!signedPdf || confirmMutation.isPending}
          onClick={() => signedPdf && confirmMutation.mutate(signedPdf)}
        >
          <CheckCircle size={14} />
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
