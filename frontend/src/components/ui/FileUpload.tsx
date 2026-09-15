import { useCallback, useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { Upload, X, FileText, CheckCircle2 } from 'lucide-react'
import { cn } from '@/lib/utils'

interface FileUploadProps {
  accept?: Record<string, string[]>
  onFile: (file: File | null) => void
  label?: string
  className?: string
}

// Deriva a mensagem de formatos aceitos a partir do que o chamador de fato
// passou em `accept` — Remoção e Vínculo passam PDF + imagem, e a mensagem
// fixa antiga ("apenas PDF") mentia para esses dois fluxos.
function describeAcceptedTypes(accept: Record<string, string[]>): string {
  const mimeTypes = Object.keys(accept)
  const hasPdf = mimeTypes.some(m => m === 'application/pdf')
  const hasImage = mimeTypes.some(m => m.startsWith('image/'))
  const outrasExtensoes = mimeTypes
    .filter(m => m !== 'application/pdf' && !m.startsWith('image/'))
    .flatMap(m => accept[m])

  const partes: string[] = []
  if (hasPdf) partes.push('documentos em formato PDF')
  if (hasImage) partes.push('imagens')
  if (outrasExtensoes.length > 0) partes.push(`arquivos ${outrasExtensoes.join(', ')}`)

  if (partes.length === 0) return 'Selecione um arquivo'
  if (partes.length === 1) return `Suporta apenas ${partes[0]}`
  const ultima = partes.pop()
  return `Suporta ${partes.join(', ')} ou ${ultima}`
}

export function FileUpload({
  accept = { 'application/pdf': ['.pdf'] },
  onFile,
  label = 'Arraste ou clique para selecionar PDF assinado',
  className,
}: FileUploadProps) {
  const [file, setFile] = useState<File | null>(null)

  const onDrop = useCallback(
    (accepted: File[]) => {
      const f = accepted[0] ?? null
      setFile(f)
      onFile(f)
    },
    [onFile]
  )

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop, accept, maxFiles: 1 })

  function clear(e: React.MouseEvent) {
    e.stopPropagation()
    setFile(null)
    onFile(null)
  }

  return (
    <div
      {...getRootProps()}
      className={cn(
        'group relative flex flex-col items-center justify-center gap-3 rounded-md border-2 border-dashed p-7 cursor-pointer transition-colors duration-micro select-none',
        isDragActive || file
          ? 'border-border-strong bg-surface-alt'
          : 'border-border bg-surface hover:border-border-strong hover:bg-surface-alt',
        className
      )}
    >
      <input {...getInputProps()} />
      {file ? (
        <div className="flex items-center gap-3 px-4 py-2 rounded-md border border-border bg-surface">
          <div className="h-8 w-8 rounded text-foreground flex items-center justify-center">
            <CheckCircle2 size={18} />
          </div>
          <div className="flex flex-col min-w-0">
            <span className="text-xs font-semibold text-foreground truncate max-w-xs">{file.name}</span>
            <span className="text-[10px] font-semibold text-muted-foreground">Arquivo Pronto</span>
          </div>
          <button
            onClick={clear}
            className="ml-2 text-muted-foreground hover:text-destructive p-1 rounded hover:bg-surface-alt transition-colors duration-micro"
            title="Remover arquivo"
          >
            <X size={16} />
          </button>
        </div>
      ) : (
        <>
          <div className="h-12 w-12 rounded-md bg-surface-alt text-foreground flex items-center justify-center">
            <Upload size={22} />
          </div>
          <div className="text-center space-y-1">
            <p className="text-xs font-semibold text-foreground">{label}</p>
            <p className="text-[11px] font-medium text-muted-foreground">{describeAcceptedTypes(accept)}</p>
          </div>
        </>
      )}
    </div>
  )
}

