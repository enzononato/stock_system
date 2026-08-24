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
        'group relative flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed p-7 cursor-pointer transition-all duration-200 backdrop-blur-sm select-none',
        isDragActive
          ? 'border-primary bg-primary/100/10 shadow-glow-indigo scale-[1.01]'
          : file
          ? 'border-emerald-400/80 bg-emerald-50/50'
          : 'border-border/80 bg-muted/50 hover:border-primary hover:bg-primary/10/30',
        className
      )}
    >
      <input {...getInputProps()} />
      {file ? (
        <div className="flex items-center gap-3 px-4 py-2 rounded-xl bg-popover shadow-sm border border-emerald-200 animate-scale-in">
          <div className="h-8 w-8 rounded-lg bg-emerald-100 text-emerald-600 flex items-center justify-center">
            <CheckCircle2 size={18} />
          </div>
          <div className="flex flex-col min-w-0">
            <span className="text-xs font-bold text-foreground truncate max-w-xs">{file.name}</span>
            <span className="text-[10px] font-semibold text-emerald-600">Arquivo Pronto</span>
          </div>
          <button
            onClick={clear}
            className="ml-2 text-muted-foreground hover:text-rose-500 p-1 rounded-lg hover:bg-secondary transition-colors"
            title="Remover arquivo"
          >
            <X size={16} />
          </button>
        </div>
      ) : (
        <>
          <div className="h-12 w-12 rounded-2xl bg-primary/10 text-primary flex items-center justify-center group-hover:scale-110 transition-transform shadow-sm">
            <Upload size={22} />
          </div>
          <div className="text-center space-y-1">
            <p className="text-xs font-bold text-foreground">{label}</p>
            <p className="text-[11px] font-medium text-muted-foreground">Suporta apenas documentos em formato PDF</p>
          </div>
        </>
      )}
    </div>
  )
}

