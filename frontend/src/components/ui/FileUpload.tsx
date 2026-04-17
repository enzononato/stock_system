import { useCallback, useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { Upload, X, FileText } from 'lucide-react'
import { cn } from '@/lib/utils'

interface FileUploadProps {
  accept?: Record<string, string[]>
  onFile: (file: File | null) => void
  label?: string
  className?: string
}

export function FileUpload({ accept = { 'application/pdf': ['.pdf'] }, onFile, label = 'Arraste ou clique para selecionar PDF', className }: FileUploadProps) {
  const [file, setFile] = useState<File | null>(null)

  const onDrop = useCallback((accepted: File[]) => {
    const f = accepted[0] ?? null
    setFile(f)
    onFile(f)
  }, [onFile])

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
        'flex flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed p-6 cursor-pointer transition-colors',
        isDragActive ? 'border-blue-400 bg-blue-50' : 'border-slate-200 hover:border-slate-400',
        className
      )}
    >
      <input {...getInputProps()} />
      {file ? (
        <div className="flex items-center gap-2 text-sm text-slate-700">
          <FileText size={20} className="text-blue-500" />
          <span className="max-w-xs truncate">{file.name}</span>
          <button onClick={clear} className="ml-1 text-slate-400 hover:text-red-500">
            <X size={16} />
          </button>
        </div>
      ) : (
        <>
          <Upload size={24} className="text-slate-400" />
          <p className="text-sm text-slate-500">{label}</p>
        </>
      )}
    </div>
  )
}
