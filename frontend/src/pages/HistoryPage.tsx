import { useEffect, useState } from 'react'
import { useQuery, useMutation, useQueryClient, keepPreviousData } from '@tanstack/react-query'
import { listHistoryPaginated, reverseEntryWithPassword, type HistoryEntry } from '@/api/history'
import { downloadAuthenticated } from '@/api/client'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import { getErrorMessage } from '@/lib/api-error'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { useAuth } from '@/contexts/AuthContext'
import type { ColumnDef } from '@tanstack/react-table'
import { formatCpf, formatDateTime } from '@/lib/utils'
import { RotateCcw, Paperclip, Loader2 } from 'lucide-react'

const REVERSIBLE_OPS = ['Cadastro','Empréstimo','Confirmação Empréstimo','Devolução','Confirmação Devolução']

const PAGE_SIZE = 20

function OperationBadge({ op }: { op?: string }) {
  if (!op) return <Badge>-</Badge>
  // Monocromático: só as operações destrutivas (Exclusão, Estorno) carregam
  // cor de sinal (danger). Todas as demais são neutras (default).
  if (op === 'Exclusão' || op === 'Estorno') return <Badge variant="danger">{op}</Badge>
  return <Badge variant="default">{op}</Badge>
}

/**
 * Deriva um nome de arquivo legível para download a partir da chave de storage
 * ("categoria/arquivo.ext" — ver `app/core/storage.py` no backend). O trecho
 * após a primeira barra já é o nome original enviado pelo usuário (prefixado
 * pelo backend com contexto, ex.: "remocao_12_nota.pdf"), então não há por que
 * desmontar o padrão — só extrair esse trecho.
 */
function attachmentFilename(key: string): string {
  const idx = key.indexOf('/')
  return idx >= 0 ? key.slice(idx + 1) : key
}

interface AttachmentDescriptor {
  key: string
  label: string
}

function AttachmentCell({
  entry,
  downloadingKey,
  onDownload,
}: {
  entry: HistoryEntry
  downloadingKey: string | null
  onDownload: (key: string) => void
}) {
  const attachments: AttachmentDescriptor[] = []
  if (entry.operacao_anexo) attachments.push({ key: entry.operacao_anexo, label: 'Comprovante' })
  if (entry.termo_assinado_anexo) attachments.push({ key: entry.termo_assinado_anexo, label: 'Termo' })

  if (attachments.length === 0) return <span className="text-muted-foreground">-</span>

  return (
    <div className="flex flex-col items-start gap-1">
      {attachments.map(({ key, label }) => (
        <Button
          key={key}
          type="button"
          size="sm"
          variant="ghost"
          className="h-7 px-2 text-muted-foreground hover:text-foreground"
          disabled={downloadingKey === key}
          onClick={() => onDownload(key)}
        >
          <Paperclip size={12} />
          {downloadingKey === key ? 'Baixando...' : label}
        </Button>
      ))}
    </div>
  )
}

export default function HistoryPage() {
  const { hasRole } = useAuth()
  const queryClient = useQueryClient()

  const [pageIndex, setPageIndex] = useState(0)
  // `search` é o que o usuário digita; `searchAplicado` é o que vai para a
  // query, com atraso, para não disparar uma requisição por tecla.
  const [search, setSearch] = useState('')
  const [searchAplicado, setSearchAplicado] = useState('')
  const [downloadingKey, setDownloadingKey] = useState<string | null>(null)

  // Confirmação de estorno: entrada sendo estornada (null = painel fechado),
  // senha digitada e mensagem de erro da última tentativa (ex.: 403).
  const [reversingEntry, setReversingEntry] = useState<HistoryEntry | null>(null)
  const [password, setPassword] = useState('')
  const [reverseError, setReverseError] = useState<string | null>(null)

  const { data, isLoading } = useQuery({
    queryKey: ['history', pageIndex, searchAplicado],
    queryFn: () =>
      listHistoryPaginated({
        search: searchAplicado || undefined,
        limit: PAGE_SIZE,
        offset: pageIndex * PAGE_SIZE,
      }),
    placeholderData: keepPreviousData,
  })

  const history = data?.items ?? []
  const total = data?.total ?? 0

  // Se um estorno (ou qualquer outra invalidação) reduzir o total de
  // registros a ponto da página atual deixar de existir (ex.: estornar o
  // único item da última página), volta para a última página válida em vez
  // de deixar a tabela "presa" numa página vazia.
  useEffect(() => {
    if (!data || data.total <= 0) return
    const maxPageIndex = Math.max(0, Math.ceil(data.total / PAGE_SIZE) - 1)
    if (pageIndex > maxPageIndex) setPageIndex(maxPageIndex)
  }, [data, pageIndex])

  // Espera o usuário parar de digitar antes de consultar o servidor, e volta
  // para a primeira página: o resultado da nova busca não tem relação com a
  // página em que ele estava.
  useEffect(() => {
    const timer = setTimeout(() => {
      setSearchAplicado(search.trim())
      setPageIndex(0)
    }, 400)
    return () => clearTimeout(timer)
  }, [search])

  function openReverseConfirm(entry: HistoryEntry) {
    setReversingEntry(entry)
    setPassword('')
    setReverseError(null)
  }

  function closeReverseConfirm() {
    setReversingEntry(null)
    setPassword('')
    setReverseError(null)
  }

  const reverseMutation = useMutation({
    mutationFn: ({ id, password }: { id: number; password: string }) => reverseEntryWithPassword(id, password),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['history'] })
      queryClient.invalidateQueries({ queryKey: ['items'] })
      toast('Operação estornada com sucesso!')
      closeReverseConfirm()
    },
    onError: (err: unknown) => {
      const status = (err as { response?: { status?: number } })?.response?.status
      // 403 é a resposta específica de senha incorreta (ver T4): mensagem fixa
      // e clara, sem fechar o painel, para o usuário poder tentar de novo.
      const msg = status === 403 ? 'Senha incorreta. Ação não autorizada.' : getErrorMessage(err, 'Erro ao estornar.')
      setReverseError(msg)
      toast(msg, 'error')
    },
  })

  function handleConfirmReverse() {
    if (!reversingEntry || !password) return
    reverseMutation.mutate({ id: reversingEntry.id, password })
  }

  async function handleDownloadAttachment(key: string) {
    setDownloadingKey(key)
    try {
      await downloadAuthenticated(`/api/documents/files/${key}`, attachmentFilename(key))
    } catch (err) {
      toast(getErrorMessage(err, 'Erro ao baixar anexo.'), 'error')
    } finally {
      setDownloadingKey(null)
    }
  }

  const columns: ColumnDef<HistoryEntry, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
    { accessorKey: 'item_id', header: 'Item', size: 60, cell: ({ getValue }) => getValue() as number ?? '-' },
    { accessorKey: 'peripheral_id', header: 'Periférico', size: 80, cell: ({ getValue }) => getValue() as number ?? '-' },
    { accessorKey: 'operador', header: 'Operador' },
    { accessorKey: 'operation', header: 'Operação', cell: ({ row }) => <OperationBadge op={row.original.operation} /> },
    { accessorKey: 'tipo', header: 'Tipo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'marca', header: 'Marca', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'modelo', header: 'Modelo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'identificador', header: 'Identificador', cell: ({ getValue }) => <span className="num">{(getValue() as string) || '-'}</span> },
    { accessorKey: 'nota_fiscal', header: 'Nota Fiscal', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'usuario', header: 'Usuário', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'cpf', header: 'CPF', cell: ({ getValue }) => <span className="num">{formatCpf(getValue() as string)}</span> },
    { accessorKey: 'cargo', header: 'Cargo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'setor', header: 'Setor', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'revenda', header: 'Revenda', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'data_operacao', header: 'Data', cell: ({ getValue }) => <span className="num">{formatDateTime(getValue() as string)}</span> },
    { accessorKey: 'details', header: 'Detalhes', cell: ({ getValue }) => getValue() as string || '-' },
    {
      id: 'anexo',
      header: 'Anexo',
      cell: ({ row }) => (
        <AttachmentCell entry={row.original} downloadingKey={downloadingKey} onDownload={handleDownloadAttachment} />
      ),
    },
    ...(hasRole('Gestor') ? [{
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: HistoryEntry } }) =>
        REVERSIBLE_OPS.includes(row.original.operation ?? '') ? (
          <Button
            size="sm"
            variant="ghost"
            className="text-muted-foreground hover:text-destructive"
            onClick={() => openReverseConfirm(row.original)}
          >
            <RotateCcw size={14} />Estornar
          </Button>
        ) : null,
    } as ColumnDef<HistoryEntry, unknown>] : []),
  ]

  return (
    <div className="space-y-4">
      <PageHeader
        eyebrow="Auditoria e Rastreabilidade"
        eyebrowDetail="Linha Temporal"
        title="Histórico"
        description="Registro completo de todas as operações."
      />

      {/* Confirmação de estorno — exige a senha do operador logado (T4). Alto
          contraste proposital: essa ação reescreve o histórico registrado. */}
      {reversingEntry && (
        <div className="figure-ground-panel space-y-4">
          <PanelHeader title="Confirmação de Estorno" />
          <h3 className="text-heading-sm text-foreground">Confirmar Estorno — Operação #<span className="num">{reversingEntry.id}</span></h3>
          <p className="text-body-sm text-muted-foreground">
            Isso desfará a operação <strong>&quot;{reversingEntry.operation ?? '-'}&quot;</strong>
            {reversingEntry.item_id != null && ` do item #${reversingEntry.item_id}`}
            {reversingEntry.peripheral_id != null && ` do periférico #${reversingEntry.peripheral_id}`}
            {reversingEntry.operador && ` (operador: ${reversingEntry.operador})`}. Esta ação não pode ser desfeita.
          </p>
          <div className="flex flex-col gap-1.5 max-w-xs">
            <Label htmlFor="senha-estorno">Confirme sua senha</Label>
            <Input
              id="senha-estorno"
              type="password"
              value={password}
              onChange={(e) => { setPassword(e.target.value); setReverseError(null) }}
              placeholder="Sua senha de acesso"
              autoFocus
            />
          </div>
          {reverseError && <p className="text-body-sm text-destructive">{reverseError}</p>}
          <div className="flex gap-3">
            <Button
              variant="destructive"
              disabled={!password || reverseMutation.isPending}
              onClick={handleConfirmReverse}
            >
              {reverseMutation.isPending ? 'Estornando...' : 'Confirmar'}
            </Button>
            <Button variant="ghost" onClick={closeReverseConfirm}>Cancelar</Button>
          </div>
        </div>
      )}

      {isLoading ? (
        <div className="py-8 flex items-center justify-center gap-2 text-body-sm text-muted-foreground">
          <Loader2 className="h-4 w-4 animate-spin" />
          Carregando...
        </div>
      ) : (
        <div className="space-y-3">
          <PanelHeader
            title="Consulta de Operações"
            description="Busque por operador, usuário, tipo, marca ou identificador."
          />
          <DataTable
            data={history}
            columns={columns}
            searchPlaceholder="Buscar por operador, usuário, tipo, marca, identificador..."
            pagination={{
              total,
              pageIndex,
              pageSize: PAGE_SIZE,
              onPageChange: setPageIndex,
              search,
              onSearchChange: setSearch,
            }}
          />
        </div>
      )}
    </div>
  )
}
