import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listPeripherals, createPeripheral, deletePeripheral } from '@/api/peripherals'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { KpiCard } from '@/components/ui/KpiCard'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Badge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from '@/components/ui/alert-dialog'
import { useAuth } from '@/contexts/AuthContext'
import { useConstants } from '@/hooks/useConstants'
import type { ColumnDef } from '@tanstack/react-table'
import type { Peripheral } from '@/api/peripherals'
import { formatDate } from '@/lib/utils'
import { Trash2 } from 'lucide-react'

function PeripheralStatusBadge({ status }: { status?: string }) {
  if (status === 'Disponível') return <Badge variant="success">{status}</Badge>
  if (status === 'Em Uso') return <Badge variant="warning">{status}</Badge>
  if (status === 'Substituido') return <Badge variant="danger">{status}</Badge>
  return <Badge>{status ?? '-'}</Badge>
}

export default function PeripheralsPage() {
  const queryClient = useQueryClient()
  const { hasRole } = useAuth()
  const { peripheralTypes, isLoading: constantsLoading } = useConstants()
  const [tipo, setTipo] = useState('')
  const [brand, setBrand] = useState('')
  const [model, setModel] = useState('')
  const [identificador, setIdentificador] = useState('')

  const { data: peripherals = [], isLoading } = useQuery({
    queryKey: ['peripherals'],
    queryFn: () => listPeripherals(),
  })

  const mutation = useMutation({
    mutationFn: createPeripheral,
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      setTipo(''); setBrand(''); setModel(''); setIdentificador('')
      toast('Periférico cadastrado com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao cadastrar.'
      toast(msg, 'error')
    },
  })

  // T4: DELETE /api/peripherals/{id} (Gestor e Técnico) inativa um periférico
  // cadastrado por engano; o backend recusa com 400 se ele estiver em uso —
  // a mensagem exibida no toast de erro vem direto do `detail` da resposta.
  const deleteMutation = useMutation({
    mutationFn: (peripheralId: number) => deletePeripheral(peripheralId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['peripherals'] })
      toast('Periférico inativado com sucesso.')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao inativar periférico.'
      toast(msg, 'error')
    },
  })

  const columns: ColumnDef<Peripheral, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'model', header: 'Modelo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'identificador', header: 'Identificador (S/N)', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'status', header: 'Status', cell: ({ row }) => <PeripheralStatusBadge status={row.original.status} /> },
    { accessorKey: 'date_registered', header: 'Cadastro', cell: ({ getValue }) => formatDate(getValue() as string) },
    // Ação de inativar só aparece para quem o backend de fato autoriza
    // (Gestor e Técnico), mesmo padrão de gating usado em StockPage.
    ...(hasRole('Gestor', 'Técnico') ? [{
      id: 'actions',
      header: '',
      cell: ({ row }: { row: { original: Peripheral } }) => (
        <AlertDialog>
          <AlertDialogTrigger asChild>
            <Button size="sm" variant="ghost" className="text-red-500 hover:text-red-700" disabled={deleteMutation.isPending}>
              <Trash2 size={14} />
            </Button>
          </AlertDialogTrigger>
          <AlertDialogContent>
            <AlertDialogHeader>
              <AlertDialogTitle>Inativar periférico #{row.original.id}?</AlertDialogTitle>
              <AlertDialogDescription>
                {row.original.tipo} — {row.original.brand || '-'} {row.original.model || ''} (S/N: {row.original.identificador || '-'}).
                {' '}Periféricos em uso não podem ser inativados. Esta ação não pode ser desfeita.
              </AlertDialogDescription>
            </AlertDialogHeader>
            <AlertDialogFooter>
              <AlertDialogCancel>Cancelar</AlertDialogCancel>
              <AlertDialogAction onClick={() => deleteMutation.mutate(row.original.id)}>
                Inativar
              </AlertDialogAction>
            </AlertDialogFooter>
          </AlertDialogContent>
        </AlertDialog>
      ),
    } as ColumnDef<Peripheral, unknown>] : []),
  ]

  return (
    <div className="space-y-8">
      <div>
        <h2 className="text-xl font-semibold tracking-tight text-foreground">Periféricos</h2>
        <p className="text-sm text-muted-foreground">Cadastre e gerencie periféricos como mouse, teclado, monitor etc.</p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <KpiCard label="Total de Periféricos" value={peripherals.length} tone="primary" highlighted />
        <KpiCard label="Vinculados" value={peripherals.filter((p) => p.link_id).length} tone="success" />
        <KpiCard label="Disponíveis" value={peripherals.filter((p) => !p.link_id).length} tone="info" />
      </div>

      {/* Formulário de cadastro */}
      <form
        onSubmit={(e) => { e.preventDefault(); mutation.mutate({ tipo, brand, model, identificador }) }}
        className="bg-card rounded-xl border p-6 space-y-4"
      >
        <h3 className="font-medium text-foreground">Cadastrar Periférico</h3>
        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Tipo *</Label>
            <Select value={tipo} onValueChange={setTipo} required disabled={constantsLoading}>
              <SelectTrigger><SelectValue placeholder="Selecione o tipo" /></SelectTrigger>
              <SelectContent>
                {peripheralTypes.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
              </SelectContent>
            </Select>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Identificador (S/N) *</Label>
            <Input value={identificador} onChange={e => setIdentificador(e.target.value)} required placeholder="Número de série único" />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Marca</Label>
            <Input value={brand} onChange={e => setBrand(e.target.value)} />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Modelo</Label>
            <Input value={model} onChange={e => setModel(e.target.value)} />
          </div>
        </div>
        <div className="flex justify-end">
          <Button type="submit" disabled={mutation.isPending}>
            {mutation.isPending ? 'Cadastrando...' : 'Cadastrar Periférico'}
          </Button>
        </div>
      </form>

      {/* Lista */}
      {isLoading ? (
        <div className="py-8 text-center text-muted-foreground">Carregando...</div>
      ) : (
        <DataTable data={peripherals} columns={columns} searchPlaceholder="Buscar periférico..." />
      )}
    </div>
  )
}
