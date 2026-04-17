import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listPeripherals, createPeripheral } from '@/api/peripherals'
import { DataTable } from '@/components/ui/DataTable'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Badge } from '@/components/ui/badge'
import { toast } from '@/components/ui/toast'
import type { ColumnDef } from '@tanstack/react-table'
import type { Peripheral } from '@/api/peripherals'
import { formatDate } from '@/lib/utils'

const PERIPHERAL_TYPES = ['Mouse','Teclado','Monitor','Mousepad','Suporte','Estabilizador','Mochila','Fone','Webcam','Adaptador','Outro']

function PeripheralStatusBadge({ status }: { status?: string }) {
  if (status === 'Disponível') return <Badge variant="success">{status}</Badge>
  if (status === 'Em Uso') return <Badge variant="warning">{status}</Badge>
  if (status === 'Substituido') return <Badge variant="danger">{status}</Badge>
  return <Badge>{status ?? '-'}</Badge>
}

export default function PeripheralsPage() {
  const queryClient = useQueryClient()
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

  const columns: ColumnDef<Peripheral, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'tipo', header: 'Tipo' },
    { accessorKey: 'brand', header: 'Marca', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'model', header: 'Modelo', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'identificador', header: 'Identificador (S/N)', cell: ({ getValue }) => getValue() as string || '-' },
    { accessorKey: 'status', header: 'Status', cell: ({ row }) => <PeripheralStatusBadge status={row.original.status} /> },
    { accessorKey: 'date_registered', header: 'Cadastro', cell: ({ getValue }) => formatDate(getValue() as string) },
  ]

  return (
    <div className="space-y-8">
      <div>
        <h2 className="text-xl font-semibold">Periféricos</h2>
        <p className="text-sm text-slate-500">Cadastre e gerencie periféricos como mouse, teclado, monitor etc.</p>
      </div>

      {/* Formulário de cadastro */}
      <form
        onSubmit={(e) => { e.preventDefault(); mutation.mutate({ tipo, brand, model, identificador }) }}
        className="bg-white rounded-xl border p-6 space-y-4"
      >
        <h3 className="font-medium text-slate-700">Cadastrar Periférico</h3>
        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Tipo *</Label>
            <Select value={tipo} onValueChange={setTipo} required>
              <SelectTrigger><SelectValue placeholder="Selecione o tipo" /></SelectTrigger>
              <SelectContent>
                {PERIPHERAL_TYPES.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
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
        <div className="py-8 text-center text-slate-400">Carregando...</div>
      ) : (
        <DataTable data={peripherals} columns={columns} searchPlaceholder="Buscar periférico..." />
      )}
    </div>
  )
}
