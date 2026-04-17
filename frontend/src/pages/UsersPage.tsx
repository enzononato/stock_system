import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listUsers, createUser, removeUser, updatePassword, type User } from '@/api/users'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Badge } from '@/components/ui/badge'
import { DataTable } from '@/components/ui/DataTable'
import { toast } from '@/components/ui/toast'
import type { ColumnDef } from '@tanstack/react-table'
import { UserPlus, Trash2, Key } from 'lucide-react'

const ROLES = ['Gestor','Técnico','Jovem Aprendiz']

export default function UsersPage() {
  const queryClient = useQueryClient()
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [role, setRole] = useState('')
  const [changingPasswordId, setChangingPasswordId] = useState<number | null>(null)
  const [newPassword, setNewPassword] = useState('')

  const { data: users = [] } = useQuery({ queryKey: ['users'], queryFn: listUsers })

  const createMutation = useMutation({
    mutationFn: () => createUser({ username, password, role }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['users'] })
      setUsername(''); setPassword(''); setRole('')
      toast('Usuário criado com sucesso!')
    },
    onError: (err: unknown) => {
      const msg = (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? 'Erro ao criar usuário.'
      toast(msg, 'error')
    },
  })

  const removeMutation = useMutation({
    mutationFn: (id: number) => removeUser(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['users'] })
      toast('Usuário removido.')
    },
    onError: () => toast('Erro ao remover usuário.', 'error'),
  })

  const passwordMutation = useMutation({
    mutationFn: ({ id, pass }: { id: number; pass: string }) => updatePassword(id, pass),
    onSuccess: () => {
      setChangingPasswordId(null); setNewPassword('')
      toast('Senha alterada com sucesso!')
    },
    onError: () => toast('Erro ao alterar senha.', 'error'),
  })

  const columns: ColumnDef<User, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60 },
    { accessorKey: 'username', header: 'Usuário' },
    {
      accessorKey: 'role',
      header: 'Função',
      cell: ({ getValue }) => {
        const r = getValue() as string
        const v = r === 'Gestor' ? 'info' : r === 'Técnico' ? 'warning' : 'success'
        return <Badge variant={v as 'info' | 'warning' | 'success'}>{r}</Badge>
      },
    },
    {
      id: 'actions',
      header: '',
      cell: ({ row }) => (
        <div className="flex gap-2">
          <Button size="sm" variant="ghost" onClick={() => setChangingPasswordId(row.original.id)}>
            <Key size={14} />
          </Button>
          <Button
            size="sm"
            variant="ghost"
            className="text-red-500 hover:text-red-700"
            onClick={() => removeMutation.mutate(row.original.id)}
          >
            <Trash2 size={14} />
          </Button>
        </div>
      ),
    },
  ]

  return (
    <div className="space-y-8 max-w-2xl">
      <div>
        <h2 className="text-xl font-semibold">Usuários</h2>
        <p className="text-sm text-slate-500">Gerencie os usuários do sistema.</p>
      </div>

      {/* Criar usuário */}
      <form
        onSubmit={(e) => { e.preventDefault(); createMutation.mutate() }}
        className="bg-white rounded-xl border p-6 space-y-4"
      >
        <h3 className="font-medium flex items-center gap-2"><UserPlus size={16} />Novo Usuário</h3>
        <div className="grid grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Nome de Usuário *</Label>
            <Input value={username} onChange={e => setUsername(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Senha *</Label>
            <Input type="password" value={password} onChange={e => setPassword(e.target.value)} required />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Função *</Label>
            <Select value={role} onValueChange={setRole} required>
              <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
              <SelectContent>{ROLES.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}</SelectContent>
            </Select>
          </div>
        </div>
        <div className="flex justify-end">
          <Button type="submit" disabled={createMutation.isPending}>
            {createMutation.isPending ? 'Criando...' : 'Criar Usuário'}
          </Button>
        </div>
      </form>

      {/* Alterar senha */}
      {changingPasswordId && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 space-y-3">
          <h3 className="font-medium text-amber-800">Alterar Senha — Usuário #{changingPasswordId}</h3>
          <div className="flex gap-3">
            <Input
              type="password"
              value={newPassword}
              onChange={e => setNewPassword(e.target.value)}
              placeholder="Nova senha"
              className="flex-1"
            />
            <Button
              onClick={() => passwordMutation.mutate({ id: changingPasswordId, pass: newPassword })}
              disabled={!newPassword || passwordMutation.isPending}
            >
              Salvar
            </Button>
            <Button variant="ghost" onClick={() => { setChangingPasswordId(null); setNewPassword('') }}>Cancelar</Button>
          </div>
        </div>
      )}

      <DataTable data={users} columns={columns} searchPlaceholder="Buscar usuário..." />
    </div>
  )
}
