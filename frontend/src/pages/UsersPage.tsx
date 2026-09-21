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
import { getErrorMessage } from '@/lib/api-error'
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
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import type { ColumnDef } from '@tanstack/react-table'
import { Trash2, Key } from 'lucide-react'

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
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao criar usuário.'), 'error'),
  })

  // O backend agora recusa com 400 a auto-remoção e a remoção do último
  // Gestor do sistema (ver backend/app/routers/users.py) — antes esta
  // mutation descartava a mensagem de erro e sempre mostrava um toast
  // genérico, então o usuário nunca sabia por que a remoção falhou.
  const removeMutation = useMutation({
    mutationFn: (id: number) => removeUser(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['users'] })
      toast('Usuário removido.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao remover usuário.'), 'error'),
  })

  const passwordMutation = useMutation({
    mutationFn: ({ id, pass }: { id: number; pass: string }) => updatePassword(id, pass),
    onSuccess: () => {
      setChangingPasswordId(null); setNewPassword('')
      toast('Senha alterada com sucesso!')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao alterar senha.'), 'error'),
  })

  const columns: ColumnDef<User, unknown>[] = [
    { accessorKey: 'id', header: 'ID', size: 60, cell: ({ getValue }) => <span className="num">{getValue() as number}</span> },
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
          {/* T4: confirmação antes de remover — não havia nenhuma, o clique
              disparava a exclusão direto. */}
          <AlertDialog>
            <AlertDialogTrigger asChild>
              <Button
                size="sm"
                variant="ghost"
                className="text-muted-foreground hover:text-destructive"
                disabled={removeMutation.isPending}
              >
                <Trash2 size={14} />
              </Button>
            </AlertDialogTrigger>
            <AlertDialogContent>
              <AlertDialogHeader>
                <AlertDialogTitle>Remover usuário "{row.original.username}"?</AlertDialogTitle>
                <AlertDialogDescription>
                  Esta ação não pode ser desfeita. Não é possível remover seu próprio usuário
                  nem o último Gestor do sistema.
                </AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>Cancelar</AlertDialogCancel>
                <AlertDialogAction onClick={() => removeMutation.mutate(row.original.id)}>
                  Remover
                </AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>
        </div>
      ),
    },
  ]

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Administração"
        eyebrowDetail="Controle de Acesso"
        title="Usuários"
        description="Gerencie os usuários do sistema."
      />

      {/* Criar usuário */}
      <form
        onSubmit={(e) => { e.preventDefault(); createMutation.mutate() }}
        className="surface-panel p-6 space-y-4"
      >
        <PanelHeader title="Cadastro de Usuário" description="Defina usuário, senha e função de acesso." />
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
        <div className="figure-ground-panel space-y-3">
          <PanelHeader title="Segurança da Conta" />
          <h3 className="text-heading-sm text-foreground">Alterar Senha — Usuário #<span className="num">{changingPasswordId}</span></h3>
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
