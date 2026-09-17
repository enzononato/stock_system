import { useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import { listUsers, createUser, removeUser, updatePassword, type User } from '@/api/users'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
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
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { Trash2, Key, ShieldCheck, Wrench, GraduationCap, X, Check } from 'lucide-react'

const ROLES = ['Gestor', 'Técnico', 'Jovem Aprendiz'] as const

const ROLE_META: Record<
  string,
  {
    label: string
    sublabel: string
    description: string
    icon: typeof ShieldCheck
  }
> = {
  Gestor: {
    label: 'Gestor',
    sublabel: 'NÍVEL ADMINISTRADOR',
    description: 'ACESSO TOTAL: CADASTROS, RELATÓRIOS, ESTORNOS E CONFIGURAÇÕES DO SISTEMA.',
    icon: ShieldCheck,
  },
  Técnico: {
    label: 'Técnico',
    sublabel: 'NÍVEL OPERADOR',
    description: 'OPERAÇÕES DE ESTOQUE: EMPRÉSTIMOS, DEVOLUÇÕES, CADASTRO DE PERIFÉRICOS.',
    icon: Wrench,
  },
  'Jovem Aprendiz': {
    label: 'Jovem Aprendiz',
    sublabel: 'NÍVEL ASSISTENTE',
    description: 'CONSULTA E OPERAÇÕES LIMITADAS SOB SUPERVISÃO DE GESTOR OU TÉCNICO.',
    icon: GraduationCap,
  },
}

function errorDetail(err: unknown, fallback: string): string {
  return (
    (err as { response?: { data?: { detail?: string } } })?.response?.data?.detail ?? fallback
  )
}

export default function UsersPage() {
  const queryClient = useQueryClient()
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [role, setRole] = useState('')
  const [changingPasswordId, setChangingPasswordId] = useState<number | null>(null)
  const [changingUsername, setChangingUsername] = useState<string>('')
  const [newPassword, setNewPassword] = useState('')

  const { data: users = [] } = useQuery({ queryKey: ['users'], queryFn: listUsers })

  const createMutation = useMutation({
    mutationFn: () => createUser({ username, password, role }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      setUsername('')
      setPassword('')
      setRole('')
      toast.success('Usuário criado com sucesso!')
    },
    onError: (err: unknown) => toast.error(errorDetail(err, 'Erro ao criar usuário.')),
  })

  const removeMutation = useMutation({
    mutationFn: (id: number) => removeUser(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      toast.success('Usuário removido.')
    },
    onError: (err: unknown) => toast.error(errorDetail(err, 'Erro ao remover usuário.')),
  })

  const passwordMutation = useMutation({
    mutationFn: ({ id, pass }: { id: number; pass: string }) => updatePassword(id, pass),
    onSuccess: () => {
      setChangingPasswordId(null)
      setChangingUsername('')
      setNewPassword('')
      toast.success('Senha alterada com sucesso!')
    },
    onError: (err: unknown) => toast.error(errorDetail(err, 'Erro ao alterar senha.')),
  })

  const byRole = ROLES.reduce<Record<string, User[]>>(
    (acc, r) => {
      acc[r] = users.filter((u) => u.role === r)
      return acc
    },
    {} as Record<string, User[]>
  )

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Administração"
        eyebrowDetail="Controle de Acesso"
        title="Usuários"
        description="Gerencie os usuários do sistema e seus níveis de acesso."
      />

      {/* Criar usuário */}
      <form
        onSubmit={(e) => {
          e.preventDefault()
          createMutation.mutate()
        }}
        className="surface-panel p-6 space-y-4"
      >
        <PanelHeader
          title="Cadastro de Usuário"
          description="Defina usuário, senha e função de acesso."
        />
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label>Nome de Usuário *</Label>
            <Input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              required
              placeholder="ex: joao.silva"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Senha *</Label>
            <Input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              placeholder="••••••••"
            />
          </div>
          <div className="flex flex-col gap-1.5">
            <Label>Função *</Label>
            <Select value={role} onValueChange={setRole} required>
              <SelectTrigger>
                <SelectValue placeholder="Selecione" />
              </SelectTrigger>
              <SelectContent>
                {ROLES.map((r) => (
                  <SelectItem key={r} value={r}>
                    {r}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>
        <div className="flex justify-end">
          <Button type="submit" disabled={createMutation.isPending || !username || !password || !role}>
            {createMutation.isPending ? 'Criando...' : 'Criar Usuário'}
          </Button>
        </div>
      </form>

      {/* Alterar senha — inline banner */}
      {changingPasswordId && (
        <div className="surface-panel p-5 border-l-2 border-primary space-y-3">
          <div className="flex items-center justify-between">
            <p className="text-body-sm font-semibold text-foreground">
              Alterar senha — <span className="font-mono text-muted-foreground">{changingUsername}</span>{' '}
              <span className="text-muted-foreground font-normal">#{changingPasswordId}</span>
            </p>
            <Button
              variant="ghost"
              size="icon"
              className="size-7 text-muted-foreground"
              onClick={() => {
                setChangingPasswordId(null)
                setChangingUsername('')
                setNewPassword('')
              }}
            >
              <X size={14} />
            </Button>
          </div>
          <div className="flex gap-2">
            <Input
              type="password"
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              placeholder="Nova senha"
              className="flex-1 h-9"
              autoFocus
            />
            <Button
              size="sm"
              onClick={() =>
                passwordMutation.mutate({ id: changingPasswordId, pass: newPassword })
              }
              disabled={!newPassword || passwordMutation.isPending}
              className="h-9"
            >
              <Check size={14} className="mr-1.5" />
              {passwordMutation.isPending ? 'Salvando...' : 'Salvar'}
            </Button>
          </div>
        </div>
      )}

      {/* Cards por função */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        {ROLES.map((r) => {
          const meta = ROLE_META[r]
          const Icon = meta.icon
          const roleUsers = byRole[r] ?? []

          return (
            <div
              key={r}
              className="surface-panel p-5 flex flex-col gap-4"
            >
              {/* Cabeçalho do card de função */}
              <div className="space-y-1">
                <div className="flex items-center gap-2">
                  <Icon className="size-4 text-muted-foreground shrink-0" aria-hidden />
                  <span className="font-semibold text-foreground text-body-sm">{meta.label}</span>
                </div>
                <p className="text-[10px] font-bold uppercase tracking-widest text-muted-foreground">
                  {meta.sublabel}
                </p>
                <p className="text-[11px] text-muted-foreground leading-snug">
                  {meta.description}
                </p>
              </div>

              {/* Lista de usuários */}
              <div className="flex flex-col gap-1">
                {roleUsers.length === 0 ? (
                  <p className="text-caption text-muted-foreground italic py-2">
                    Nenhum usuário neste nível.
                  </p>
                ) : (
                  roleUsers.map((user) => (
                    <div
                      key={user.id}
                      className="flex items-center justify-between px-3 py-2 rounded border border-border bg-surface-alt hover:bg-surface transition-colors group"
                    >
                      <div>
                        <p className="text-body-sm font-medium text-foreground">{user.username}</p>
                        <p className="text-[11px] font-mono text-muted-foreground">#{user.id}</p>
                      </div>

                      <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                        {/* Alterar senha */}
                        <Button
                          size="icon"
                          variant="ghost"
                          className="size-7 text-muted-foreground hover:text-foreground"
                          title="Alterar senha"
                          onClick={() => {
                            setChangingPasswordId(user.id)
                            setChangingUsername(user.username)
                            setNewPassword('')
                          }}
                        >
                          <Key size={13} />
                        </Button>

                        {/* Remover usuário */}
                        <AlertDialog>
                          <AlertDialogTrigger asChild>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="size-7 text-muted-foreground hover:text-destructive"
                              title="Remover usuário"
                              disabled={removeMutation.isPending}
                            >
                              <Trash2 size={13} />
                            </Button>
                          </AlertDialogTrigger>
                          <AlertDialogContent>
                            <AlertDialogHeader>
                              <AlertDialogTitle>
                                Remover usuário "{user.username}"?
                              </AlertDialogTitle>
                              <AlertDialogDescription>
                                Esta ação não pode ser desfeita. Não é possível remover seu
                                próprio usuário nem o último Gestor do sistema.
                              </AlertDialogDescription>
                            </AlertDialogHeader>
                            <AlertDialogFooter>
                              <AlertDialogCancel>Cancelar</AlertDialogCancel>
                              <AlertDialogAction
                                onClick={() => removeMutation.mutate(user.id)}
                              >
                                Remover
                              </AlertDialogAction>
                            </AlertDialogFooter>
                          </AlertDialogContent>
                        </AlertDialog>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Rodapé do card */}
              <div className="border-t border-border pt-3 mt-auto">
                <p className="text-[10px] font-bold uppercase tracking-widest text-muted-foreground">
                  {roleUsers.length}{' '}
                  {roleUsers.length === 1 ? 'OPERADOR' : 'OPERADORES'}
                </p>
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}
