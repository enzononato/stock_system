import { useMemo, useState } from 'react'
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query'
import {
  listUsers,
  createUser,
  removeUser,
  enviarLinkSenha,
  updateUserEmail,
  zerar2fa,
  type User,
} from '@/api/users'
import { useAuth } from '@/contexts/AuthContext'
import { usePageSize } from '@/hooks/usePageSize'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Pagination } from '@/components/ui/Pagination'
import { StatusPill } from '@/components/ui/StatusPill'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  Dialog,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
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
import { Trash2, Mail, ShieldOff, Search, Send, KeyRound } from 'lucide-react'

const ROLES = ['Gestor', 'Técnico', 'Jovem Aprendiz'] as const
const TODAS_AS_FUNCOES = 'todas'

function normalizar(texto: string): string {
  return texto
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
}

function PilulaDe2fa({ user }: { user: User }) {
  if (user.totp_ativo) return <StatusPill status="Ativo" />
  if (user.role === 'Gestor') return <StatusPill status="Pendente" />
  return <StatusPill status="Inativo" tone="neutral" />
}

/** Situação da conta: convite ainda não aceito, e-mail a confirmar no próximo login, ou pronta. */
function PilulaDaConta({ user }: { user: User }) {
  if (!user.senha_definida) return <StatusPill status="Convite pendente" tone="pending" />
  if (!user.email_confirmado) return <StatusPill status="Aguardando confirmação" tone="return" />
  return <StatusPill status="Confirmado" tone="available" />
}

export default function UsersPage() {
  const queryClient = useQueryClient()
  const { user: logado } = useAuth()
  const [role, setRole] = useState('')
  const [email, setEmail] = useState('')

  const [busca, setBusca] = useState('')
  const [filtroFuncao, setFiltroFuncao] = useState(TODAS_AS_FUNCOES)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = usePageSize('usuarios')

  const [editando, setEditando] = useState<User | null>(null)
  const [emailEdicao, setEmailEdicao] = useState('')
  const [erroEdicao, setErroEdicao] = useState('')

  const { data: users = [] } = useQuery({ queryKey: ['users'], queryFn: listUsers })

  const createMutation = useMutation({
    mutationFn: () => createUser({ email: email.trim(), role }),
    onSuccess: (data) => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      setRole('')
      setEmail('')
      // Sem o envio do convite a conta existe, mas a pessoa não tem como entrar: aviso de erro.
      toast(data.detail, data.convite_enviado ? undefined : 'error')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao criar usuário.'), 'error'),
  })

  // O backend recusa com 400 a auto-remoção e a remoção do último Gestor do
  // sistema (ver backend/app/routers/users.py) — por isso a mensagem de erro
  // usa getErrorMessage em vez de um toast genérico fixo, para o usuário
  // entender por que a remoção falhou.
  const removeMutation = useMutation({
    mutationFn: (id: number) => removeUser(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      toast('Usuário removido.')
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao remover usuário.'), 'error'),
  })

  // Ninguém define a senha de outra pessoa: o Gestor só manda o link (convite ou nova senha).
  const linkSenhaMutation = useMutation({
    mutationFn: (id: number) => enviarLinkSenha(id),
    onSuccess: (data) => toast(data.detail),
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao enviar o link.'), 'error'),
  })

  const emailMutation = useMutation({
    mutationFn: ({ id, valor }: { id: number; valor: string | null }) => updateUserEmail(id, valor),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      fecharEdicao()
      toast('E-mail atualizado.')
    },
    onError: (err: unknown) => setErroEdicao(getErrorMessage(err, 'Erro ao atualizar o e-mail.')),
  })

  const zerarMutation = useMutation({
    mutationFn: (id: number) => zerar2fa(id),
    onSuccess: (data) => {
      void queryClient.invalidateQueries({ queryKey: ['users'] })
      toast(data.detail)
    },
    onError: (err: unknown) => toast(getErrorMessage(err, 'Erro ao zerar a verificação em duas etapas.'), 'error'),
  })

  function abrirEdicao(u: User) {
    setEditando(u)
    setEmailEdicao(u.email ?? '')
    setErroEdicao('')
  }

  function fecharEdicao() {
    setEditando(null)
    setEmailEdicao('')
    setErroEdicao('')
  }

  // A lista é pequena e vem inteira de /api/users (sem paginação no servidor):
  // filtramos e paginamos aqui, no cliente.
  const filtrados = useMemo(() => {
    const termo = normalizar(busca.trim())
    return users.filter((u) => {
      if (filtroFuncao !== TODAS_AS_FUNCOES && u.role !== filtroFuncao) return false
      if (!termo) return true
      return normalizar(u.username).includes(termo) || normalizar(u.email ?? '').includes(termo)
    })
  }, [users, busca, filtroFuncao])

  const ultimaPagina = Math.max(1, Math.ceil(filtrados.length / pageSize))
  const paginaAtual = Math.min(page, ultimaPagina) // corrige quando a lista encolhe (remoção, filtro)
  const visiveis = filtrados.slice((paginaAtual - 1) * pageSize, paginaAtual * pageSize)

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Administração"
        eyebrowDetail="Controle de Acesso"
        title="Usuários"
        description="Gerencie os usuários do sistema, seus e-mails e níveis de acesso."
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
          description="Informe o e-mail corporativo e o papel. A pessoa recebe um convite por e-mail e cria a própria senha."
        />
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="novo-email">E-mail *</Label>
            <Input
              id="novo-email"
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              required
              placeholder="nome@revalle.com.br"
              autoComplete="off"
              aria-describedby="novo-email-dica"
            />
            <p id="novo-email-dica" className="text-caption text-muted-foreground">
              Só e-mails @revalle.com.br.
            </p>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="novo-role">Papel *</Label>
            <Select value={role} onValueChange={setRole} required>
              <SelectTrigger id="novo-role">
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
          <Button
            type="submit"
            loading={createMutation.isPending}
            disabled={!email.trim() || !role}
          >
            Criar e enviar convite
          </Button>
        </div>
      </form>

      {/* Lista de usuários */}
      <section className="surface-panel p-6 space-y-4">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="relative w-full sm:max-w-sm">
            <Search
              size={16}
              className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground"
              aria-hidden
            />
            <Input
              value={busca}
              onChange={(e) => {
                setBusca(e.target.value)
                setPage(1)
              }}
              placeholder="Buscar por usuário ou e-mail"
              aria-label="Buscar por usuário ou e-mail"
              className="pl-9"
            />
          </div>
          <div className="w-full sm:w-52">
            <Select
              value={filtroFuncao}
              onValueChange={(v) => {
                setFiltroFuncao(v)
                setPage(1)
              }}
            >
              <SelectTrigger aria-label="Filtrar por função">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value={TODAS_AS_FUNCOES}>Todas as funções</SelectItem>
                {ROLES.map((r) => (
                  <SelectItem key={r} value={r}>
                    {r}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-body-sm">
            <thead>
              <tr className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
                <th className="py-2 pr-4 font-semibold">E-mail</th>
                <th className="py-2 pr-4 font-semibold">Situação</th>
                <th className="py-2 pr-4 font-semibold">Função</th>
                <th className="py-2 pr-4 font-semibold">2FA</th>
                <th className="py-2 pr-4 font-semibold">Usuário (histórico)</th>
                <th className="py-2 font-semibold text-right">Ações</th>
              </tr>
            </thead>
            <tbody>
              {visiveis.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-6 text-center text-muted-foreground italic">
                    Nenhum usuário encontrado.
                  </td>
                </tr>
              ) : (
                visiveis.map((u) => (
                  <tr key={u.id} className="border-b border-border last:border-0">
                    <td className="py-2.5 pr-4">
                      <span className="block font-medium text-foreground">{u.email ?? '—'}</span>
                      <span className="block text-[11px] font-mono text-muted-foreground">#{u.id}</span>
                    </td>
                    <td className="py-2.5 pr-4">
                      <PilulaDaConta user={u} />
                    </td>
                    <td className="py-2.5 pr-4 text-foreground">{u.role}</td>
                    <td className="py-2.5 pr-4">
                      <PilulaDe2fa user={u} />
                    </td>
                    <td className="py-2.5 pr-4 font-mono text-[12px] text-muted-foreground">{u.username}</td>
                    <td className="py-2.5">
                      <div className="flex items-center justify-end gap-1">
                        <Button
                          size="icon"
                          variant="ghost"
                          className="size-7 text-muted-foreground hover:text-foreground"
                          title="Editar e-mail"
                          aria-label="Editar e-mail"
                          onClick={() => abrirEdicao(u)}
                        >
                          <Mail size={13} />
                        </Button>

                        {u.senha_definida ? (
                          <Button
                            size="icon"
                            variant="ghost"
                            className="size-7 text-muted-foreground hover:text-foreground"
                            title="Enviar link para criar nova senha"
                            aria-label="Enviar link para criar nova senha"
                            disabled={!u.email || linkSenhaMutation.isPending}
                            onClick={() => linkSenhaMutation.mutate(u.id)}
                          >
                            <KeyRound size={13} />
                          </Button>
                        ) : (
                          <Button
                            size="icon"
                            variant="ghost"
                            className="size-7 text-muted-foreground hover:text-foreground"
                            title="Reenviar convite"
                            aria-label="Reenviar convite"
                            disabled={!u.email || linkSenhaMutation.isPending}
                            onClick={() => linkSenhaMutation.mutate(u.id)}
                          >
                            <Send size={13} />
                          </Button>
                        )}

                        {u.totp_ativo && u.id !== logado?.id && (
                          <AlertDialog>
                            <AlertDialogTrigger asChild>
                              <Button
                                size="icon"
                                variant="ghost"
                                className="size-7 text-muted-foreground hover:text-foreground"
                                title="Zerar 2FA"
                                aria-label="Zerar 2FA"
                                disabled={zerarMutation.isPending}
                              >
                                <ShieldOff size={13} />
                              </Button>
                            </AlertDialogTrigger>
                            <AlertDialogContent>
                              <AlertDialogHeader>
                                <AlertDialogTitle>
                                  Zerar a verificação em duas etapas de "{u.username}"?
                                </AlertDialogTitle>
                                <AlertDialogDescription>
                                  As sessões abertas do usuário serão encerradas. Um Gestor terá de cadastrar
                                  um novo autenticador no próximo login; os demais voltam a entrar só com a
                                  senha.
                                </AlertDialogDescription>
                              </AlertDialogHeader>
                              <AlertDialogFooter>
                                <AlertDialogCancel>Cancelar</AlertDialogCancel>
                                <AlertDialogAction onClick={() => zerarMutation.mutate(u.id)}>
                                  Zerar 2FA
                                </AlertDialogAction>
                              </AlertDialogFooter>
                            </AlertDialogContent>
                          </AlertDialog>
                        )}

                        <AlertDialog>
                          <AlertDialogTrigger asChild>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="size-7 text-muted-foreground hover:text-destructive"
                              title="Remover usuário"
                              aria-label="Remover usuário"
                              disabled={removeMutation.isPending}
                            >
                              <Trash2 size={13} />
                            </Button>
                          </AlertDialogTrigger>
                          <AlertDialogContent>
                            <AlertDialogHeader>
                              <AlertDialogTitle>Remover usuário "{u.username}"?</AlertDialogTitle>
                              <AlertDialogDescription>
                                Esta ação não pode ser desfeita. Não é possível remover seu próprio usuário
                                nem o último Gestor do sistema.
                              </AlertDialogDescription>
                            </AlertDialogHeader>
                            <AlertDialogFooter>
                              <AlertDialogCancel>Cancelar</AlertDialogCancel>
                              <AlertDialogAction onClick={() => removeMutation.mutate(u.id)}>
                                Remover
                              </AlertDialogAction>
                            </AlertDialogFooter>
                          </AlertDialogContent>
                        </AlertDialog>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        <Pagination
          page={paginaAtual}
          pageSize={pageSize}
          total={filtrados.length}
          onPageChange={setPage}
          onPageSizeChange={(n) => {
            setPageSize(n)
            setPage(1)
          }}
        />
      </section>

      {/* Editar e-mail */}
      <Dialog open={editando !== null} onOpenChange={(aberto) => !aberto && fecharEdicao()}>
        <DialogContent>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              if (editando) emailMutation.mutate({ id: editando.id, valor: emailEdicao.trim() || null })
            }}
            className="space-y-4"
          >
            <DialogHeader>
              <DialogTitle>Editar e-mail — {editando?.username}</DialogTitle>
            </DialogHeader>
            <div className="space-y-1.5">
              <Label htmlFor="email-edicao">E-mail</Label>
              <Input
                id="email-edicao"
                type="email"
                value={emailEdicao}
                onChange={(e) => setEmailEdicao(e.target.value)}
                placeholder="nome@revalle.com.br"
                autoComplete="off"
                autoFocus
              />
              <p className="text-caption text-muted-foreground">
                É o e-mail de login e de recuperação de senha. Um e-mail novo precisa ser confirmado
                pela pessoa no próximo login.
              </p>
            </div>
            {erroEdicao && (
              <p role="alert" className="text-body-sm text-destructive">
                {erroEdicao}
              </p>
            )}
            <DialogFooter>
              <Button type="button" variant="outline" onClick={fecharEdicao}>
                Cancelar
              </Button>
              <Button type="submit" loading={emailMutation.isPending}>
                Salvar
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}
