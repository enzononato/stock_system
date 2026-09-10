import { useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Key, Lock, Plus, Shield, Trash2, UserPlus, Users, Wrench } from "lucide-react";
import { listUsers, createUser, removeUser, updatePassword, type User } from "@/api/users";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
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
} from "@/components/ui/alert-dialog";
import { getErrorMessage } from "@/lib/api-error";
import { toast } from "sonner";

const ROLES = ["Gestor", "Técnico", "Jovem Aprendiz"];

const ROLE_META: Record<string, { label: string; iconSrc: string; description: string }> = {
  Gestor: {
    label: "Gestor",
    iconSrc: "/icons/role-gestor.png",
    description: "Acesso total: cadastros, relatórios, estornos e configurações do sistema.",
  },
  Técnico: {
    label: "Técnico",
    iconSrc: "/icons/role-tecnico.png",
    description: "Operações de estoque: empréstimos, devoluções, cadastro de periféricos.",
  },
  "Jovem Aprendiz": {
    label: "Jovem Aprendiz",
    iconSrc: "/icons/role-aprendiz.png",
    description: "Consulta e operações limitadas sob supervisão de Gestor ou Técnico.",
  },
};

function RoleBadge({ role }: { role: string }) {
  const meta = ROLE_META[role] ?? { label: role, iconSrc: "/icons/role-gestor.png" };
  return (
    <span className="inline-flex items-center gap-1.5 font-medium text-caption text-foreground">
      {meta.iconSrc && (
        <span className="size-4 rounded-xs border border-border/80 bg-surface-alt flex items-center justify-center p-0.5 shrink-0">
          <img src={meta.iconSrc} alt={meta.label} className="size-full object-contain dark:invert" />
        </span>
      )}
      <span>{meta.label}</span>
    </span>
  );
}

export function UsersPage() {
  const queryClient = useQueryClient();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [role, setRole] = useState("");
  const [showCreateForm, setShowCreateForm] = useState(false);
  const [changingPasswordId, setChangingPasswordId] = useState<number | null>(null);
  const [newPassword, setNewPassword] = useState("");
  const [usersPage, setUsersPage] = useState(0);
  const USERS_PAGE_SIZE = 7;

  const { data: users = [], isLoading, error, refetch } = useQuery({
    queryKey: ["users"],
    queryFn: listUsers,
  });

  const createMutation = useMutation({
    mutationFn: () => createUser({ username, password, role }),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["users"] });
      setUsername(""); setPassword(""); setRole("");
      setShowCreateForm(false);
      toast.success("Usuário criado com sucesso!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao criar usuário.")),
  });

  const removeMutation = useMutation({
    mutationFn: (id: number) => removeUser(id),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["users"] });
      toast.success("Usuário removido.");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao remover usuário.")),
  });

  const passwordMutation = useMutation({
    mutationFn: ({ id, pass }: { id: number; pass: string }) => updatePassword(id, pass),
    onSuccess: () => {
      setChangingPasswordId(null);
      setNewPassword("");
      toast.success("Senha alterada com sucesso!");
    },
    onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao alterar senha.")),
  });

  const changingUser = users.find((u) => u.id === changingPasswordId);

  const gestores = users.filter((u) => u.role === "Gestor");
  const tecnicos = users.filter((u) => u.role === "Técnico");
  const aprendizes = users.filter((u) => u.role === "Jovem Aprendiz");

  return (
    <div className="page-container-dense space-y-6">
      {/* Header */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between border-b border-border pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
              Administração de Acessos
            </span>
            <span className="text-muted-foreground">•</span>
            <span className="text-caption text-foreground font-medium">Matriz de Governança</span>
          </div>
          <h1 className="text-heading font-semibold tracking-tight text-foreground mt-0.5">
            Gestão de Usuários
          </h1>
          <p className="text-body-sm text-muted-foreground mt-1">
            Contas de operadores, perfis de acesso e credenciais de segurança corporativa.
          </p>
        </div>
        <Button
          size="sm"
          onClick={() => setShowCreateForm((p) => !p)}
          className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8"
        >
          <UserPlus className="mr-1.5 size-3.5" />
          Novo Usuário
        </Button>
      </div>

      {/* Formulário de Criação (colapsável) */}
      {showCreateForm && (
        <div className="rounded-[6px] border border-border bg-surface p-5 space-y-4">
          <div className="border-b border-border pb-2.5">
            <h2 className="text-body font-semibold text-foreground">Criar Nova Conta de Operador</h2>
            <p className="text-caption text-muted-foreground">
              Defina o nome de usuário, senha inicial e o perfil de permissões conforme o princípio de menor privilégio.
            </p>
          </div>
          <form
            onSubmit={(e) => { e.preventDefault(); createMutation.mutate(); }}
            className="space-y-4"
          >
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Nome de usuário *</Label>
                <Input
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Ex.: joao.silva"
                  required
                  className="rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Senha inicial *</Label>
                <Input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Mínimo 6 caracteres"
                  required
                  className="rounded-[4px] border-border bg-background text-body-sm h-9"
                />
              </div>
              <div className="space-y-1.5">
                <Label className="text-caption font-medium">Perfil de acesso *</Label>
                <Select value={role} onValueChange={setRole}>
                  <SelectTrigger className="h-9 rounded-[4px] border-border bg-background text-body-sm">
                    <SelectValue placeholder="Selecione o perfil" />
                  </SelectTrigger>
                  <SelectContent className="rounded-[4px] border-border bg-surface">
                    {ROLES.map((r) => (
                      <SelectItem key={r} value={r} className="text-body-sm rounded-[2px]">
                        {ROLE_META[r]?.symbol} {r}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            {/* Preview do perfil selecionado */}
            {role && ROLE_META[role] && (
              <div className="rounded-[4px] border border-border bg-surface-alt px-3 py-2 text-caption text-muted-foreground">
                <span className="font-semibold text-foreground font-mono mr-2">{ROLE_META[role]?.symbol}</span>
                {ROLE_META[role]?.description}
              </div>
            )}

            <div className="flex justify-end gap-2 pt-1">
              <Button type="button" variant="outline" size="sm" onClick={() => setShowCreateForm(false)} className="rounded-[4px] border-border text-xs h-8">
                Cancelar
              </Button>
              <Button
                type="submit"
                size="sm"
                disabled={createMutation.isPending || !username || !password || !role}
                className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs h-8"
              >
                <UserPlus className="mr-1.5 size-3.5" />
                {createMutation.isPending ? "Criando conta…" : "Criar usuário"}
              </Button>
            </div>
          </form>
        </div>
      )}

      {/* Aviso de governança */}
      <div className="rounded-[4px] border border-border bg-surface-alt px-4 py-3 text-caption text-muted-foreground">
        <span className="font-semibold text-foreground">Princípio de menor privilégio —</span> Atribua a cada operador somente as permissões estritamente necessárias para sua função corporativa. Gestores possuem acesso irreversível ao estorno de operações.
      </div>

      {/* Matriz de Governança por Perfil */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        {[
          { role: "Gestor", users: gestores, iconSrc: "/icons/role-gestor.png", level: "Nível Administrador" },
          { role: "Técnico", users: tecnicos, iconSrc: "/icons/role-tecnico.png", level: "Nível Operador" },
          { role: "Jovem Aprendiz", users: aprendizes, iconSrc: "/icons/role-aprendiz.png", level: "Nível Assistente" },
        ].map(({ role: r, users: roleUsers, iconSrc, level }) => (
          <div key={r} className="rounded-[6px] border border-border bg-surface p-4 space-y-3">
            <div className="border-b border-border pb-2.5">
              <div className="flex items-center gap-2.5">
                <div className="size-8 rounded-[4px] border border-border bg-surface-alt flex items-center justify-center p-1.5 shrink-0 shadow-2xs">
                  <img
                    src={iconSrc}
                    alt={r}
                    className="size-full object-contain dark:invert"
                  />
                </div>
                <div className="min-w-0">
                  <span className="text-body-sm font-semibold text-foreground block leading-tight">{r}</span>
                  <span className="text-[10px] text-muted-foreground uppercase font-mono tracking-wider">
                    {level}
                  </span>
                </div>
              </div>
              <p className="text-caption text-muted-foreground mt-2">
                {ROLE_META[r]?.description}
              </p>
            </div>

            <div className="space-y-1">
              {isLoading ? (
                <p className="text-caption text-muted-foreground py-2">Carregando…</p>
              ) : roleUsers.length === 0 ? (
                <p className="text-caption text-muted-foreground italic py-2">Nenhum {r.toLowerCase()} cadastrado.</p>
              ) : (
                roleUsers.map((u) => (
                  <div
                    key={u.id}
                    className="flex items-center justify-between gap-2 rounded-[4px] px-3 py-2 bg-surface-alt border border-border/50"
                  >
                    <div className="min-w-0">
                      <span className="text-body-sm font-medium text-foreground block truncate">{u.username}</span>
                      <span className="font-mono text-caption text-muted-foreground">#{u.id}</span>
                    </div>
                    <div className="flex items-center gap-1 shrink-0">
                      <Button
                        size="sm"
                        variant="ghost"
                        onClick={() => setChangingPasswordId(u.id)}
                        className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
                        title="Alterar senha"
                      >
                        <Key className="size-3" />
                      </Button>
                      <AlertDialog>
                        <AlertDialogTrigger asChild>
                          <Button
                            size="sm"
                            variant="ghost"
                            className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
                            title="Remover usuário"
                          >
                            <Trash2 className="size-3" />
                          </Button>
                        </AlertDialogTrigger>
                        <AlertDialogContent className="rounded-[6px] border-border bg-surface">
                          <AlertDialogHeader>
                            <AlertDialogTitle className="text-body-lg font-semibold">
                              Remover "{u.username}"?
                            </AlertDialogTitle>
                            <AlertDialogDescription className="text-caption text-muted-foreground">
                              Esta ação não pode ser desfeita. Não é possível remover seu próprio usuário nem o último Gestor ativo no sistema.
                            </AlertDialogDescription>
                          </AlertDialogHeader>
                          <AlertDialogFooter>
                            <AlertDialogCancel className="rounded-[4px] border-border text-xs">Cancelar</AlertDialogCancel>
                            <AlertDialogAction
                              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
                              onClick={() => removeMutation.mutate(u.id)}
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

            <div className="pt-1 border-t border-border text-caption text-muted-foreground font-mono">
              {roleUsers.length} operador{roleUsers.length !== 1 ? "es" : ""}
            </div>
          </div>
        ))}
      </div>

      {/* Tabela Flat Completa de todos os usuários */}
      <div className="rounded-[6px] border border-border bg-surface p-5 space-y-3">
        <div className="border-b border-border pb-2.5">
          <span className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">
            Todos os Operadores — {users.length} conta{users.length !== 1 ? "s" : ""}
          </span>
        </div>
        <table className="w-full text-left text-body-sm">
          <thead className="border-b border-border text-caption uppercase tracking-wider text-muted-foreground">
            <tr>
              <th className="py-2 px-3">ID</th>
              <th className="py-2 px-3">Usuário</th>
              <th className="py-2 px-3">Perfil</th>
              <th className="py-2 px-3 text-right">Ações</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-border">
            {users.length === 0 ? (
              <tr>
                <td colSpan={4} className="py-8 text-center text-caption text-muted-foreground">
                  {isLoading ? "Carregando usuários…" : "Nenhum usuário cadastrado."}
                </td>
              </tr>
            ) : (
              users
                .slice(usersPage * USERS_PAGE_SIZE, (usersPage + 1) * USERS_PAGE_SIZE)
                .map((u) => (
                <tr key={u.id} className="hover:bg-muted/30">
                  <td className="py-2 px-3 font-mono text-caption text-muted-foreground">#{u.id}</td>
                  <td className="py-2 px-3 font-medium text-foreground">{u.username}</td>
                  <td className="py-2 px-3">
                    <RoleBadge role={u.role} />
                  </td>
                  <td className="py-2 px-3 text-right">
                    <div className="flex items-center gap-1 justify-end">
                      <Button
                        size="sm"
                        variant="ghost"
                        onClick={() => setChangingPasswordId(u.id)}
                        className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
                        title="Alterar senha"
                      >
                        <Key className="size-3.5" />
                      </Button>
                      <AlertDialog>
                        <AlertDialogTrigger asChild>
                          <Button
                            size="sm"
                            variant="ghost"
                            className="size-7 p-0 rounded-[2px] text-muted-foreground hover:text-foreground"
                            title="Remover"
                          >
                            <Trash2 className="size-3.5" />
                          </Button>
                        </AlertDialogTrigger>
                        <AlertDialogContent className="rounded-[6px] border-border bg-surface">
                          <AlertDialogHeader>
                            <AlertDialogTitle className="text-body-lg font-semibold">Remover "{u.username}"?</AlertDialogTitle>
                            <AlertDialogDescription className="text-caption text-muted-foreground">
                              Esta ação não pode ser desfeita. O último Gestor não pode ser removido.
                            </AlertDialogDescription>
                          </AlertDialogHeader>
                          <AlertDialogFooter>
                            <AlertDialogCancel className="rounded-[4px] border-border text-xs">Cancelar</AlertDialogCancel>
                            <AlertDialogAction
                              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
                              onClick={() => removeMutation.mutate(u.id)}
                            >
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

        {Math.ceil(users.length / USERS_PAGE_SIZE) > 1 && (
          <div className="flex items-center justify-between border-t border-border pt-3">
            <span className="text-caption text-muted-foreground">
              Página {usersPage + 1} de {Math.ceil(users.length / USERS_PAGE_SIZE)} ({users.length} operadores)
            </span>
            <div className="flex items-center gap-2">
              <Button
                size="sm"
                variant="outline"
                disabled={usersPage === 0}
                onClick={() => setUsersPage((p) => Math.max(0, p - 1))}
                className="rounded-[4px] border-border text-xs h-7 px-2.5"
              >
                ← Anterior
              </Button>
              <Button
                size="sm"
                variant="outline"
                disabled={usersPage >= Math.ceil(users.length / USERS_PAGE_SIZE) - 1}
                onClick={() => setUsersPage((p) => p + 1)}
                className="rounded-[4px] border-border text-xs h-7 px-2.5"
              >
                Próxima →
              </Button>
            </div>
          </div>
        )}
      </div>

      {/* Dialog: Alterar Senha */}
      <Dialog
        open={changingPasswordId !== null}
        onOpenChange={(open) => { if (!open) { setChangingPasswordId(null); setNewPassword(""); } }}
      >
        <DialogContent className="max-w-sm rounded-[6px] border border-border bg-surface p-6">
          <DialogHeader>
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 rounded-[4px] border border-border bg-surface-alt">
                <Lock className="size-4 text-foreground" />
              </div>
              <div>
                <DialogTitle className="text-body-lg font-semibold text-foreground">
                  Alterar Senha
                </DialogTitle>
                <DialogDescription className="text-caption text-muted-foreground">
                  {changingUser?.username ?? `Operador #${changingPasswordId}`}
                </DialogDescription>
              </div>
            </div>
          </DialogHeader>
          <div className="space-y-2">
            <Label className="text-caption font-medium">Nova senha corporativa *</Label>
            <Input
              type="password"
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              placeholder="Mínimo 6 caracteres"
              autoFocus
              className="rounded-[4px] border-border bg-background text-body-sm h-9"
            />
          </div>
          <DialogFooter className="pt-4 border-t border-border flex justify-end gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => { setChangingPasswordId(null); setNewPassword(""); }}
              className="rounded-[4px] border-border text-xs"
            >
              Cancelar
            </Button>
            <Button
              size="sm"
              onClick={() =>
                changingPasswordId !== null &&
                passwordMutation.mutate({ id: changingPasswordId, pass: newPassword })
              }
              disabled={!newPassword || passwordMutation.isPending}
              className="rounded-[4px] bg-foreground text-background hover:bg-foreground/90 text-xs"
            >
              {passwordMutation.isPending ? "Salvando…" : "Salvar nova senha"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
