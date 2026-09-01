import { useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { Key, Trash2, UserPlus, ShieldCheck } from "lucide-react";
import { listUsers, createUser, removeUser, updatePassword, type User } from "@/api/users";
import { DataTable, type Column } from "@/components/app/DataTable";
import { KpiCard } from "@/components/app/KpiCard";
import { PageHeader, Section } from "@/components/app/PageHeader";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { getErrorMessage } from "@/lib/api-error";
import { toast } from "sonner";

const ROLES = ["Gestor", "Técnico", "Jovem Aprendiz"];

export function UsersPage() {
  const queryClient = useQueryClient();
  const [username, setUsername] = useState(""); const [password, setPassword] = useState(""); const [role, setRole] = useState("");
  const [changingPasswordId, setChangingPasswordId] = useState<number | null>(null); const [newPassword, setNewPassword] = useState("");
  const { data: users = [], isLoading, error, refetch } = useQuery({ queryKey: ["users"], queryFn: listUsers });
  const createMutation = useMutation({ mutationFn: () => createUser({ username, password, role }), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["users"] }); setUsername(""); setPassword(""); setRole(""); toast.success("Usuário criado com sucesso!"); }, onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao criar usuário.")) });
  const removeMutation = useMutation({ mutationFn: (id: number) => removeUser(id), onSuccess: () => { queryClient.invalidateQueries({ queryKey: ["users"] }); toast.success("Usuário removido."); }, onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao remover usuário.")) });
  const passwordMutation = useMutation({ mutationFn: ({ id, pass }: { id: number; pass: string }) => updatePassword(id, pass), onSuccess: () => { setChangingPasswordId(null); setNewPassword(""); toast.success("Senha alterada com sucesso!"); }, onError: (err: unknown) => toast.error(getErrorMessage(err, "Erro ao alterar senha.")) });
  const roleBadgeClass = (r: string) => r === "Gestor" ? "bg-blue-500/12 text-blue-700 dark:text-blue-400 border-blue-500/25" : r === "Técnico" ? "bg-amber-500/14 text-amber-700 dark:text-amber-400 border-amber-500/30" : "bg-emerald-500/12 text-emerald-700 dark:text-emerald-400 border-emerald-500/25";
  const columns: Column<User>[] = [
    { key: "id", header: "ID", cell: (u) => u.id }, { key: "username", header: "Usuário", cell: (u) => u.username, primary: true },
    { key: "role", header: "Função", cell: (u) => <Badge variant="outline" className={roleBadgeClass(u.role)}>{u.role}</Badge> },
    { key: "actions", header: "Ações", cell: (u) => <div className="flex gap-1"><Button size="sm" variant="ghost" aria-label={`Alterar senha de ${u.username}`} onClick={() => setChangingPasswordId(u.id)}><Key className="size-3.5" /></Button><AlertDialog><AlertDialogTrigger asChild><Button size="sm" variant="ghost" aria-label={`Remover ${u.username}`} className="text-destructive hover:text-destructive/80"><Trash2 className="size-3.5" /></Button></AlertDialogTrigger><AlertDialogContent><AlertDialogHeader><AlertDialogTitle>Remover usuário &quot;{u.username}&quot;?</AlertDialogTitle><AlertDialogDescription>Esta ação não pode ser desfeita. Não é possível remover seu próprio usuário nem o último Gestor do sistema.</AlertDialogDescription></AlertDialogHeader><AlertDialogFooter><AlertDialogCancel>Cancelar</AlertDialogCancel><AlertDialogAction onClick={() => removeMutation.mutate(u.id)}>Remover</AlertDialogAction></AlertDialogFooter></AlertDialogContent></AlertDialog></div> }
  ];
  const changingUser = users.find((u) => u.id === changingPasswordId);
  return <div className="space-y-7">
    <PageHeader eyebrow="Administração" title="Gestão de Usuários" description="Gerencie usuários, funções e credenciais de acesso em um único lugar." />
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-3"><KpiCard label="Total de usuários" value={users.length} /><KpiCard label="Gestores" value={users.filter((u) => u.role === "Gestor").length} /><KpiCard label="Técnicos" value={users.filter((u) => u.role === "Técnico").length} /></div>
    <Section title="Novo usuário" description="Cadastre uma conta e defina sua função de acesso."><form onSubmit={(e) => { e.preventDefault(); createMutation.mutate(); }} className="space-y-4"><div className="grid grid-cols-1 gap-4 md:grid-cols-3"><div className="flex flex-col gap-1.5"><Label htmlFor="new-username">Nome de usuário *</Label><Input id="new-username" value={username} onChange={(e) => setUsername(e.target.value)} required /></div><div className="flex flex-col gap-1.5"><Label htmlFor="new-password">Senha *</Label><Input id="new-password" type="password" value={password} onChange={(e) => setPassword(e.target.value)} required /></div><div className="flex flex-col gap-1.5"><Label>Função *</Label><Select value={role} onValueChange={setRole}><SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger><SelectContent>{ROLES.map((r) => <SelectItem key={r} value={r}>{r}</SelectItem>)}</SelectContent></Select></div></div><div className="flex justify-end"><Button type="submit" disabled={createMutation.isPending || !username || !password || !role}><UserPlus className="mr-2 size-4" />{createMutation.isPending ? "Criando…" : "Criar usuário"}</Button></div></form></Section>
    <Section title="Usuários cadastrados" description={`${users.length} conta(s) encontrada(s).`}><DataTable data={users} columns={columns} rowKey={(u) => u.id} isLoading={isLoading} error={error} onRetry={() => void refetch()} emptyTitle="Nenhum usuário cadastrado" /></Section>
    <div className="flex items-start gap-3 rounded-lg border border-border bg-muted/30 p-4 text-sm text-muted-foreground"><ShieldCheck className="mt-0.5 size-4 shrink-0" /><p>Princípio de menor privilégio: atribua a cada usuário somente a função necessária para suas atividades.</p></div>
    <Dialog open={changingPasswordId !== null} onOpenChange={(open) => { if (!open) { setChangingPasswordId(null); setNewPassword(""); } }}><DialogContent><DialogHeader><DialogTitle>Alterar senha — {changingUser?.username ?? `Usuário #${changingPasswordId}`}</DialogTitle></DialogHeader><div className="flex flex-col gap-1.5"><Label htmlFor="change-password-input">Nova senha</Label><Input id="change-password-input" type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} placeholder="Nova senha" autoFocus /></div><DialogFooter><Button variant="ghost" onClick={() => { setChangingPasswordId(null); setNewPassword(""); }}>Cancelar</Button><Button onClick={() => changingPasswordId !== null && passwordMutation.mutate({ id: changingPasswordId, pass: newPassword })} disabled={!newPassword || passwordMutation.isPending}>{passwordMutation.isPending ? "Salvando…" : "Salvar"}</Button></DialogFooter></DialogContent></Dialog>
  </div>;
}
