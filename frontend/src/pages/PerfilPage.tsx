import { useState } from 'react'
import { useMutation, useQuery, useQueryClient } from '@tanstack/react-query'
import {
  confirmarCadastro2fa,
  desativar2fa,
  getPerfil,
  iniciarCadastro2fa,
  regenerarCodigos,
} from '@/api/perfil'
import { PageHeader, PanelHeader } from '@/components/layout/PageHeader'
import { CadastroAutenticador } from '@/components/auth/CadastroAutenticador'
import { CodigosRecuperacao } from '@/components/auth/CodigosRecuperacao'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { StatusPill } from '@/components/ui/StatusPill'
import { toast } from '@/components/ui/toast'
import { mensagemErroAuth } from '@/lib/auth-errors'

type Fluxo =
  | { tipo: 'nenhum' }
  | { tipo: 'cadastro' }
  | { tipo: 'codigos'; codigos: string[]; aoConcluir: () => void }

type AcaoComSenha = 'desativar' | 'codigos'

function Campo({ rotulo, children }: { rotulo: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-0.5">
      <dt className="text-caption font-semibold uppercase tracking-wider text-muted-foreground">{rotulo}</dt>
      <dd className="text-body-sm text-foreground">{children}</dd>
    </div>
  )
}

export default function PerfilPage() {
  const queryClient = useQueryClient()
  const { data: perfil, isLoading, isError } = useQuery({ queryKey: ['perfil'], queryFn: getPerfil })
  const [fluxo, setFluxo] = useState<Fluxo>({ tipo: 'nenhum' })
  const [acao, setAcao] = useState<AcaoComSenha | null>(null)
  const [senha, setSenha] = useState('')
  const [erroSenha, setErroSenha] = useState('')

  function fecharDialogo() {
    setAcao(null)
    setSenha('')
    setErroSenha('')
  }

  const senhaMutation = useMutation({
    mutationFn: async ({ qual, valor }: { qual: AcaoComSenha; valor: string }) => {
      if (qual === 'desativar') {
        await desativar2fa(valor)
        return null
      }
      return (await regenerarCodigos(valor)).codigos_recuperacao
    },
    onSuccess: (codigos, { qual }) => {
      fecharDialogo()
      if (qual === 'desativar') {
        void queryClient.invalidateQueries({ queryKey: ['perfil'] })
        toast('Verificação em duas etapas desativada.')
      } else if (codigos) {
        setFluxo({
          tipo: 'codigos',
          codigos,
          aoConcluir: () => {
            setFluxo({ tipo: 'nenhum' })
            void queryClient.invalidateQueries({ queryKey: ['perfil'] })
          },
        })
      }
    },
    onError: (err: unknown) => setErroSenha(mensagemErroAuth(err, 'Não foi possível concluir a operação.')),
  })

  if (isLoading) {
    return <p className="text-sm text-muted-foreground">Carregando...</p>
  }
  if (isError || !perfil) {
    return <p className="text-sm text-destructive">Não foi possível carregar o seu perfil.</p>
  }

  const podeDesativar = perfil.totp_ativo && !perfil.mfa_obrigatorio

  return (
    <div className="page-container-reading space-y-8">
      <PageHeader
        eyebrow="Conta"
        eyebrowDetail="Segurança"
        title="Meu perfil"
        description="Seus dados de acesso e a verificação em duas etapas."
      />

      <section className="surface-panel p-6 space-y-4">
        <PanelHeader title="Dados" />
        <dl className="grid grid-cols-1 gap-4 sm:grid-cols-3">
          <Campo rotulo="Usuário">{perfil.username}</Campo>
          <Campo rotulo="Função">{perfil.role}</Campo>
          <Campo rotulo="E-mail">{perfil.email ?? 'Não cadastrado'}</Campo>
        </dl>
        <p className="text-caption text-muted-foreground">Para alterar o e-mail, fale com um Gestor.</p>
      </section>

      <section className="surface-panel p-6 space-y-4">
        <PanelHeader
          title="Verificação em duas etapas"
          description="Além da senha, o login pede um código do aplicativo autenticador."
        />

        <div className="flex flex-wrap items-center gap-3">
          {perfil.totp_ativo ? (
            <StatusPill status="Ativo" />
          ) : (
            <StatusPill status="Inativo" tone="neutral" />
          )}
          {perfil.mfa_obrigatorio && (
            <span className="text-body-sm text-muted-foreground">Obrigatória para Gestor.</span>
          )}
          {perfil.totp_ativo && (
            <span className="text-body-sm text-muted-foreground">
              Códigos de recuperação restantes: {perfil.codigos_recuperacao_restantes}
            </span>
          )}
        </div>

        {!perfil.mfa_disponivel && (
          <p role="alert" className="text-body-sm text-destructive">
            A verificação em duas etapas está indisponível neste servidor. Avise o administrador.
          </p>
        )}

        {fluxo.tipo === 'cadastro' && (
          <div className="mx-auto max-w-sm rounded-lg border border-border p-5">
            <CadastroAutenticador
              iniciar={iniciarCadastro2fa}
              confirmar={async (codigo) => (await confirmarCadastro2fa(codigo)).codigos_recuperacao}
              onConcluido={(codigos) =>
                setFluxo({
                  tipo: 'codigos',
                  codigos,
                  aoConcluir: () => {
                    setFluxo({ tipo: 'nenhum' })
                    void queryClient.invalidateQueries({ queryKey: ['perfil'] })
                    toast('Verificação em duas etapas ativada.')
                  },
                })
              }
            />
          </div>
        )}

        {fluxo.tipo === 'codigos' && (
          <div className="mx-auto max-w-sm rounded-lg border border-border p-5">
            <CodigosRecuperacao
              codigos={fluxo.codigos}
              rotuloContinuar="Concluir"
              onContinuar={fluxo.aoConcluir}
            />
          </div>
        )}

        {fluxo.tipo === 'nenhum' && (
          <div className="flex flex-wrap gap-2">
            {!perfil.totp_ativo && perfil.mfa_disponivel && (
              <Button onClick={() => setFluxo({ tipo: 'cadastro' })}>
                Ativar verificação em duas etapas
              </Button>
            )}
            {perfil.totp_ativo && perfil.mfa_disponivel && (
              <Button variant="outline" onClick={() => setAcao('codigos')}>
                Gerar novos códigos de recuperação
              </Button>
            )}
            {podeDesativar && (
              <Button variant="outline" onClick={() => setAcao('desativar')}>
                Desativar
              </Button>
            )}
          </div>
        )}
      </section>

      <Dialog open={acao !== null} onOpenChange={(aberto) => !aberto && fecharDialogo()}>
        <DialogContent>
          <form
            onSubmit={(e) => {
              e.preventDefault()
              if (acao) senhaMutation.mutate({ qual: acao, valor: senha })
            }}
            className="space-y-4"
          >
            <DialogHeader>
              <DialogTitle>Confirme sua senha</DialogTitle>
              <DialogDescription>
                {acao === 'desativar'
                  ? 'Para desativar a verificação em duas etapas, digite a sua senha.'
                  : 'Os códigos atuais deixarão de valer. Digite a sua senha para gerar novos.'}
              </DialogDescription>
            </DialogHeader>
            <div className="space-y-1.5">
              <Label htmlFor="senha-confirmacao">Senha</Label>
              <Input
                id="senha-confirmacao"
                type="password"
                value={senha}
                onChange={(e) => setSenha(e.target.value)}
                autoComplete="current-password"
                autoFocus
              />
            </div>
            {erroSenha && (
              <p role="alert" className="text-body-sm text-destructive">
                {erroSenha}
              </p>
            )}
            <DialogFooter>
              <Button type="button" variant="outline" onClick={fecharDialogo}>
                Cancelar
              </Button>
              <Button type="submit" loading={senhaMutation.isPending} disabled={!senha}>
                Confirmar
              </Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  )
}
