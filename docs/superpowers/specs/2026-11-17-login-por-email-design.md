# Login por e-mail — design

Data: 2026-11-17. Pedido do responsável pelo sistema; aprovado na conversa ("todos têm e-mail @revalle.com.br — pode fazer a implementação").

## Intenção
Hoje cada pessoa entra com um nome de usuário, e o responsável não acha isso profissional. Ele quer que todas as contas (Gestor, Técnico e Jovem Aprendiz) entrem com o e-mail corporativo `@revalle.com.br`.

As contas existentes não podem ser apagadas nem recriadas, para não perder o histórico (`history.usuario` guarda o nome de usuário). Quem já tem conta entra uma última vez com usuário e senha, cadastra e confirma o e-mail, e daí em diante entra com o e-mail.

Produção tem 7 contas, nenhuma com e-mail. Os Gestores ainda não cadastraram o 2FA.

## Objetivos
- Login com **e-mail + senha** para todos.
- **Migração no próximo login.** Uma conta sem e-mail confirmado entra com usuário e senha e é levada a informar o e-mail. Um código de 6 dígitos é enviado para esse e-mail, e a pessoa só entra no sistema depois de digitar o código. O 2FA do Gestor vem logo depois, no mesmo acesso.
- **A conta continua a mesma.** Mesmo `id` e mesmo `username`, que fica guardado por dentro e não aparece mais na tela de login. O histórico e a auditoria continuam mostrando quem fez o quê.
- **Só `@revalle.com.br`.** O domínio fica numa configuração (`LOGIN_EMAIL_DOMINIO`, padrão `revalle.com.br`), sem maiúsculas e com espaços removidos.
- **Contas novas** são criadas pelo Gestor com e-mail obrigatório. O `username` é gerado a partir da parte antes do `@`; se já existir, ganha um número no fim (`joao.silva2`).

## O que fica de fora
- Entrar sem senha (link mágico) e login com Google ou Microsoft.
- A pessoa trocar o próprio e-mail pela tela de Perfil. Quem troca é o Gestor, e a pessoa confirma no próximo login.
- Recriar contas ou mexer no histórico.

## O que tornaria a mudança um fracasso
- Alguém ficar trancado fora da conta. Isso aconteceria se um e-mail digitado errado passasse a valer, ou se um usuário antigo não tivesse mais como entrar.
- Dar para descobrir quais e-mails ou usuários existem pelas mensagens de erro, antes da senha ser conferida.
- Um Gestor conseguir pular o 2FA, ou alguém pular a confirmação do e-mail.

## Abordagens consideradas
1. **Escolhida: uma etapa nova no login em etapas que já existe.** O login já devolve `etapa` (`ok` | `mfa` | `cadastro_2fa`) com um token temporário de propósito único. Entra a etapa `cadastro_email`, antes do 2FA. Ela reaproveita o token de etapa, o limite de tentativas e o envio de e-mail.
2. Trocar o `username` pelo e-mail na mesma coluna. Foi descartada porque o histórico guarda o `username` como texto: os registros antigos ficariam apontando para nomes que não existem mais.
3. O Gestor preenche os e-mails de todo mundo no banco e desliga o login por usuário. Foi descartada porque um e-mail digitado errado deixa a pessoa trancada fora, e o responsável pediu que cada pessoa cadastre o próprio e-mail.

## Como funciona

**Login**

O campo da tela passa a se chamar "E-mail". O backend continua aceitando o campo `username` na requisição (compatibilidade) e decide assim:

| O que foi digitado | Como o backend trata |
|---|---|
| Tem `@` | Procura a conta pelo e-mail (sem maiúsculas). Vale com o e-mail confirmado ou não. |
| Não tem `@` | Procura pelo `username`, mas só aceita se a conta **não** tem e-mail confirmado. |

- Se a conta já tem e-mail confirmado e a pessoa usou o usuário, a resposta é 401 "Esta conta agora entra com o e-mail. Use o seu e-mail @revalle.com.br." Essa mensagem só aparece depois da senha conferida, então não revela que a conta existe.
- Erros antes da senha continuam com a mensagem genérica atual. O tempo de resposta continua o mesmo para contas que existem e que não existem (o bcrypt de mentira continua).
- Ordem das etapas depois da senha certa:
  1. sem e-mail confirmado → `cadastro_email`;
  2. senão, Gestor com 2FA ativo → `mfa`;
  3. senão, Gestor sem 2FA → `cadastro_2fa`;
  4. senão → `ok`.

**Etapa `cadastro_email`**

A resposta traz um `mfa_token` de propósito `"email"` e o e-mail atual, se houver, para preencher o campo.

- `POST /api/auth/email/iniciar {mfa_token, email}`:
  - confere o domínio, o formato e se outra conta já usa esse e-mail (409 "Este e-mail já está em uso por outra conta.");
  - gera um código de 6 dígitos e grava só o hash, vinculado à conta e ao e-mail, válido por 15 minutos;
  - envia o código por e-mail;
  - pode reenviar depois de 60 segundos; o código anterior deixa de valer.
- `POST /api/auth/email/confirmar {mfa_token, codigo}`:
  - com o código certo, grava `usuarios.email` e `email_confirmado = 1`, apaga os códigos pendentes e registra na auditoria (`usuario` / `confirmar_email`, com o e-mail mascarado);
  - em seguida segue a ordem de etapas acima: Gestor sem 2FA → `cadastro_2fa` com um token novo; senão abre a sessão.
- Limites:
  - 5 códigos errados invalidam o código e o token da etapa;
  - os dois endpoints também têm o limite por IP do login.

**Dados**

Migração aditiva nova, com `.down.sql`, na faixa livre seguinte:
- `usuarios.email_confirmado TINYINT(1) NOT NULL DEFAULT 0`. Todas as contas atuais começam em 0.
- Tabela `email_confirmacoes`: `id`, `usuario_id` (FK), `email`, `codigo_hash`, `expira_em`, `tentativas`, `criado_em`.
- Sem charset ou collation fixos (herda o padrão do banco, Ruling 39).

**Gestor (tela Usuários)**
- Ao criar uma conta, informa e-mail, senha e papel. O `username` é gerado sozinho.
- Ao editar o e-mail de alguém, o `email_confirmado` volta para 0 e a pessoa confirma no próximo login. A lista mostra o e-mail e se está confirmado.
- O `username` continua aparecendo na lista, como informação secundária.

**Outras telas**
- "Esqueci minha senha" só manda o link para e-mail **confirmado**.
- O cabeçalho e o perfil mostram o e-mail.

**Tela de login**
- Campo "E-mail", com a dica: "Primeiro acesso depois da mudança? Entre uma vez com o seu usuário antigo."
- Nova etapa com o campo de e-mail, o botão "Enviar código", o campo do código e "Reenviar código".

## Restrições
- Seguir as convenções do projeto: mensagens em português, `audit.record`, componentes da F0, TDD e migrações aditivas.
- Os fluxos de 2FA, de sessão (`sessao_versao`) e de redefinição de senha continuam funcionando como hoje.

## Decidido por mim (o responsável pode corrigir)
- Código de 6 dígitos, válido por 15 minutos, com reenvio após 60 segundos e 5 tentativas.
- O `username` é gerado da parte antes do `@`.
- Trocar o próprio e-mail pelo Perfil fica fora desta entrega.
