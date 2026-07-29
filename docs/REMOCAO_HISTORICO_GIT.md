# Remoção de dados pessoais do histórico do Git

> **Este documento é apenas referência. NÃO execute os comandos abaixo sem
> planejamento e comunicação prévia com todos os envolvidos** — eles reescrevem
> o histórico inteiro do repositório e exigem `push --force`.

## Contexto

Em 2026-07-28, as pastas `terms/`, `termos_assinados/`, `termos_devolucao/`,
`termos_devolucao_assinados/` e `notas_remocao/` foram **desrastreadas**
(`git rm -r --cached`) porque contêm documentos reais de colaboradores: PDFs e
DOCX com CPF, nome completo, currículos e boletos.

Desrastrear resolve o problema **daqui para frente** (novos commits não vão
mais incluir esses arquivos). Mas os arquivos que já foram commitados
continuam recuperáveis por qualquer pessoa com acesso ao histórico do
repositório — em cada commit anterior ao da remoção, em `git log -p`, em
`git show <sha>:<caminho>`, em clones antigos, em forks, etc. Removê-los de
verdade exige reescrever o histórico.

## Quando fazer isso

Só depois de:
1. Confirmar com o time que ninguém depende de commits antigos que serão
   reescritos (SHAs vão todos mudar).
2. Ter um backup completo do repositório (`.git` incluso) antes de começar.
3. Escolher uma janela sem pushes concorrentes de outras pessoas.

## Passo a passo

### 1. Instalar o git-filter-repo

`git-filter-repo` é a ferramenta recomendada pelo próprio time do Git para
esse tipo de operação (substitui o antigo `git filter-branch`, que é lento e
cheio de armadilhas, e o BFG Repo-Cleaner, menos flexível para filtrar por
caminho).

```bash
# Via pip (recomendado, multiplataforma)
python -m pip install git-filter-repo

# Ou via gerenciador de pacotes
# Debian/Ubuntu:
sudo apt install git-filter-repo
# macOS:
brew install git-filter-repo
```

### 2. Trabalhar em um clone novo e isolado

O `git-filter-repo` se recusa a rodar em um clone "sujo" (com remotes,
branches de trabalho, etc.) por segurança. Faça um clone novo só para essa
operação:

```bash
git clone <url-do-repositorio> stock_system_purge
cd stock_system_purge
```

### 3. Remover as 5 pastas de TODO o histórico

```bash
git filter-repo \
  --path terms \
  --path termos_assinados \
  --path termos_devolucao \
  --path termos_devolucao_assinados \
  --path notas_remocao \
  --invert-paths
```

- `--path <pasta>` seleciona os caminhos afetados.
- `--invert-paths` inverte a seleção: em vez de "manter só isso", vira
  "remover isso e manter o resto" — ou seja, os PDFs/DOCX somem de **todos**
  os commits onde aparecem, do primeiro ao último.

O `git-filter-repo` já roda a limpeza de objetos soltos e o `gc` internamente,
então ao final os blobs antigos deixam de existir no repositório local.

### 4. Conferir o resultado antes de publicar

```bash
# Não deve retornar nada:
git log --all --diff-filter=A -- terms termos_assinados termos_devolucao termos_devolucao_assinados notas_remocao

# Os SHAs de todos os commits mudaram — isso é esperado e inevitável.
git log --oneline | head -20
```

### 5. Publicar (⚠️ reescreve o histórico remoto)

O `git-filter-repo` remove o remote `origin` automaticamente (proteção contra
push acidental). Adicione de novo e force o push:

```bash
git remote add origin <url-do-repositorio>
git push origin --force --all
git push origin --force --tags
```

**A partir daqui, os SHAs antigos deixam de existir no remoto.** Qualquer PR
aberto, branch não mesclada ou fork baseado no histórico antigo vai divergir
ou quebrar.

## Checklist de coordenação (fazer ANTES do push --force)

- [ ] Avisar todos os colaboradores com clone local da data/hora da operação.
- [ ] Confirmar que não há PRs abertos importantes que serão invalidados (ou
      documentar que precisarão ser recriados após a reescrita).
- [ ] Combinar um horário sem commits concorrentes de terceiros.
- [ ] Depois do push --force, instruir cada colaborador a **descartar o clone
      antigo e clonar de novo do zero** — não tentar `git pull`/`git rebase`
      em cima do histórico antigo (na prática gera conflitos e pode
      reintroduzir os blobs removidos).
- [ ] Se o repositório estiver hospedado em GitHub/GitLab, verificar se há
      forks públicos — eles preservam os blobs antigos independentemente do
      que acontece no repositório de origem. Nesse caso, também vale abrir um
      chamado de suporte da plataforma pedindo a purga de caches internos
      (diffs de PR, releases, etc.) que possam ter uma cópia dos arquivos.
- [ ] Revalidar pipelines de CI/CD que referenciam SHAs fixos.

## Nota importante e independente: rotação de credenciais

A limpeza acima cobre **apenas os documentos pessoais** (PDFs/DOCX). Ela **não
resolve** o fato de que a senha do MySQL (`DB_PASSWORD`) e o host do banco
ficaram commitados como valor padrão em `backend/app/core/config.py` e em
`backend/.env.example` em commits anteriores a este módulo de refatoração.

Reescrever o histórico é opcional e disruptivo; **trocar a senha não é
opcional**. Portanto, **independentemente de este expurgo de histórico ser
executado ou não, e independentemente de quando for executado**:

1. **Rotacione a senha do usuário MySQL no servidor real** (a que estava
   hardcoded no código). Uma vez que uma credencial esteve exposta em texto
   plano em um repositório Git — mesmo privado, mesmo que só localmente por um
   tempo — ela deve ser tratada como comprometida.
2. Atualize o `backend/.env` local (fora do controle de versão) com a nova
   senha.
3. Se o mesmo usuário/senha for reaproveitado em outro sistema, troque lá
   também.

Se, além dos documentos, também for necessário remover a senha antiga do
histórico de commits (não apenas trocá-la no servidor), o mesmo
`git-filter-repo` pode ser usado com `--replace-text` apontando para um
arquivo com o valor antigo, em uma segunda operação sobre o mesmo clone antes
do push. Isso está fora do escopo deste documento porque o valor exposto já
precisa ser considerado inválido a partir da rotação no passo 1 acima — a
reescrita de histórico neste caso é apenas higiene adicional, não a correção
em si.

## Referência: schema legado da tabela `usuarios`

Preservado aqui porque o script original (`inicializar_db.py`, removido da
raiz do repositório) foi apagado como parte da limpeza do app desktop legado.
O schema completo, com o comentário de contexto, também está documentado no
`README.md` (seção "Schema herdado"):

```sql
CREATE TABLE IF NOT EXISTS usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role ENUM('Gestor', 'Técnico', 'Jovem Aprendiz') NOT NULL
)
```
