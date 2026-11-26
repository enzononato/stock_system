# Coletor de dados da máquina — design

Data: 2026-11-19. Pedido do responsável pelo sistema; abordagem escolhida na conversa ("sim" para o comando no Win+R, com o download como alternativa).

## Intenção
Hoje o técnico digita à mão os dados de cada computador no cadastro do Patrimônio (marca, modelo, host, MAC, processador, memória, disco, sistema, licença, AnyDesk). Isso é lento e dá erro de digitação. Existia um script de outra pessoa que fazia a coleta no computador, mas ninguém tem acesso a ele (era o item 8.3 do cartão do Trello). O responsável quer "a forma mais automática possível, sem o técnico precisar fazer tanta coisa".

Quando é usado:
- **Máquina nova**: o técnico instala e configura tudo (Windows, AnyDesk, licença, domínio) e, no fim, abre o Patrimônio no navegador daquela máquina para cadastrá-la.
- **Máquina que já está na empresa**: o técnico passa nela para colocá-la no Patrimônio ou completar os dados.

Todas as máquinas são Windows (10 ou 11). Linux, se aparecer, o responsável avisa.

## O que existe hoje (relevante)
- Cadastro em `/register` (Gestor e Técnico), `RegisterItemPage.tsx`; os campos por tipo ficam em `TypeSpecificFields.tsx` (`TYPE_FIELDS`). Notebook e Desktop mostram: Domínio (Sim/Não), Host, Endereço Físico (MAC), Processador, Memória RAM, Armazenamento, Sistema Operacional, Licença Windows e AnyDesk. **Não mostram número de série**, embora a coluna `items.identificador` exista e seja usada como serial em Celular, Tablet e Nobreak.
- Produção (19/11): 74 notebooks e 34 desktops; nenhum tem serial; todos têm `endereco_fisico` (MAC) e `anydesk`; `licenca` guarda o **Product ID** do Windows (`00355-62588-…`), não a chave. Formatos usados à mão: CPU `i5-1345U`, RAM `16GB`, armazenamento `512 SSD`, sistema `Windows 11`.
- MAC e IP já passam por `normalizar_mac` / `validar_ip` no backend e por máscara/validação no formulário.
- Sessão: access token em memória e refresh token em cookie httpOnly (persiste no navegador até o logout).
- Limite por IP com `slowapi` (`limiter` em `app/dependencies.py`). Agendador com jobs em `app/services/scheduler.py`. Migrações aditivas em `backend/migrations/NNNN_*.sql` + `.down.sql`.

## Objetivos
- Na máquina que está sendo cadastrada, o técnico faz: entrar no Patrimônio → "Cadastrar esta máquina" → **Copiar** → `Windows + R` → `Ctrl+V` → `Enter`. Em poucos segundos a página mostra o cadastro preenchido.
- O técnico **sempre confere e confirma**; o coletor nunca grava um equipamento sozinho.
- Unidade, local, colaborador, nota fiscal e dados de compra continuam com o técnico.
- Máquina já cadastrada é reconhecida e o técnico atualiza aquele cadastro em vez de criar outro.
- Ao terminar, a página **sai da conta** do técnico naquele computador.

## Não objetivos
- Linux e macOS; celulares, tablets, impressoras e periféricos (monitor, teclado).
- Coleta sem login do técnico; agente instalado que atualiza sozinho de tempos em tempos.
- Executável assinado ou instalador.
- Guardar a chave do Windows (só o Product ID, como já é feito).

## O que tornaria isso um fracasso
- O acesso do técnico ficar salvo no computador que vai para outra pessoa.
- Equipamento duplicado porque a máquina já estava cadastrada.
- Precisar de permissão de administrador, instalar algo ou passar por vários avisos do Windows a cada máquina.
- Alguém de fora conseguir mandar dados para o Patrimônio ou ler os dados coletados.
- Dado errado gravado sem o técnico ver (o preenchimento substitui o que ele vê na tela antes de salvar, nunca o banco direto).

## Abordagens consideradas
1. **Escolhida: comando colado no Win+R, com download como alternativa.** A página gera um código de uso único e mostra um comando curto (`powershell … irm <url com o código> | iex`). Não baixa arquivo, não aparece aviso do SmartScreen, são três teclas. Na mesma página, "Baixar o coletor" entrega um `.cmd` com o mesmo conteúdo, para quando o comando for bloqueado; ele mostra os avisos de arquivo baixado (3–4 cliques a mais).
2. Só o arquivo para baixar e dar dois cliques: familiar, mas sem assinatura digital o navegador e o Windows avisam em toda máquina.
3. Programa assinado: sem avisos, mas exige certificado de assinatura de código pago (centenas de dólares por ano) e um processo de build; não compensa para uso interno.
4. Só o navegador: descartado. O navegador não dá acesso a serial, marca, modelo, host, MAC, disco, licença nem AnyDesk.

## Como funciona

### Fluxo do técnico
```
/register  ──[ Preencher com os dados desta máquina ]──►  /register/coletor
                                                         1. [ Copiar comando ]
                                                         2. Windows + R, Ctrl+V, Enter
                                                         ⏳ Aguardando os dados…   (alternativa: Baixar o coletor)
            ◄── dados chegam ──
   sem correspondência  → /register com o tipo sugerido e os campos preenchidos
   com correspondência  → "Esta máquina parece já estar cadastrada: #123 Notebook Dell Latitude 3420 (mesmo MAC)."
                          [ Atualizar este cadastro ] → /edit/123 com os campos coletados aplicados
                          [ Cadastrar como novo mesmo assim ] → /register preenchido
   salvar → tela "Máquina cadastrada (#id). Você saiu da sua conta neste computador."
```
- O botão "Preencher com os dados desta máquina" fica no topo do cadastro (modo criação). A página `/register/coletor` também tem link direto no menu de cadastro, para o técnico favoritar.
- A página consulta o andamento a cada 2 s enquanto está aberta. Código vencido: mostra "O código venceu" e o botão "Gerar novo código".
- Nos campos preenchidos pelo coletor aparece uma marca discreta ("do coletor"); no modo de atualização, os campos cujo valor mudou em relação ao cadastro atual ficam destacados, para o técnico ver o que vai trocar.
- **Saída automática da conta** (decidido pelo responsável: "sair logo ao salvar"): quando o cadastro ou a atualização foi aberto pelo coletor, salvar com sucesso chama o logout existente (revoga só o refresh token daquele navegador) e mostra a tela final. A entrega ao colaborador (`/loan`, que gera o termo) e a etiqueta (`/etiquetas`) ficam para depois, no computador do técnico; a tela final lembra disso e mostra o número do equipamento. A página do coletor também tem "Sair deste computador" para quando o técnico desiste. Fora desse fluxo, salvar continua como hoje.

### Código de uso único
- `POST /api/coletor` (Gestor e Técnico) cria a coleta e devolve `{id, comando, url_arquivo, expira_em}`. O código (`secrets.token_urlsafe(16)`) só aparece nessa resposta; o banco guarda o SHA-256.
- Validade de **15 minutos**. O código serve para **um** envio de dados; depois disso as rotas públicas o tratam como vencido.
- O comando precisa caber no Win+R (limite de 259 caracteres) e forçar TLS 1.2 para Windows 10 mais antigo:
  `powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';irm '<APP_PUBLIC_URL>/api/coletor/<codigo>'|iex"`
  Um teste garante o tamanho com o `APP_PUBLIC_URL` de produção.

### Rotas públicas (sem login; o código é a autorização)
Todas com limite por IP (`30/minute`) e sem revelar nada além de "válido" ou "vencido".
- `GET /api/coletor/{codigo}` → `text/plain`, o script PowerShell com a URL de envio e o código embutidos. Código inválido ou vencido → 200 com um script de uma linha que escreve "Este código venceu. Gere um novo na página do Patrimônio." e espera Enter (para o técnico não ver um erro vermelho).
- `GET /api/coletor/{codigo}/coletor.cmd` → o mesmo, embrulhado num `.cmd` (`Content-Disposition: attachment`), que chama o PowerShell com `-ExecutionPolicy Bypass`.
- `POST /api/coletor/{codigo}/dados` → JSON da coleta (abaixo), até 64 KB. Valida tipos e tamanhos, normaliza e grava. Responde 200 `{ok: true}`; 410 se vencido ou já usado; 422 se o corpo for inválido.

### Rota do técnico
- `GET /api/coletor/{id}/andamento` (o usuário que criou ou um Gestor) → `{status: "aguardando" | "recebido" | "vencido", campos?, tipo_sugerido?, correspondencias?}`.
  - `campos`: já no formato do formulário (chaves de `ItemCreate`): `brand`, `model`, `identificador`, `dominio`, `host`, `endereco_fisico`, `cpu`, `ram`, `storage`, `sistema`, `licenca`, `anydesk`.
  - `correspondencias`: equipamentos existentes (qualquer status) com o mesmo serial (`identificador`, sem diferenciar maiúsculas e espaços), o mesmo MAC (`endereco_fisico` ou `mac`, normalizados, comparados com **todos** os MACs físicos da máquina, cabo e Wi-Fi) ou o mesmo ID do AnyDesk (só dígitos). Cada uma traz `id`, `tipo`, `brand`, `model`, `status`, `revenda`, `ativo` e `motivos` (`serial`/`mac`/`anydesk`).
  - **Equipamento baixado** (`is_active=0`) não oferece "Atualizar este cadastro": a página mostra "Esta máquina consta como baixada (#123)", com link para o registro antigo, e só oferece "Cadastrar como novo". Reativar uma baixa fica fora desta entrega.

### O que o coletor lê (PowerShell 5.1, sem administrador)
Cada leitura fica num `try/catch`; o que falhar vai vazio e o resto segue.

| Dado | Fonte | Vira no formulário |
|---|---|---|
| Fabricante | `Win32_ComputerSystem.Manufacturer` | `brand`, normalizado: "Dell Inc." → Dell, "LENOVO" → Lenovo, "HP"/"Hewlett-Packard" → HP, "Daten Tecnologia Ltda" → Daten (sem "Inc.", "Ltda", "Tecnologia", "Corporation", "Co., Ltd."); demais em formato de título |
| Modelo | Lenovo: `Win32_ComputerSystemProduct.Version` (nome comercial); demais: `Win32_ComputerSystem.Model` | `model` |
| Serial | `Win32_BIOS.SerialNumber`; descarta valores genéricos ("To be filled by O.E.M.", "Default string", "System Serial Number", "0", vazio) | `identificador` |
| Formato | `Win32_SystemEnclosure.ChassisTypes` (8, 9, 10, 14, 30, 31, 32 → notebook; 3–7, 13, 15, 16, 35, 36 → desktop) | `tipo_sugerido` (Notebook/Desktop; outro → técnico escolhe) |
| Host | `Win32_ComputerSystem.Name` | `host` |
| Domínio | `Win32_ComputerSystem.PartOfDomain` (+ `Domain`, só para exibir) | `dominio` = "Sim"/"Não" |
| Processador | `Win32_Processor.Name` | `cpu`, encurtado: Intel Core → `i5-1345U`, Core Ultra → `Ultra 5 125U`; AMD Ryzen → `Ryzen 5 5500U`; outros sem "(R)", "(TM)", "CPU" e "@ x GHz" |
| Memória | soma de `Win32_PhysicalMemory.Capacity` | `ram` = `16GB` |
| Disco | `Get-PhysicalDisk` (MediaType, Size, BusType; ignora USB) | `storage` = `512 SSD`; vários discos: `512 SSD + 1000 HDD`; tamanho arredondado ao comercial (238 GiB → 256) |
| Sistema | `Win32_OperatingSystem.Caption` sem "Microsoft" | `sistema` = `Windows 11 Pro` |
| Licença | `Win32_OperatingSystem.SerialNumber` (Product ID) | `licenca` |
| AnyDesk | `C:\ProgramData\AnyDesk\system.conf` (`ad.anynet.id`); se não achar, `anydesk.exe --get-id` no caminho padrão | `anydesk` (só dígitos) |
| MAC | todos os adaptadores físicos (`Get-NetAdapter -Physical`) vão no JSON; o gravado é o **que está em uso na hora** (decidido pelo responsável): o adaptador ativo com gateway padrão (`Get-NetIPConfiguration`); sem gateway, o primeiro Ethernet físico ativo, depois Wi-Fi | `endereco_fisico` (`AA:BB:CC:DD:EE:FF`); na atualização, se mudar, aparece destacado como os demais campos |

O JSON enviado é o dado bruto (`fabricante`, `modelo`, `versao_produto`, `serial`, `chassi`, `host`, `no_dominio`, `dominio`, `cpu`, `ram_bytes`, `discos[]`, `sistema`, `product_id`, `anydesk`, `adaptadores[]`, `versao_coletor`). **A normalização para o formato do formulário fica no backend**, testável em Python com amostras de Dell, Lenovo e HP; o script só lê e envia.

Janela do coletor: mostra "Coletando os dados desta máquina para o Patrimônio…", depois "Pronto! Volte ao navegador." e fecha sozinha em 5 s. Se o envio falhar, mostra a mensagem do servidor e espera Enter.

### Dados
Migração `1000_coletas_maquina.sql` (+ `.down.sql`), aditiva, sem charset fixo:
- `coletas_maquina (id, codigo_hash CHAR(64) UNIQUE, criado_por_id INT FK usuarios, criado_em, expira_em, recebido_em NULL, recebido_ip VARCHAR(45) NULL, dados JSON NULL)`.
- Um job diário do agendador apaga coletas com mais de 7 dias.

### Formulário
- `TYPE_FIELDS` de Notebook e Desktop ganha `{ key: 'identificador', label: 'Nº de Série' }`, logo depois de Host. Só tela; a coluna já existe.

## Restrições
- Convenções do projeto: mensagens em português, `audit.record` (o salvamento já audita; a coleta não precisa de auditoria própria), componentes da F0, TDD, migração aditiva.
- O script servido não pode conter nada além da URL pública e do próprio código; não guarda arquivo no disco; não pede administrador.
- Dados recebidos pela rota pública são não confiáveis: tamanho máximo por campo igual ao da coluna de destino, sem HTML, e nunca gravados em `items` sem o técnico salvar.

## Decidido por mim (o responsável pode corrigir)
- Validade de 15 minutos e uso único.
- Reconhecer máquina já cadastrada por serial, MAC ou AnyDesk (o serial sozinho não basta: nenhuma máquina em produção tem serial ainda).
- Campos vindos do coletor substituem os da tela, com destaque no que muda; o técnico pode editar antes de salvar.
- Logout automático só quando o cadastro foi aberto pelo coletor (o momento, "logo ao salvar", foi o responsável quem escolheu).
- Correspondência com equipamento baixado só permite "Cadastrar como novo".
- Formatos curtos iguais aos que já estão no banco (`i5-1345U`, `16GB`, `512 SSD`).
- IP não é guardado (Notebook/Desktop não têm esse campo na tela).

## Teste de viabilidade (19/11)
No PC do responsável (no domínio `revalle.local`, Windows Defender, 1 regra ASR), `irm <url> | iex` com um script inofensivo rodou sem bloqueio, e as leituras (fabricante, modelo, serial, disco, configuração do AnyDesk) funcionaram sem administrador. Não deu para testar, de dentro do ambiente do assistente, a abertura de um novo `powershell.exe` pelo Win+R. **A primeira tarefa do plano é esse teste numa máquina real da empresa**, com uma rota provisória que só devolve um script inofensivo, antes de construir as telas. Se for bloqueado, a abordagem volta para discussão.

## Deixado para quem constrói
- Nomes de componentes e arquivos do frontend, desde que sigam as pastas por módulo (`features/coletor/`).
- Como o template do script é montado no backend (arquivo em `app/templates/coletor/`, com substituição segura da URL e do código).
- Uma forma de testar o script localmente sem enviar (ex.: parâmetro que só imprime o JSON).
