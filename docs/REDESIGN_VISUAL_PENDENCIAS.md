# Port visual monocromático — pendências e decisões em aberto

Referente ao merge `92b53de` (branch `feat/port-visual-redesign`, 26 commits).

O port está completo e verificado: `tsc` limpo, 133 testes passando em 13 arquivos,
`vite build` limpo, `docker compose build frontend` funcionando. Nada abaixo é um defeito
que impeça o uso — são decisões de design que cabem a você, mais uma verificação que
nenhum agente consegue fazer.

## 1. Ninguém olhou o sistema rodando

Esta é a lacuna que mais pesa. Toda a verificação foi estática: compilador, suíte de
testes e análise de código. O plano previa uma passagem visual humana como último passo e
ela não aconteceu.

Para fazer:

```bash
docker compose -f docker-compose.test.yml up -d
cd frontend && npm run dev
```

Olhe, no mínimo: o login (duas colunas), o estoque (tabela densa), o botão de tema
claro/escuro na barra superior, um modal de detalhes de item, e a tela de remoção
(contraste de ação destrutiva). Depois derrube o banco com
`docker compose -f docker-compose.test.yml down -v`.

## 2. Contraste das linhas divisórias

O design separa superfícies por linha, não por sombra. As linhas medem:

| par | claro | escuro |
|---|---|---|
| `--border` sobre `--surface` | 1.42:1 | 1.31:1 |
| `--border-strong` sobre `--surface` | 1.89:1 | 1.72:1 |

Duas consequências: um campo de formulário difere do painel em que está apenas por essa
linha (a WCAG 1.4.11 pede 3:1 para identificar um controle), e `figure-ground-panel` —
usado em toda decisão crítica, incluindo estorno e remoção — difere de um painel comum só
por trocar `--border` por `--border-strong`.

Não alterei porque a paleta é cópia literal do design de referência que você escolheu, e
escurecer esses tokens muda a aparência de todo painel, todo campo e toda tela de
confirmação do sistema. Se quiser mudar, os valores sugeridos são `#8A8A84` no claro e
`#616161` no escuro, mais apontar `--input` para `--border-strong` e dar 2px de borda ao
`figure-ground-panel`. Tudo em `frontend/src/index.css`.

## 3. Cores das operações no histórico

`OperationBadge` em `frontend/src/pages/HistoryPage.tsx` passou de quatro códigos de cor
para dois: só `Exclusão` e `Estorno` ficaram coloridos (vermelho), todo o resto é neutro.

A favor: a regra do design é que cor marca status, e tipo de operação não é status.
Contra: a tabela de auditoria é a mais densa do sistema (20 linhas por página) e era
justamente onde a cor acelerava a leitura. Reversível em duas linhas.

## 4. Botão "Desvincular" preenchido em cada linha

Em `frontend/src/pages/LinkPeripheralPage.tsx`, o botão destrutivo virou preenchido
(sólido). Como ele aparece em toda linha de periférico vinculado, uma lista com vários
itens fica com vários botões vermelhos sólidos. Desvincular é reversível — talvez mereça
a versão só com contorno, deixando o preenchimento para a confirmação final.

## 5. "Confirmar Renomeação" pintado como exclusão

`AlertDialogAction` fixa o variant destrutivo, então o diálogo de renomear unidade em
`frontend/src/pages/UnidadesPage.tsx` mostra um botão vermelho sólido para uma ação que
não apaga nada. Correção de uma linha no próprio call site, sem mexer no componente
(o `cn` com tailwind-merge deixa o call site sobrescrever):

```tsx
<AlertDialogAction className="bg-primary text-primary-foreground" onClick={confirmRename}>
  Confirmar Renomeação
</AlertDialogAction>
```

## Armadilha que vale conhecer antes de mexer no CSS

Os tokens de cor estão registrados no `tailwind.config.js` como string `var(--x)`, não no
formato `rgb(var(--x) / <alpha-value>)`. Por isso **modificador de opacidade sobre token
deste projeto não gera CSS nenhum** — `bg-primary/50` é descartado em silêncio, sem erro
de build, sem warning, e o elemento renderiza sem cor. Quatro ocorrências assim chegaram a
existir neste port (um ponto separador invisível no login, um spinner que não girava).

Existe um teste-guarda em `frontend/src/lib/tokens.test.ts` que deriva a lista de tokens
do próprio config e reprova apontando `arquivo:linha:classe`. Se precisar de uma cor mais
clara, use um token real (`text-muted-foreground`, `border-border`, `border-border-strong`)
em vez da opacidade.
