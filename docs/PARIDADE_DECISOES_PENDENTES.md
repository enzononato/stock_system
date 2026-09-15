# Paridade com a branch redesign-frontend — o que falta decidir

Referente à branch `feat/paridade-funcional`. Tudo abaixo está **implementado e funcionando**
do jeito atual; nada aqui é defeito. São pontos onde o sistema atual tem **mais** que a
referência, e ficar literalmente idêntico a ela exigiria **remover** algo que hoje funciona.

O dono pediu para ver cada remoção antes. Esta é a lista.

## Grupo A — Colunas de tabela

A referência mostra menos colunas que nós. Para ficar idêntico, sairiam:

**Estoque** (`frontend/src/pages/StockPage.tsx`): `peripheral_count`, `identificador`,
`setor`, `ip`. A referência mostra esses quatro só na ficha lateral do item, não na grade.
Perda: deixa de dar para varrer esses campos sem abrir o detalhe de cada item.

**Relatório** (`frontend/src/pages/ReportPage.tsx`): `fornecedor`, `cargo`, `center_cost`,
`data_confirmacao`, `data_devolucao`, `details` — seis colunas que a referência não tem
(19 nossas contra 11 dela).

## Grupo B — Recursos de tabela

**Ordenação clicável por coluna.** Vem de graça do nosso `DataTable`. As tabelas da
referência são `<table>` estáticas, sem `onClick` no cabeçalho. Some de todas as telas.

**Busca global no relatório.** A referência não tem campo de busca nessa tela, só paginação.

## Grupo C — Indicadores do Estoque

Os quatro cartões no topo do Estoque (Total, Disponíveis, Em Empréstimo, Ações Pendentes).
A referência os tem **no Dashboard dela, não na tela de Estoque** — lá o Estoque mostra
apenas uma contagem textual no cabeçalho.

**Este caso mudou de figura:** antes desta rodada nós não tínhamos Dashboard, então os
cartões no Estoque eram o único lugar com essa informação. Agora temos Dashboard. Remover os
do Estoque passou a ser uma opção razoável em vez de perda pura — a informação continua a um
clique.

## Grupo D — Gráficos

A `<Legend>` do recharts no primeiro gráfico. A referência desenha uma legenda estática à mão
só para esse gráfico e deixa os outros dois **sem legenda nenhuma**. A nossa é interativa e
acompanha a troca de tema automaticamente. Ficar idêntico seria trocar por pior.

## Grupo E — Redesenhos de tela inteira

Estes não são remoções, são reconstruções — e foram deixados de fora porque você escolheu
"visual + componentes estruturais novos", não "reestruturar telas":

- **Empréstimo** como assistente de 4 passos, com trilho de contexto fixo de 340px à direita.
- **Devolução** com modal de checklist físico: três marcações obrigatórias (chassi íntegro,
  tela/teclado/fonte testados, periféricos conferidos) antes de liberar a confirmação.
  Atenção: isso é **mudança de regra de negócio**, não de layout.
- **Termos** com visualizador de documento em duas colunas, exibindo o clausulado por extenso
  na tela. Hoje o termo existe só em PDF, para gerar e baixar.
- **Histórico** como linha do tempo de duas colunas em vez da tabela.

## Grupo F — Shell

A sidebar da referência **esconde** `/remove`, `/terms` e `/link` — lá só se chega por Ctrl+K.
A nossa mostra os três, e agora eles aparecem nos dois lugares. E o logout dela não
redireciona; o nosso sim.

## Grupo G — O que ficou pendente por escopo, não por decisão

- **Exportação CSV** em Periféricos e em Histórico. Exigiria um utilitário novo de CSV.
  Barato de fazer, só não estava nos passos das tasks.
- **"Visualização por Equipamento"** em Periféricos: duplicaria integralmente a página
  `/link` e reintroduziria uma busca de 500 linhas.

---

## Defeitos da referência que NÃO foram copiados

Registrado para quem for comparar as duas versões e estranhar a diferença:

1. **Devolução contada no mês errado.** A referência agrupa `data_devolucao` por dia sem
   conferir o mês. Como o relatório mensal traz só empréstimos emitidos no mês, e
   `data_devolucao` vem de subconsulta sem limite de data, uma máquina emprestada em março e
   devolvida em abril era contada como devolução de abril **no gráfico de março**.
2. **403 para Jovem Aprendiz ao filtrar filial.** `GET /reports/monthly` exige gestor ou
   técnico, mas os endpoints de gráfico não têm guarda de papel. O usuário abria Indicadores
   normalmente e quebrava ao escolher uma unidade. Agora o seletor não é renderizado para
   esse papel.
3. **Empréstimos inflados.** A coluna `data_emprestimo` é reaproveitada pelo UNION do
   relatório para toda operação — cadastro, periférico, exclusão. Contar por ela ser
   preenchida inflava o total. Agora conta por `operation_type === 'Empréstimo'`.
4. **Tempestade N+1 calculando valor descartado.** Em Periféricos, a referência dispara uma
   requisição HTTP por item com periférico para montar um mapa que não é lido em lugar
   nenhum — só aparece na lista de dependências de um `useMemo`.
5. **Desalinhamento dia/valor** no recorte dos últimos 15 dias do gráfico do Dashboard.
6. **Indicadores errados acima de 500 itens.** A referência agrega sobre
   `listItemsPaginated({ limit: 500 })` no cliente, e 500 é o teto do backend, não folga.
   Aqui todo total vem do campo `total` de uma resposta paginada.

## Limitações reais do backend, que não dá para contornar sem alterá-lo

- `GET /api/peripherals` aceita só `status`, `tipo` e `include_inactive`. **Não tem `search`,
  `limit` nem `offset`** — por isso ali a busca textual e a paginação são no cliente.
- `GET /api/items` **não tem parâmetro `setor`**, então filtro de setor no servidor é
  impossível. A referência também não tem.
- **Filtrar o relatório por unidade descarta toda operação de periférico**, porque o ramo de
  periféricos do UNION seleciona `NULL` na coluna de revenda. A interface sinaliza isso
  quando o filtro está ativo.
- **A exportação CSV ignora os filtros de tela** — o endpoint recebe só ano e mês. A interface
  avisa que o arquivo traz sempre o mês inteiro.
- **"Devoluções" no relatório não existe como operação.** O UNION nunca gera essa linha; a
  devolução aparece só como a coluna `data_devolucao` preenchida na linha de Empréstimo. Por
  isso o bloco se chama "Empréstimos já devolvidos", com a qualificação "dos emitidos no
  período" — o número não é "devoluções que ocorreram no mês".
