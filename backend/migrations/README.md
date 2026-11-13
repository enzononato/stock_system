# Migrações do banco

Toda mudança nova de schema (tabelas, colunas, índices) entra aqui como migração
versionada. As tabelas legadas (`items`, `history`, `peripherals`,
`equipment_peripherals`, `revoked_tokens`, `unidades`) continuam criadas pelos
managers em `app/db/`; as migrações partem delas.

## Convenção de nomes

- `NNNN_descricao.sql` aplica a migração. `NNNN` tem 4 dígitos; a descrição usa
  minúsculas e `_`.
- `NNNN_descricao.down.sql` reverte a migração. O par é obrigatório: sem ele
  `python -m scripts.migrate reverter` recusa reverter aquela versão.
- Mudanças só aditivas: criar tabela, adicionar coluna anulável ou com padrão,
  criar índice. Nunca apagar nem renomear coluna ou tabela existente.
- Dentro do arquivo, cada comando termina com `;` no fim da linha. Linhas que
  começam com `--` são comentários.
- `CREATE TABLE` sem `CHARACTER SET`/`COLLATE`: a tabela nova herda o padrão do
  banco, como as legadas (evita "Illegal mix of collations" em junções com elas).
- Prefira `IF NOT EXISTS` / `IF EXISTS`: o MySQL faz commit implícito em DDL, então
  uma migração que falha no meio pode deixar parte aplicada, e reaplicá-la precisa
  funcionar.

## Faixas por fase

| Faixa | Fase |
|-------|------|
| 0001-0099 | Fase 0 (base: colaboradores, locais, auditoria, agendador) |
| 0100-0199 | Fase 1 (estoque e telas existentes; sem migração até agora) |
| 0200-0299 | Fase 2 (usuários: e-mail, 2FA, redefinição de senha) |
| 0300-0399 | Fase 3 (termos de responsabilidade) |
| 0400-0499 | Fase 4 (compra, manutenções, transferências, baixa) |
| 0500-0599 | Fase 5 (linhas telefônicas) |
| 0600-0699 | Fase 6 (inventário físico) |
| 0700-0799 | Fase 7 (painel, alertas e exportações; sem migração até agora) |
| 0800-0899 | Fase 8 (acessos) |

## Como roda

- No boot da API (`app/main.py`): `garantir_baseline()` e `aplicar_pendentes()`.
  Se o banco estiver fora do ar, o erro é registrado e a API sobe mesmo assim.
- Manualmente, a partir de `backend/`:
  - `python -m scripts.migrate` ou `python -m scripts.migrate aplicar`
  - `python -m scripts.migrate pendentes`
  - `python -m scripts.migrate reverter` (reverte só a última aplicada)
- Cada versão ausente de `schema_migrations` é aplicada em ordem crescente, mesmo
  que uma versão maior já esteja aplicada. Um `GET_LOCK` impede que dois processos
  migrem ao mesmo tempo.
