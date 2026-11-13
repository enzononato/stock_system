-- Estado das tarefas periódicas do agendador (última execução e resultado).
-- ultima_execucao é sempre gravada em UTC.
CREATE TABLE IF NOT EXISTS scheduler_execucoes (
    nome VARCHAR(100) NOT NULL PRIMARY KEY,
    ultima_execucao DATETIME NULL,
    ultimo_status ENUM('ok','erro') NULL,
    ultimo_erro VARCHAR(500) NULL
) CHARACTER SET utf8mb4;
