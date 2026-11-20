-- Coletor de dados da máquina: cada coleta é um código de uso único e curto (15 min).
-- Só o SHA-256 do código fica no banco (codigo_hash); datas em UTC, passadas pelo Python.
-- `dados` guarda a coleta bruta (`ColetaBruta.model_dump()`) enviada pelo script, sem normalizar; nada é gravado em `items`.
CREATE TABLE IF NOT EXISTS coletas_maquina (
    id INT AUTO_INCREMENT PRIMARY KEY,
    codigo_hash CHAR(64) NOT NULL UNIQUE,
    criado_por_id INT NOT NULL,
    criado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    expira_em DATETIME NOT NULL,
    recebido_em DATETIME NULL,
    recebido_ip VARCHAR(45) NULL,
    dados JSON NULL,
    INDEX idx_coletas_criado_em (criado_em),
    CONSTRAINT fk_coletas_usuario FOREIGN KEY (criado_por_id) REFERENCES usuarios(id) ON DELETE CASCADE
);
