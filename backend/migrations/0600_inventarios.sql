-- Inventário físico por unidade: cabeçalho, leituras e fotografia dos não encontrados.
-- A coluna gerada unidade_aberta garante um único inventário aberto por unidade.
CREATE TABLE IF NOT EXISTS inventarios (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    unidade_id INT NOT NULL,
    status ENUM('aberto','encerrado','cancelado') NOT NULL DEFAULT 'aberto',
    iniciado_por VARCHAR(100) NOT NULL,
    iniciado_em DATETIME NOT NULL,
    encerrado_por VARCHAR(100) NULL,
    encerrado_em DATETIME NULL,
    observacao VARCHAR(500) NULL,
    motivo_cancelamento VARCHAR(255) NULL,
    unidade_aberta INT AS (IF(status = 'aberto', unidade_id, NULL)) STORED,
    UNIQUE KEY uq_inventario_aberto_por_unidade (unidade_aberta),
    INDEX idx_inventarios_unidade_status (unidade_id, status),
    INDEX idx_inventarios_iniciado_em (iniciado_em),
    CONSTRAINT fk_inventarios_unidade FOREIGN KEY (unidade_id) REFERENCES unidades(id)
);
CREATE TABLE IF NOT EXISTS inventario_leituras (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    inventario_id INT NOT NULL,
    item_id INT NULL,
    chave VARCHAR(120) NOT NULL,
    codigo_lido VARCHAR(255) NOT NULL,
    lido_por VARCHAR(100) NOT NULL,
    lido_em DATETIME NOT NULL,
    local_id_encontrado INT NULL,
    usuario_conferido TINYINT(1) NOT NULL DEFAULT 0,
    colaborador_id_encontrado INT NULL,
    revenda_item VARCHAR(100) NULL,
    local_id_esperado INT NULL,
    colaborador_id_esperado INT NULL,
    usuario_esperado VARCHAR(200) NULL,
    resultado ENUM('ok','local_errado','usuario_diferente','outra_unidade','nao_cadastrado') NOT NULL,
    detalhe VARCHAR(500) NULL,
    observacao VARCHAR(500) NULL,
    atualizado_em DATETIME NULL,
    UNIQUE KEY uq_leitura_item (inventario_id, item_id),
    UNIQUE KEY uq_leitura_chave (inventario_id, chave),
    CONSTRAINT fk_leituras_inventario FOREIGN KEY (inventario_id) REFERENCES inventarios(id) ON DELETE CASCADE,
    CONSTRAINT fk_leituras_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE SET NULL,
    CONSTRAINT fk_leituras_local FOREIGN KEY (local_id_encontrado) REFERENCES locais(id) ON DELETE SET NULL,
    CONSTRAINT fk_leituras_colaborador FOREIGN KEY (colaborador_id_encontrado) REFERENCES colaboradores(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS inventario_nao_encontrados (
    inventario_id INT NOT NULL,
    item_id INT NOT NULL,
    status_item VARCHAR(50) NULL,
    local_id INT NULL,
    colaborador_id INT NULL,
    usuario VARCHAR(200) NULL,
    PRIMARY KEY (inventario_id, item_id),
    CONSTRAINT fk_nao_encontrados_inventario FOREIGN KEY (inventario_id) REFERENCES inventarios(id) ON DELETE CASCADE,
    CONSTRAINT fk_nao_encontrados_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE
);
