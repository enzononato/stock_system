-- Linhas telefônicas (Fase 5): uma linha por número, com aparelho (Celular/Tablet) e colaborador opcionais.
-- Datas em hora local do servidor. item_id e colaborador_id viram NULL se o registro de origem for apagado.
CREATE TABLE IF NOT EXISTS linhas (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    numero VARCHAR(11) NOT NULL,
    operadora VARCHAR(50) NOT NULL,
    plano VARCHAR(150) NULL,
    custo_mensal DECIMAL(10,2) NOT NULL,
    status ENUM('ativa','suspensa','cancelada') NOT NULL DEFAULT 'ativa',
    unidade_id INT NOT NULL,
    item_id INT NULL,
    colaborador_id INT NULL,
    iccid VARCHAR(22) NULL,
    observacoes VARCHAR(500) NULL,
    criada_em DATETIME NOT NULL,
    atualizada_em DATETIME NOT NULL,
    UNIQUE KEY uq_linhas_numero (numero),
    KEY idx_linhas_status_unidade (status, unidade_id),
    CONSTRAINT fk_linhas_unidade FOREIGN KEY (unidade_id) REFERENCES unidades (id),
    CONSTRAINT fk_linhas_item FOREIGN KEY (item_id) REFERENCES items (id) ON DELETE SET NULL,
    CONSTRAINT fk_linhas_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores (id) ON DELETE SET NULL
);

-- Histórico da linha. item_id (aparelho depois do evento) e item_id_anterior não têm FK: são registro histórico.
CREATE TABLE IF NOT EXISTS linha_eventos (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    linha_id INT NOT NULL,
    tipo ENUM('ativacao','troca_aparelho','troca_usuario','portabilidade','suspensao','reativacao','cancelamento','edicao') NOT NULL,
    campo VARCHAR(30) NULL,
    de_valor VARCHAR(500) NULL,
    para_valor VARCHAR(500) NULL,
    item_id INT NULL,
    item_id_anterior INT NULL,
    ocorrido_em DATETIME NOT NULL,
    por_usuario VARCHAR(100) NULL,
    observacao VARCHAR(255) NULL,
    KEY idx_linha_eventos_linha (linha_id, ocorrido_em),
    KEY idx_linha_eventos_item (item_id),
    KEY idx_linha_eventos_item_anterior (item_id_anterior),
    CONSTRAINT fk_linha_eventos_linha FOREIGN KEY (linha_id) REFERENCES linhas (id) ON DELETE CASCADE
);
