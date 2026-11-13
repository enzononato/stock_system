-- Registro de apagamento de dados (LGPD) na baixa de equipamentos com armazenamento.
CREATE TABLE IF NOT EXISTS baixas_apagamento (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    item_id INT NOT NULL,
    history_id INT NOT NULL,
    dados_apagados ENUM('sim','nao_aplica') NOT NULL,
    metodo ENUM('formatacao_segura','reset_fabrica','destruicao_fisica','nao_aplica') NOT NULL,
    apagado_por VARCHAR(100) NULL,
    apagado_em DATE NULL,
    comprovante_key VARCHAR(255) NULL,
    registrado_por VARCHAR(100) NOT NULL,
    registrado_em DATETIME NOT NULL,
    UNIQUE KEY uq_baixas_apagamento_history (history_id),
    KEY idx_baixas_apagamento_item (item_id),
    CONSTRAINT fk_baixas_apagamento_item FOREIGN KEY (item_id) REFERENCES items (id),
    CONSTRAINT fk_baixas_apagamento_history FOREIGN KEY (history_id) REFERENCES history (id)
) CHARACTER SET utf8mb4;
