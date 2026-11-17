-- Ordens de manutenção de equipamentos (envio à assistência, retorno com custo e anexo).
CREATE TABLE IF NOT EXISTS manutencoes (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    item_id INT NOT NULL,
    status ENUM('aberta','enviada','em_reparo','retornada','sem_conserto','cancelada') NOT NULL DEFAULT 'aberta',
    defeito VARCHAR(500) NOT NULL,
    assistencia VARCHAR(150) NULL,
    aberta_por VARCHAR(100) NOT NULL,
    aberta_em DATETIME NOT NULL,
    enviada_em DATETIME NULL,
    em_reparo_em DATETIME NULL,
    previsao_retorno DATE NULL,
    retornada_em DATETIME NULL,
    cancelada_em DATETIME NULL,
    encerrada_por VARCHAR(100) NULL,
    custo DECIMAL(12,2) NULL,
    anexo_key VARCHAR(255) NULL,
    observacoes VARCHAR(1000) NULL,
    KEY idx_manutencoes_item (item_id),
    KEY idx_manutencoes_status (status),
    KEY idx_manutencoes_aberta_em (aberta_em),
    CONSTRAINT fk_manutencoes_item FOREIGN KEY (item_id) REFERENCES items (id)
);
