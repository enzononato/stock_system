-- Transferências de equipamento entre unidades. As unidades são gravadas por nome
-- (retrato do momento, como history.revenda).
CREATE TABLE IF NOT EXISTS transferencias (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    item_id INT NOT NULL,
    de_unidade VARCHAR(100) NULL,
    para_unidade VARCHAR(100) NOT NULL,
    de_local_id INT NULL,
    para_local_id INT NULL,
    motivo VARCHAR(255) NOT NULL,
    transferido_por VARCHAR(100) NOT NULL,
    transferido_em DATETIME NOT NULL,
    KEY idx_transferencias_item (item_id),
    CONSTRAINT fk_transferencias_item FOREIGN KEY (item_id) REFERENCES items (id),
    CONSTRAINT fk_transferencias_de_local FOREIGN KEY (de_local_id) REFERENCES locais (id) ON DELETE SET NULL,
    CONSTRAINT fk_transferencias_para_local FOREIGN KEY (para_local_id) REFERENCES locais (id) ON DELETE SET NULL
);
