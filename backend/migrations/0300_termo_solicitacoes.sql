-- Solicitações de assinatura de termos (entrega, devolução e uso), com o retorno do RevSign.
-- Datas em horário de Brasília, sem fuso. Apagar o item apaga as solicitações; apagar o
-- colaborador ou a linha do histórico apenas solta o vínculo.
CREATE TABLE IF NOT EXISTS termo_solicitacoes (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    item_id INT NOT NULL,
    colaborador_id INT NULL,
    tipo ENUM('entrega','devolucao','uso') NOT NULL,
    history_id INT NULL,
    modelo_versao VARCHAR(100) NULL,
    status ENUM('agendado','enviado','assinado','recusado','expirado','cancelado','erro','manual') NOT NULL,
    destinatario_nome VARCHAR(200) NULL,
    destinatario_email VARCHAR(200) NULL,
    solicitado_por VARCHAR(100) NOT NULL,
    solicitado_em DATETIME NOT NULL,
    agendado_para DATETIME NULL,
    enviado_em DATETIME NULL,
    assinado_em DATETIME NULL,
    cancelado_em DATETIME NULL,
    revsign_documento_id VARCHAR(64) NULL,
    documento_key VARCHAR(255) NULL,
    arquivo_assinado_key VARCHAR(255) NULL,
    motivo VARCHAR(500) NULL,
    atualizado_em DATETIME NOT NULL,
    UNIQUE KEY uq_termo_revsign_doc (revsign_documento_id),
    INDEX idx_termo_item (item_id, id),
    INDEX idx_termo_colaborador (colaborador_id),
    INDEX idx_termo_status (status, agendado_para),
    CONSTRAINT fk_termo_item FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE,
    CONSTRAINT fk_termo_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores(id) ON DELETE SET NULL,
    CONSTRAINT fk_termo_history FOREIGN KEY (history_id) REFERENCES history(id) ON DELETE SET NULL
) CHARACTER SET utf8mb4;
