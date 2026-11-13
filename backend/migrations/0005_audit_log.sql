-- Trilha de auditoria: quem fez o quê, quando, de onde, com o antes/depois.
-- ocorrido_em é sempre gravado em UTC.
CREATE TABLE IF NOT EXISTS audit_log (
    id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    ocorrido_em DATETIME(3) NOT NULL,
    usuario_id INT NULL,
    username VARCHAR(100) NULL,
    papel VARCHAR(50) NULL,
    ip VARCHAR(45) NULL,
    entidade VARCHAR(50) NOT NULL,
    entidade_id VARCHAR(64) NULL,
    acao VARCHAR(50) NOT NULL,
    antes JSON NULL,
    depois JSON NULL,
    INDEX idx_audit_entidade (entidade, entidade_id),
    INDEX idx_audit_ocorrido_em (ocorrido_em)
);
