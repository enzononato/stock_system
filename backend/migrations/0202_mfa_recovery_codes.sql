-- F2: códigos de recuperação do 2FA (só o hash é guardado; cada código serve uma vez).
CREATE TABLE IF NOT EXISTS mfa_recovery_codes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    usuario_id INT NOT NULL,
    code_hash CHAR(64) NOT NULL,
    usado_em DATETIME NULL,
    UNIQUE KEY uq_mfa_recovery (usuario_id, code_hash),
    CONSTRAINT fk_mfa_recovery_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
);
