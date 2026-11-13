-- F2: links de redefinição de senha (só o hash do token é guardado; datas em UTC).
CREATE TABLE IF NOT EXISTS password_reset_tokens (
    id INT AUTO_INCREMENT PRIMARY KEY,
    usuario_id INT NOT NULL,
    token_hash CHAR(64) NOT NULL,
    criado_em DATETIME NOT NULL,
    expira_em DATETIME NOT NULL,
    usado_em DATETIME NULL,
    criado_ip VARCHAR(45) NULL,
    UNIQUE KEY uq_password_reset_token_hash (token_hash),
    INDEX idx_password_reset_usuario (usuario_id, criado_em),
    CONSTRAINT fk_password_reset_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
) CHARACTER SET utf8mb4;
