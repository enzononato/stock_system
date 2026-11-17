-- F2 (rodada 1): falhas da 2ª etapa do login (limite por token e por usuário) e bloqueio temporário.
-- Datas em UTC. Fica no banco para valer entre os vários workers.
CREATE TABLE IF NOT EXISTS mfa_falhas (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    usuario_id INT NOT NULL,
    jti VARCHAR(64) NOT NULL,
    criado_em DATETIME NOT NULL,
    INDEX idx_mfa_falhas_usuario (usuario_id, criado_em),
    INDEX idx_mfa_falhas_jti (jti),
    CONSTRAINT fk_mfa_falhas_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
);
SET @faltando := (SELECT COUNT(*) = 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'mfa_bloqueado_ate');
SET @sql := IF(@faltando, 'ALTER TABLE usuarios ADD COLUMN mfa_bloqueado_ate DATETIME NULL', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
