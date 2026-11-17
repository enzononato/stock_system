-- Login por e-mail: confirmação do e-mail corporativo no próximo login.
-- Todas as contas atuais começam com email_confirmado = 0. O código de 6 dígitos
-- é guardado só como HMAC (codigo_hash); datas em UTC, passadas pelo Python.
SET @faltando := (SELECT COUNT(*) = 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'email_confirmado');
SET @sql := IF(@faltando, 'ALTER TABLE usuarios ADD COLUMN email_confirmado TINYINT(1) NOT NULL DEFAULT 0', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
-- Mesma guarda (@faltando): só na primeira aplicação. Encerra as sessões abertas antes
-- do deploy das contas sem e-mail confirmado (todas, nesse momento), para ninguém
-- continuar renovando o refresh token sem passar pela etapa do e-mail.
SET @sql := IF(@faltando, 'UPDATE usuarios SET sessao_versao = sessao_versao + 1 WHERE email_confirmado = 0', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
CREATE TABLE IF NOT EXISTS email_confirmacoes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    usuario_id INT NOT NULL,
    email VARCHAR(200) NOT NULL,
    codigo_hash CHAR(64) NOT NULL,
    expira_em DATETIME NOT NULL,
    tentativas INT NOT NULL DEFAULT 0,
    usado TINYINT(1) NOT NULL DEFAULT 0,
    criado_em DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_email_confirmacoes_usuario (usuario_id),
    CONSTRAINT fk_email_confirmacoes_usuario FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
);
-- Contas criadas por convite nascem com senha_definida = 0 (senha inutilizável) até a
-- pessoa criar a própria senha pelo link. As contas atuais já têm senha: ficam com 1.
SET @faltando := (SELECT COUNT(*) = 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'senha_definida');
SET @sql := IF(@faltando, 'ALTER TABLE usuarios ADD COLUMN senha_definida TINYINT(1) NOT NULL DEFAULT 1', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
