-- Reverte a 0900 sem apagar a tabela usuarios nem o e-mail já gravado.
DROP TABLE IF EXISTS email_confirmacoes;
SET @existe := (SELECT COUNT(*) > 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'email_confirmado');
SET @sql := IF(@existe, 'ALTER TABLE usuarios DROP COLUMN email_confirmado', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
SET @existe := (SELECT COUNT(*) > 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'senha_definida');
SET @sql := IF(@existe, 'ALTER TABLE usuarios DROP COLUMN senha_definida', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
