DROP TABLE IF EXISTS mfa_falhas;
SET @existe := (SELECT COUNT(*) > 0 FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name = 'mfa_bloqueado_ate');
SET @sql := IF(@existe, 'ALTER TABLE usuarios DROP COLUMN mfa_bloqueado_ate', 'DO 0');
PREPARE passo FROM @sql;
EXECUTE passo;
DEALLOCATE PREPARE passo;
