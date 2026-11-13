SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'unidades' AND column_name = 'alias_rh');
SET @comando = IF(@existe = 1, 'ALTER TABLE unidades DROP COLUMN alias_rh', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
