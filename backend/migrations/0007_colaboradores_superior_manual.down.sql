SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'colaboradores' AND column_name = 'superior_manual');
SET @comando = IF(@existe = 1, 'ALTER TABLE colaboradores DROP COLUMN superior_manual', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
