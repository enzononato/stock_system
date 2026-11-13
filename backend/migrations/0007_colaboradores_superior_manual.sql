-- 1 quando o superior foi definido (ou limpo) à mão; a sincronização não mexe nesses.
-- ADD COLUMN IF NOT EXISTS só existe no MariaDB; o teste em information_schema vale nos dois.
SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'colaboradores' AND column_name = 'superior_manual');
SET @comando = IF(@existe = 0, 'ALTER TABLE colaboradores ADD COLUMN superior_manual TINYINT(1) NOT NULL DEFAULT 0', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
