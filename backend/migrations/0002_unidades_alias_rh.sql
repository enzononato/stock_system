-- Nome da unidade como aparece na base de RH, para ligar colaboradores às unidades.
-- ADD COLUMN IF NOT EXISTS só existe no MariaDB; o teste em information_schema vale nos dois.
SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'unidades' AND column_name = 'alias_rh');
SET @comando = IF(@existe = 0, 'ALTER TABLE unidades ADD COLUMN alias_rh VARCHAR(150) NULL', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
