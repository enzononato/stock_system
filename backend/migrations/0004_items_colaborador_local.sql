-- Vínculo do item com o colaborador responsável e com o local físico.
-- Colunas anuláveis: itens existentes continuam como estão. Apagar o colaborador ou o
-- local apenas solta o vínculo (ON DELETE SET NULL).
-- ADD COLUMN IF NOT EXISTS só existe no MariaDB; o teste em information_schema vale nos dois.
SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'colaborador_id');
SET @comando = IF(@existe = 0, 'ALTER TABLE items ADD COLUMN colaborador_id INT NULL', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'local_id');
SET @comando = IF(@existe = 0, 'ALTER TABLE items ADD COLUMN local_id INT NULL', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.table_constraints WHERE table_schema = DATABASE() AND table_name = 'items' AND constraint_name = 'fk_items_colaborador' AND constraint_type = 'FOREIGN KEY');
SET @comando = IF(@existe = 0, 'ALTER TABLE items ADD CONSTRAINT fk_items_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores (id) ON DELETE SET NULL', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.table_constraints WHERE table_schema = DATABASE() AND table_name = 'items' AND constraint_name = 'fk_items_local' AND constraint_type = 'FOREIGN KEY');
SET @comando = IF(@existe = 0, 'ALTER TABLE items ADD CONSTRAINT fk_items_local FOREIGN KEY (local_id) REFERENCES locais (id) ON DELETE SET NULL', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
