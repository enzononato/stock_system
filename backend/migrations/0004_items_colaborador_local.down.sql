SET @existe = (SELECT COUNT(*) FROM information_schema.table_constraints WHERE table_schema = DATABASE() AND table_name = 'items' AND constraint_name = 'fk_items_local' AND constraint_type = 'FOREIGN KEY');
SET @comando = IF(@existe = 1, 'ALTER TABLE items DROP FOREIGN KEY fk_items_local', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.table_constraints WHERE table_schema = DATABASE() AND table_name = 'items' AND constraint_name = 'fk_items_colaborador' AND constraint_type = 'FOREIGN KEY');
SET @comando = IF(@existe = 1, 'ALTER TABLE items DROP FOREIGN KEY fk_items_colaborador', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'local_id');
SET @comando = IF(@existe = 1, 'ALTER TABLE items DROP COLUMN local_id', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;

SET @existe = (SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'colaborador_id');
SET @comando = IF(@existe = 1, 'ALTER TABLE items DROP COLUMN colaborador_id', 'SELECT 1');
PREPARE passo FROM @comando;
EXECUTE passo;
DEALLOCATE PREPARE passo;
