-- Dados de compra do equipamento. Colunas anuláveis: itens existentes continuam como estão.
-- Idempotente coluna a coluna (o MySQL 8.0 não tem ADD COLUMN IF NOT EXISTS).
SET @ddl = (SELECT IF(COUNT(*) = 0, 'ALTER TABLE items ADD COLUMN valor_compra DECIMAL(12,2) NULL', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'valor_compra');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 0, 'ALTER TABLE items ADD COLUMN data_compra DATE NULL', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'data_compra');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 0, 'ALTER TABLE items ADD COLUMN garantia_ate DATE NULL', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'garantia_ate');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 0, 'ALTER TABLE items ADD COLUMN nf_anexo_key VARCHAR(255) NULL', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'nf_anexo_key');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
