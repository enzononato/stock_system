-- ATENÇÃO: reverter apaga os dados de compra (valor, data, garantia e chave da NF) de todos os itens.
SET @ddl = (SELECT IF(COUNT(*) = 1, 'ALTER TABLE items DROP COLUMN valor_compra', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'valor_compra');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 1, 'ALTER TABLE items DROP COLUMN data_compra', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'data_compra');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 1, 'ALTER TABLE items DROP COLUMN garantia_ate', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'garantia_ate');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @ddl = (SELECT IF(COUNT(*) = 1, 'ALTER TABLE items DROP COLUMN nf_anexo_key', 'SELECT 1') FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'items' AND column_name = 'nf_anexo_key');
PREPARE stmt FROM @ddl;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
