-- Locais físicos de cada unidade (sala, almoxarifado, andar...), onde os itens ficam.
CREATE TABLE IF NOT EXISTS locais (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    unidade_id INT NOT NULL,
    nome VARCHAR(150) NOT NULL,
    descricao VARCHAR(255) NULL,
    ativo TINYINT(1) NOT NULL DEFAULT 1,
    UNIQUE KEY uq_locais_unidade_nome (unidade_id, nome),
    CONSTRAINT fk_locais_unidade FOREIGN KEY (unidade_id) REFERENCES unidades (id)
);
