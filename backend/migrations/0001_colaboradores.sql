-- Cadastro de colaboradores: espelho da base de RH (somente leitura) com complemento manual.
-- Datas gravadas em UTC.
CREATE TABLE IF NOT EXISTS colaboradores (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    cpf CHAR(11) NOT NULL,
    nome VARCHAR(200) NOT NULL,
    matricula VARCHAR(50) NULL,
    cargo VARCHAR(150) NULL,
    setor VARCHAR(150) NULL,
    unidade_rh VARCHAR(150) NULL,
    unidade_id INT NULL,
    situacao ENUM('ativo', 'afastado', 'desligado') NOT NULL DEFAULT 'ativo',
    email VARCHAR(200) NULL,
    telefone VARCHAR(20) NULL,
    email_manual TINYINT(1) NOT NULL DEFAULT 0,
    telefone_manual TINYINT(1) NOT NULL DEFAULT 0,
    superior_id INT NULL,
    gestor_nome_rh VARCHAR(200) NULL,
    origem ENUM('rh', 'manual') NOT NULL DEFAULT 'rh',
    sincronizado_em DATETIME NULL,
    criado_em DATETIME NOT NULL,
    atualizado_em DATETIME NOT NULL,
    UNIQUE KEY uq_colaboradores_cpf (cpf),
    INDEX idx_colaboradores_nome (nome),
    INDEX idx_colaboradores_situacao (situacao),
    INDEX idx_colaboradores_unidade (unidade_id),
    CONSTRAINT fk_colaboradores_unidade FOREIGN KEY (unidade_id) REFERENCES unidades (id),
    CONSTRAINT fk_colaboradores_superior FOREIGN KEY (superior_id) REFERENCES colaboradores (id) ON DELETE SET NULL
);

-- Histórico das sincronizações com a base de RH (a última aparece na tela de colaboradores).
CREATE TABLE IF NOT EXISTS colaboradores_sync (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    iniciado_em DATETIME NOT NULL,
    concluido_em DATETIME NULL,
    status ENUM('ok', 'erro', 'em_andamento') NOT NULL DEFAULT 'em_andamento',
    total_base INT NOT NULL DEFAULT 0,
    criados INT NOT NULL DEFAULT 0,
    atualizados INT NOT NULL DEFAULT 0,
    desligados INT NOT NULL DEFAULT 0,
    ignorados INT NOT NULL DEFAULT 0,
    duplicados INT NOT NULL DEFAULT 0,
    erro VARCHAR(500) NULL
);
