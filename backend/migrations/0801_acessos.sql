-- Módulo de acessos (F8): catálogo de sistemas, pacotes por cargo, ficha de acessos do colaborador,
-- eventos (admissão, desligamento, mudança de função) e a lista de tarefas deles.
-- Nenhuma coluna guarda senha: o módulo só registra o login.
-- Sem COLLATE próprio: a comparação sem caixa de nome/login vem do padrão do banco e, quando precisa
-- valer em qualquer banco, é feita na consulta (LOWER) ou no código.
CREATE TABLE IF NOT EXISTS sistemas (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    descricao VARCHAR(255) NULL,
    ativo TINYINT(1) NOT NULL DEFAULT 1,
    UNIQUE KEY uq_sistemas_nome (nome)
);

CREATE TABLE IF NOT EXISTS pacotes_funcao (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    cargo VARCHAR(150) NOT NULL,
    sistema_id INT NOT NULL,
    UNIQUE KEY uq_pacotes_funcao (cargo, sistema_id),
    INDEX idx_pacotes_funcao_sistema (sistema_id),
    CONSTRAINT fk_pacotes_funcao_sistema FOREIGN KEY (sistema_id) REFERENCES sistemas (id)
);

CREATE TABLE IF NOT EXISTS colaborador_acessos (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    colaborador_id INT NOT NULL,
    sistema_id INT NOT NULL,
    login VARCHAR(150) NOT NULL,
    status ENUM('ativo', 'desativado') NOT NULL DEFAULT 'ativo',
    item_id INT NULL,
    criado_por VARCHAR(100) NOT NULL,
    criado_em DATETIME NOT NULL,
    desativado_por VARCHAR(100) NULL,
    desativado_em DATETIME NULL,
    UNIQUE KEY uq_colaborador_acessos (colaborador_id, sistema_id, login),
    INDEX idx_colaborador_acessos_item (item_id),
    INDEX idx_colaborador_acessos_sistema (sistema_id),
    CONSTRAINT fk_colaborador_acessos_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores (id),
    CONSTRAINT fk_colaborador_acessos_sistema FOREIGN KEY (sistema_id) REFERENCES sistemas (id),
    CONSTRAINT fk_colaborador_acessos_item FOREIGN KEY (item_id) REFERENCES items (id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS acesso_eventos (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    colaborador_id INT NOT NULL,
    tipo ENUM('admissao', 'desligamento', 'mudanca_funcao') NOT NULL,
    cargo_anterior VARCHAR(150) NULL,
    cargo_novo VARCHAR(150) NULL,
    status ENUM('aberto', 'concluido', 'cancelado') NOT NULL DEFAULT 'aberto',
    observacao VARCHAR(500) NULL,
    aberto_por VARCHAR(100) NOT NULL,
    aberto_em DATETIME NOT NULL,
    concluido_por VARCHAR(100) NULL,
    concluido_em DATETIME NULL,
    cancelado_por VARCHAR(100) NULL,
    cancelado_em DATETIME NULL,
    motivo_cancelamento VARCHAR(500) NULL,
    INDEX idx_acesso_eventos_colaborador (colaborador_id, status),
    INDEX idx_acesso_eventos_status (status, aberto_em),
    CONSTRAINT fk_acesso_eventos_colaborador FOREIGN KEY (colaborador_id) REFERENCES colaboradores (id)
);

CREATE TABLE IF NOT EXISTS acesso_tarefas (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    evento_id INT NOT NULL,
    sistema_id INT NOT NULL,
    acesso_id INT NULL,
    acao ENUM('criar', 'desativar') NOT NULL,
    status ENUM('pendente', 'feito', 'nao_se_aplica') NOT NULL DEFAULT 'pendente',
    login VARCHAR(150) NULL,
    feito_por VARCHAR(100) NULL,
    feito_em DATETIME NULL,
    observacao VARCHAR(500) NULL,
    INDEX idx_acesso_tarefas_evento (evento_id),
    INDEX idx_acesso_tarefas_sistema (sistema_id),
    INDEX idx_acesso_tarefas_acesso (acesso_id),
    CONSTRAINT fk_acesso_tarefas_evento FOREIGN KEY (evento_id) REFERENCES acesso_eventos (id) ON DELETE CASCADE,
    CONSTRAINT fk_acesso_tarefas_sistema FOREIGN KEY (sistema_id) REFERENCES sistemas (id),
    CONSTRAINT fk_acesso_tarefas_acesso FOREIGN KEY (acesso_id) REFERENCES colaborador_acessos (id) ON DELETE SET NULL
);

INSERT IGNORE INTO sistemas (nome, descricao, ativo) VALUES
    ('AD', 'Active Directory (usuário de rede)', 1),
    ('E-mail', 'E-mail corporativo', 1),
    ('VPN', 'Acesso remoto (VPN)', 1),
    ('Promax', 'ERP Promax', 1),
    ('Biz Deliveries', 'Biz Deliveries (entregas)', 1),
    ('WMS', 'WMS (armazém)', 1);
