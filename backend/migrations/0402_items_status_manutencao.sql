-- Novo status de equipamento: Em Manutenção.
ALTER TABLE items MODIFY status ENUM('Disponível','Indisponível','Pendente','Pendente Devolução','Em Manutenção') DEFAULT 'Disponível';
