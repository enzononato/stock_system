-- Reverter recusa (falha sem alterar nada) enquanto existir item 'Em Manutenção': encerre ou cancele as manutenções antes.
-- Trava: se houver item 'Em Manutenção', a subconsulta devolve 2 linhas e o MySQL aborta (erro 1242).
SELECT (SELECT x FROM (SELECT 1 AS x UNION ALL SELECT 2) d WHERE EXISTS (SELECT 1 FROM items WHERE status = 'Em Manutenção')) AS trava_em_manutencao;
ALTER TABLE items MODIFY status ENUM('Disponível','Indisponível','Pendente','Pendente Devolução') DEFAULT 'Disponível';
