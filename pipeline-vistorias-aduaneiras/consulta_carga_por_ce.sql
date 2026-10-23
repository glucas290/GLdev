-- Consulta de EXEMPLO.
-- A versão original usava o modelo de dados interno do sistema operacional
-- do terminal, que não é público. Adapte para o seu banco.
--
-- Recebe o CE Mercante pela bind variable :ce (nunca concatenar o valor no SQL)
-- e deve retornar uma linha por contêiner/lote com as colunas:
--   MERCID, NRO_BL, NAVIO, VIAGEM, CONTEINER, CLASSE, TAMANHO, CLIENTE
SELECT DISTINCT
    m.id                   AS MERCID,
    c.numero_bl            AS NRO_BL,
    n.nome                 AS NAVIO,
    v.numero_viagem        AS VIAGEM,
    m.identificacao        AS CONTEINER,
    CASE WHEN m.eh_conteiner = 1 THEN 'CONTEINER' ELSE 'LOTE' END AS CLASSE,
    m.codigo_iso           AS TAMANHO,
    cl.nome                AS CLIENTE
FROM conhecimentos c
JOIN mercadorias m  ON m.conhecimento_id = c.id
JOIN viagens     v  ON v.id = c.viagem_id
JOIN navios      n  ON n.id = v.navio_id
JOIN clientes    cl ON cl.id = c.cliente_id
WHERE c.ce_mercante = :ce
