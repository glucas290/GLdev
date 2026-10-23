-- Consulta de EXEMPLO.
-- A versão original usava o modelo de dados interno do sistema operacional
-- do terminal, que não é público. Adapte para o seu banco.
--
-- Contrato esperado pelo script: retornar as colunas PLACA e DOCUMENTO
-- (número da DTA) dos agendamentos ativos a partir de hoje.
SELECT
    v.placa   AS PLACA,
    d.numero  AS DOCUMENTO
FROM agendamentos a
JOIN veiculos   v ON v.id = a.veiculo_id
JOIN documentos d ON d.id = a.documento_id
WHERE d.tipo = 'DTA-E'
  AND a.cancelado IS NULL
  AND a.inicio_janela >= TRUNC(SYSDATE)
ORDER BY a.inicio_janela DESC
