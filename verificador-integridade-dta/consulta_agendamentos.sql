-- Consulta de EXEMPLO.
-- A versão original usava o modelo de dados interno do sistema operacional
-- do terminal, que não é público. Adapte para o seu banco.
--
-- Contrato esperado pelo script: uma linha por veículo/documento com as colunas
--   PLACA            placa do veículo
--   DOCUMENTO        número da DTA
--   JANELA_AGENDADA  janela de atendimento
--   DATA_PRE_GATE    quando o veículo passou pelo pré-gate
--   DATA_GATE_IN     quando entrou no pátio (nulo se ainda não entrou)
SELECT
    v.placa              AS PLACA,
    d.numero             AS DOCUMENTO,
    j.inicio_janela      AS JANELA_AGENDADA,
    pg.data_ocorrencia   AS DATA_PRE_GATE,
    gi.data_entrada      AS DATA_GATE_IN
FROM agendamentos a
JOIN veiculos      v  ON v.id = a.veiculo_id
JOIN documentos    d  ON d.id = a.documento_id
JOIN janelas       j  ON j.id = a.janela_id
JOIN pre_gate      pg ON pg.agendamento_id = a.id
LEFT JOIN gate_in  gi ON gi.agendamento_id = a.id
WHERE d.tipo = 'DTA-E'
  AND a.cancelado IS NULL
  AND j.inicio_janela >= TRUNC(SYSDATE) - 1
ORDER BY j.inicio_janela DESC
