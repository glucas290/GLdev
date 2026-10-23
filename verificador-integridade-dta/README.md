# Verificador de Integridade de DTA

Automação que cruza os agendamentos de caminhões do armazém com a consulta pública da Receita Federal e avisa a equipe, por e-mail, quando uma carga em trânsito aduaneiro (DTA) tem registro de **integridade** (ocorrência de violação de lacre ou de volumes).

## Problema

Cargas que chegam por DTA podem ter ocorrências registradas no trânsito. A equipe só descobria isso ao consultar cada DTA manualmente no site da Receita, normalmente depois de o caminhão já estar no pátio.

## Solução

1. **Banco Oracle** – busca os agendamentos do dia com DTA (placa + número) usando o SQL do arquivo [`consulta_agendamentos.sql`](consulta_agendamentos.sql).
2. **Deduplicação** – ignora as DTAs já comunicadas, guardadas num histórico CSV.
3. **Selenium** – consulta cada DTA no site da Receita e trata alerts de DTA inexistente.
4. **Filtro** – mantém só as consultas com registro de integridade e extrai data, responsável e recinto de origem.
5. **Outlook** – envia uma tabela HTML formatada para a equipe e atualiza o histórico.

## Stack

Python · oracledb · Selenium · pandas · pywin32 (Outlook) · python-dotenv

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env   # credenciais do banco e destinatários
python verificar_integridade.py
```

Requer Windows com o Outlook configurado (envio via COM) e Google Chrome.

## Observações

- O SQL deste repositório é um **exemplo**: a consulta original usava o modelo de dados interno do terminal. O script só exige que a consulta retorne as colunas `PLACA` e `DOCUMENTO`.
- Credenciais, endereço do banco e destinatários ficam no `.env`, fora do código.
