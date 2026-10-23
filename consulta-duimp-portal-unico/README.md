# Consulta de DUIMP – API do Portal Único Siscomex

Integração com a API oficial do **Portal Único Siscomex** para conferir as DUIMPs (Declaração Única de Importação) de um armazém alfandegado: compara o que está no sistema interno com o que a Receita registra (valor da mercadoria, situação, canal de parametrização e importador).

## Problema

Divergências entre o sistema operacional do armazém e a Receita (cliente diferente, valor diferente, DUIMP parametrizada em canal vermelho) só eram percebidas no faturamento ou na liberação da carga.

## Solução

1. Recebe a lista de DUIMPs da demanda (no projeto original, vinda do banco interno; aqui, de um CSV).
2. Autentica **uma única vez** na API com o certificado digital A1 (mTLS via `requests-pkcs12`) e reaproveita o token, renovando-o se expirar.
3. Consulta o endpoint de estoque pós-ACD de cada DUIMP.
4. Gera um painel (HTML + CSV) com os dados internos e os do Portal Único lado a lado, mais o status de cada consulta.

O módulo `sharepoint_client.py` reúne as operações de SharePoint usadas nas automações de conferência: ler listas grandes com paginação, editar itens, anexar arquivos, fazer upload e ler planilhas Excel do site.

## Stack

Python · requests / requests-pkcs12 · API REST Portal Único · pandas · Office365-REST-Python-Client

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env   # caminho e senha do certificado .pfx
python consulta_duimp.py --entrada exemplo_duimps.csv
```

O arquivo `exemplo_duimps.csv` tem números e empresas **fictícios**. É preciso um certificado digital com perfil habilitado no Portal Único.

## Melhorias em relação à versão original

- A versão original autenticava na API a cada DUIMP consultada. Agora autentica uma vez e renova o token só quando necessário.
- O certificado e a senha saíram do código e foram para variáveis de ambiente.
