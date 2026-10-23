# Consulta de Perdimento – Siscarga / Siscomex

Robô que mantém atualizados os indicadores de cargas em **risco de perdimento** num armazém alfandegado. Uma carga que passa do prazo de armazenagem sem ser desembaraçada pode ser declarada abandonada e perdida para a União. Por isso, a equipe precisa saber, para cada carga, em que pé está o despacho.

## Problema

Centenas de cargas armazenadas, cada uma com prazos diferentes. Consultar manualmente o CE Mercante e a DI de cada carga no Siscarga e no Siscomex levava horas e era feito de forma irregular.

## Solução

```
SharePoint (cargas + regras)  ──►  seleciona as cargas do dia
                                          │
                       Siscarga: CE Mercante (situação, bloqueios, nº da DI)
                       Siscarga: entrega da DI (desembaraço, VMLE, pendências)
                                          │
                       grava os indicadores na lista do SharePoint
                                          │
                       Siscomex Web: UF de destino das DIs desembaraçadas
                                          │
                       e-mail de log com sucessos e falhas
```

- **Frequência inteligente**: uma lista de regras define, por faixa de dias em estoque, em que dias da semana a carga é consultada. Cargas perto do prazo são consultadas com mais frequência.
- **Detecção de pendências**: identifica nas mensagens da DI pendências de ICMS, AFRMM, frete do armador e TUM.
- **Resiliência**: se a sessão do Siscarga ou o token do SharePoint expirarem, o robô reautentica e segue para o próximo item, sem perder a execução.
- **Log**: ao final, envia por e-mail um resumo com sucessos e falhas de cada etapa.

## Estrutura

| Arquivo | Responsabilidade |
|---|---|
| `main.py` | Orquestra o fluxo e as regras de negócio |
| `siscomex.py` | Robô Selenium para Siscarga e Siscomex Web |
| `sharepoint_client.py` | Leitura e edição de listas do SharePoint |
| `notificacao.py` | E-mails via Outlook com tabelas formatadas |

## Stack

Python · Selenium · pandas · Office365-REST-Python-Client · pywin32 (Outlook) · python-dotenv

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env
python main.py
```

Requisitos: Windows com Outlook, Google Chrome e um **certificado digital** (e-CPF/e-CNPJ) com acesso ao Siscarga instalado no navegador.

## Observações

- A versão original tinha um caso em que um aviso inesperado na consulta da DI derrubava o processamento do item. Agora ele é tratado e registrado.
- Os nomes das colunas (`SCMX_*`) são o modelo da lista de exemplo; ajuste-os à sua lista.
