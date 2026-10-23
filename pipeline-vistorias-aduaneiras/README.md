# Pipeline de Vistorias Aduaneiras

Pipeline em três etapas que organiza as **solicitações de vistoria** feitas pela fiscalização aduaneira a um armazém alfandegado. Ele junta tudo numa lista do SharePoint que a equipe operacional usa para planejar o trabalho e acompanhar o SLA.

## Problema

As solicitações de vistoria chegavam por um sistema de agendamento, com texto livre. A equipe precisava classificar cada uma à mão (tipo de declaração, equipe da fiscalização, tipo de serviço) e calcular o prazo de atendimento. Depois, consultava o Siscomex e o sistema interno para descobrir quais contêineres separar e o que havia dentro de cada carga.

## Solução

```
Etapa 1  API de agendamento ──► classificação + prazo (SLA) ──► lista SharePoint
Etapa 2  DI pendentes ──► Siscomex (CE, volumes, peso) + Oracle (BL, navio, contêineres) ──► lista
Etapa 3  DI detalhadas ──► Siscomex (adições e itens) ──► termos de risco ──► Excel anexado + flags
```

### Etapa 1 – Sincronizar agendamentos
- Autentica na API por OAuth2 e busca as solicitações do período.
- **Classifica** pelo texto: tipo de declaração (DI, DUIMP, DSI, DUE), equipe da fiscalização e serviço (desova total, parcial, por adição, abertura, posicionamento). As regras ficam em [`regras.py`](regras.py).
- **Calcula o prazo de atendimento** em dias úteis, considerando feriados nacionais e do Rio de Janeiro (`workalendar`). Cada serviço tem seu número de dias e horário limite.
- Insere as solicitações novas e marca as que foram canceladas no sistema de origem.

### Etapa 2 – Detalhar a DI
- Consulta no Siscomex Importação o CE Mercante, os volumes e o peso bruto.
- Com o CE, busca no Oracle a carga: BL, navio, viagem, contêineres e cliente. A consulta usa *bind variable*; a versão original concatenava o valor no SQL.
- Se a DI tem vários contêineres, cria uma linha por contêiner na lista.
- Converte cada valor para o tipo do campo no SharePoint (texto, número, data).

### Etapa 3 – Adições e itens de risco
- Lê todas as adições da DI e os itens de mercadoria de cada adição.
- Compara as descrições com uma lista de **termos de risco** mantida pela equipe no SharePoint.
- Gera uma planilha Excel por declaração, anexa aos itens da lista e preenche a quantidade de referências e a marcação de item de risco.

## Estrutura

| Arquivo | Responsabilidade |
|---|---|
| `etapa1_sincronizar_agendamentos.py` | API → classificação → SharePoint |
| `etapa2_detalhar_di.py` | Siscomex + Oracle → SharePoint |
| `etapa3_adicoes_e_riscos.py` | Adições, itens de risco e anexos |
| `regras.py` | Regras de classificação e de prazo (funções puras) |
| `siscomex_importacao.py` | Robô Selenium para o Siscomex Importação |
| `sharepoint_client.py` | Listas, anexos e conversão de tipos no SharePoint |
| `consulta_carga_por_ce.sql` | Consulta de exemplo (contrato de colunas) |

## Stack

Python · REST/OAuth2 · Selenium · oracledb · pandas · workalendar · openpyxl · Office365-REST-Python-Client

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env
python etapa1_sincronizar_agendamentos.py
python etapa2_detalhar_di.py
python etapa3_adicoes_e_riscos.py
```

As etapas 2 e 3 exigem um **certificado digital** com acesso ao Siscomex. A seleção do certificado acontece numa janela do sistema operacional, então o robô pausa até o operador confirmar no terminal. Entre uma declaração e outra há uma pausa configurável, porque o Siscomex limita a frequência de consultas.

## Observações

- O SQL deste repositório é um **exemplo**: a consulta original usava o modelo de dados interno do terminal.
- O calendário de feriados da versão original fixava Corpus Christi em 20/06, mas esse feriado é móvel. Agora ele é calculado a partir da Páscoa.
