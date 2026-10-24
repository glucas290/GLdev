# KPI Diário – Armazém (Power BI)

Painel usado pela equipe operacional do armazém alfandegado de um terminal portuário para acompanhar, todos os dias, os três indicadores principais da operação. O painel original consolidava consultas ao banco Oracle do sistema operacional do terminal (TOS) e planilhas de controle mantidas pelas equipes no SharePoint.

> Esta versão foi **reconstruída para o portfólio** com bases fictícias: não contém dados de clientes, endereços internos, nomes da empresa ou de colegas. A lógica de negócio (etapas, metas em dias úteis, tratamentos no Power Query) segue a do painel original.

## Indicadores

| KPI | O que mede | Início → Fim | Meta |
|---|---|---|---|
| **Conferência Aduaneira** | Tempo para atender a conferência/vistoria fiscal | Pedido de posicionamento do contêiner → envio do relatório ao fiscal (RFB, MAPA ou ANVISA) | D+2 a D+3 úteis, conforme o serviço |
| **Entrega de Carga Solta** | Tempo do caminhão dentro do terminal | Gate-in → gate-out | 2h (normal) · 3h (especial e DTA) |
| **Recebimento e Desunitização** | Visão ponta a ponta da desova | Solicitação de posicionamento do cliente → armazenagem da carga e envio do relatório | Por etapa (posicionamento, desova, presença de carga, relatório) |

As metas ficam na planilha `Parametros e Metas.xlsx`, junto com o calendário de feriados. A gestão pode alterar um prazo sem mexer no Power BI.

## Arquitetura dos dados

```
dados/
├── oracle_tos/      extrações do TOS (no original: Oracle.Database com SQL nativo)
│   ├── tos_posicionamento_conferencia.csv
│   ├── tos_agendamento_entrega.csv
│   ├── tos_movimentacao_conteiner.csv
│   ├── tos_clientes.csv
│   └── tos_transportadoras.csv
└── sharepoint/      planilhas de controle das equipes (no original: SharePoint.Contents)
    ├── Controle de Conferencia Aduaneira.xlsx
    ├── Controle do Recebimento - Desova.xlsx
    ├── Solicitacao de Posicionamento.xlsx
    ├── Justificativas.xlsx
    └── Parametros e Metas.xlsx
```

As planilhas reproduzem os problemas reais das bases manuais: linhas de título antes do cabeçalho, contêiner digitado com espaços ou em minúsculas, nome de cliente fora do padrão e linhas duplicadas. O Power Query trata tudo isso.

## Power Query

As consultas estão organizadas em camadas:

| Grupo | Conteúdo |
|---|---|
| `0 Parâmetros` | `CaminhoDados`: pasta das bases |
| `1 Funções` | `fnPrazoDiaUtil` (prazo em dias úteis com hora de corte e feriados), `fnAplicarMeta`, `fnAbaControle`, `fnLimparContainer`, `fnStatusPrazo`, entre outras |
| `2 Origem Oracle (TOS)` | Staging das extrações do TOS (não carregadas) |
| `3 Origem SharePoint` | Staging das planilhas, já limpas e sem duplicidade (não carregadas) |
| `4 Dimensões` | `dCalendario`, `dCliente`, `dTransportadora`, `dMeta` |
| `5 Fatos` | `fConferencia`, `fEntregaCargaSolta`, `fRecebimento`, `fSolicitacaoPosicionamento` |

Destaques:
- **Metas em dias úteis calculadas no Power Query.** Exemplo: um pedido feito até as 19h deve ser posicionado até as 07h do próximo dia útil, desconsiderando sábados, domingos e feriados. No painel original essa regra era uma coluna DAX repetida em várias tabelas; aqui ela virou uma única função M.
- **Cruzamento de origens.** O recebimento parte da planilha da equipe, busca a solicitação de posicionamento na planilha do pátio e o horário real do posicionamento no TOS (se o posicionamento foi refeito, vale o primeiro).
- **Justificativas.** São mescladas pela chave usada pela equipe (janela agendada + placa). Os horários vindos do Excel são arredondados ao minuto para a mesclagem não falhar.

Modelo estrela: as quatro fatos se ligam a `dCalendario` (pela data de referência) e a `dCliente`; a entrega também se liga a `dTransportadora`.

## Como abrir

1. Instale o [Power BI Desktop](https://aka.ms/pbidesktopstore).
2. Abra `powerbi/KPI Armazem Portuario.pbip`.
3. Em **Transformar dados → Editar parâmetros**, aponte `CaminhoDados` para a pasta `dados\` do seu clone (com a barra no final).
4. Clique em **Atualizar**.

O projeto está no formato **Power BI Project (.pbip)**: o modelo fica em arquivos TMDL de texto, versionáveis no Git.

## Regenerar as bases fictícias

```bash
cd scripts
npm install
npm run gerar
```

O gerador (`scripts/gerar_dados.js`) usa semente fixa, então sempre produz as mesmas bases: cerca de 4,3 mil conferências, 16,7 mil agendamentos de entrega e 3,1 mil desovas entre jan/2025 e out/2026.

## Stack

Power BI Desktop · Power Query (M) · DAX · TMDL / PBIP · Node.js (gerador de dados)
