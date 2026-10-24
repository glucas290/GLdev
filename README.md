# GLdev – Portfólio de Automações

Projetos de automação de processos (RPA e integrações) que desenvolvi para a operação de um armazém alfandegado em terminal portuário. Eles automatizam consultas a sistemas da Receita Federal, cruzam dados de banco e SharePoint e notificam as equipes por e-mail.

> Os códigos foram **reescritos e anonimizados** para este portfólio: não contêm credenciais, endereços internos, dados de clientes nem nomes da empresa ou de colegas. Configurações sensíveis ficam em variáveis de ambiente (`.env`, ver `.env.example` em cada projeto) e os dados de exemplo são fictícios.

## Projetos

| Projeto | O que faz | Stack |
|---|---|---|
| [Monitor de Programação de Navios](monitor-programacao-navios/) | Coleta a programação de navios de um terminal e publica em CSV ou SharePoint | Selenium, pandas, SharePoint |
| [Monitor de DTA – Gate-in e Integridade](verificador-integridade-dta/) | Registra e valida no Siscomex a chegada dos veículos em trânsito aduaneiro e alerta a equipe sobre ocorrências de integridade | Oracle, Selenium, Outlook |
| [Pipeline de Vistorias Aduaneiras](pipeline-vistorias-aduaneiras/) | Sincroniza solicitações de vistoria por API, calcula o SLA em dias úteis, enriquece com Siscomex e banco e sinaliza itens de risco | REST/OAuth2, Selenium, Oracle, SharePoint |
| [Consulta de Perdimento – Siscarga/Siscomex](consulta-perdimento-siscomex/) | Atualiza os indicadores de cargas em risco de perdimento consultando CE Mercante e DI | Selenium, SharePoint, Outlook |
| [Consulta de DUIMP – Portal Único](consulta-duimp-portal-unico/) | Confere DUIMPs na API oficial do Portal Único Siscomex com certificado digital | API REST, mTLS, pandas |
| [KPI Diário – Armazém](kpi-armazem-portuario/) | Painel Power BI com os KPIs de conferência aduaneira, entrega de carga solta e recebimento/desunitização, consolidando TOS (Oracle) e planilhas SharePoint · [ver painel ao vivo](https://app.powerbi.com/view?r=eyJrIjoiN2M0MDBjMjEtNjU4NS00NTZhLThkM2YtYjMyM2UzMmQ2ZjljIiwidCI6ImU2OWJmYmEwLTQwZjQtNDMwMi1hZjc4LTJjNTkxYTU3ZmFmYiJ9) | Power BI, Power Query (M), DAX, PBIP |

## Tecnologias

Python 3.10+ · Selenium · pandas · REST APIs / OAuth2 · Oracle (oracledb) · SharePoint Online (Office365-REST-Python-Client) · Outlook via COM (pywin32) · workalendar

## Como rodar um projeto

```bash
cd nome-do-projeto
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env   # preencha com as suas credenciais
python <script>.py
```

Cada pasta tem um README com o problema, a solução e os requisitos específicos.
