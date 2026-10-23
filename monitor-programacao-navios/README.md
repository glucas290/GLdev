# Monitor de Programação de Navios

Robô que coleta a programação de navios publicada no site de um terminal portuário e a disponibiliza para a equipe de planejamento: em CSV e, opcionalmente, numa lista do SharePoint que alimenta painéis internos.

## Problema

A equipe de planejamento do armazém precisava acompanhar a chegada, a atracação e a saída de navios de um terminal vizinho. A consulta era manual, no site, várias vezes ao dia, e os dados eram copiados para planilhas.

## Solução

1. Abre a página de programação com Selenium (Chrome em modo headless).
2. Aplica o filtro de status (por exemplo, navios já operados) e espera a tabela recarregar.
3. Lê a tabela com `pandas.read_html` e descarta colunas que não interessam.
4. Salva um CSV (`;` como separador, pronto para Excel/Power BI).
5. Com `--sharepoint`: limpa a lista de destino e regrava os dados, mapeando cada coluna para o nome interno da lista.

## Stack

Python · Selenium · pandas · Office365-REST-Python-Client · python-dotenv

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env   # preencha só se for usar --sharepoint
python programacao_navios.py
python programacao_navios.py --sharepoint
```

Opções: `--status`, `--saida`, `--url` e `--com-janela` (mostra o navegador; útil para depurar).

## Observações

- A autenticação no SharePoint é por usuário e senha (legado). Em tenants com MFA, use um app registrado no Azure AD.
- A primeira versão usava `time.sleep(5)` para esperar a tabela. Esta versão espera a tabela antiga sair do DOM, o que é mais rápido e mais confiável.
