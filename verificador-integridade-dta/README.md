# Monitor de DTA – Gate-in e Integridade

Robô que acompanha as cargas em **trânsito aduaneiro (DTA)** que chegam ao armazém. Ele faz duas coisas que antes eram manuais:

1. **Registro de chegada:** quando o caminhão passa pelo gate-in, o robô registra a chegada do veículo no Siscomex Trânsito, confere se o registro aparece no fluxo da declaração e avisa a equipe sobre o status de cada veículo.
2. **Integridade:** consulta cada DTA no site da Receita e alerta a equipe quando há ocorrência de integridade (violação de lacre ou de volumes) registrada no trânsito.

## Problema

Os registros de chegada eram feitos à mão, veículo a veículo, e atrasos geravam pendências no despacho. As ocorrências de integridade só eram descobertas quando alguém consultava a DTA, muitas vezes com o caminhão já no pátio.

## Solução

```
Oracle (agendamentos DTA)
   ├─ veículos com gate-in ─► registra chegada no Siscomex ─► confere no fluxo ─► e-mail com status
   └─ DTAs ainda não verificadas ─► consulta integridade ─► e-mail se houver ocorrência
Históricos em CSV evitam reprocessar o mesmo documento.
```

- **Validação em duas etapas:** se a tela de registro não confirma na hora, o robô verifica no fluxo da declaração se a chegada consta antes de marcá-la como falha.
- **E-mail de alerta:** o assunto muda conforme o resultado, e a tabela destaca em cores cada status (confirmada, não encontrada, pendente).
- **Log controlado:** o robô roda a cada poucos minutos, mas só manda o e-mail de log no início de cada hora, quando processa algo ou quando há falha.

## Estrutura

| Arquivo | Responsabilidade |
|---|---|
| `verificar_integridade.py` | Orquestração, banco, históricos e e-mails |
| `robo_transito.py` | Robô Selenium para o Siscomex Trânsito |
| `consulta_agendamentos.sql` | Consulta de exemplo (contrato de colunas) |

## Stack

Python · oracledb · Selenium · pandas · pywin32 (Outlook) · python-dotenv

## Como executar

```bash
pip install -r requirements.txt
cp .env.example .env
python verificar_integridade.py
```

Requer Windows com Outlook, Google Chrome e um **certificado digital** com acesso ao Siscomex Trânsito instalado no navegador.

## Observações

- O SQL deste repositório é um **exemplo**: a consulta original usava o modelo de dados interno do terminal.
- A versão original tinha dois bugs: vírgulas no fim da linha transformavam usuário e senha em tuplas, e um `replace("OK", ...)` no HTML inteiro podia alterar placas que contêm "OK". Os dois foram corrigidos.
