"""Envio de e-mail pelo Outlook (Windows) com tabelas pandas no corpo."""
from datetime import datetime, timedelta

import pandas as pd

ESTILO_TABELA = """
<style>
    .dataframe thead { background: #ED6B10; color: white; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    tr, td, th { text-align: center; padding: 3px; border: 1px solid black; }
    th { padding: 5px; white-space: nowrap; }
</style>
"""


class EmailOutlook:
    def __init__(self):
        import win32com.client as win32

        self.outlook = win32.Dispatch("Outlook.Application")
        self.mail = self.outlook.CreateItem(0)
        self.mail.HTMLBody = ""

    def escrever(self, texto: str) -> None:
        self.mail.HTMLBody += f"<br>{texto}" if self.mail.HTMLBody else texto

    def anexar(self, caminho_arquivo: str) -> None:
        self.mail.Attachments.Add(caminho_arquivo)

    def adicionar_tabela(self, df: pd.DataFrame) -> None:
        tabela = df.to_html(index=False, justify="left")
        self.mail.HTMLBody += f"<html><head>{ESTILO_TABELA}</head><body>{tabela}</body></html>"

    def enviar(self, destinatarios: str, assunto: str, importancia: int = 1, lembrete: bool = False, sinalizar: bool = False) -> None:
        self.mail.To = destinatarios
        self.mail.Subject = assunto
        self.mail.Importance = importancia
        self.mail.ReminderSet = lembrete
        if lembrete:
            self.mail.ReminderTime = datetime.now() + timedelta(days=1)
        if sinalizar:
            self.mail.FlagRequest = "Acompanhar"
        self.mail.Send()


def enviar_log_execucao(destinatarios: str, nome_rotina: str, status: str, detalhe: str) -> None:
    """Envia um resumo da execução (sucesso ou falha) para os responsáveis."""
    email = EmailOutlook()
    email.adicionar_tabela(pd.DataFrame({
        "Rotina": [nome_rotina],
        "Hora da Execução": [datetime.now().strftime("%d-%m-%Y %H:%M:%S")],
        "Status": [status],
        "Execução/Erro": [detalhe],
    }))
    email.enviar(destinatarios, f"Log de automação - {nome_rotina} - {status}")
