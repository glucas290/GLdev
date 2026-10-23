"""
Monitora as cargas em trânsito aduaneiro (DTA) agendadas para o armazém.

Fluxo:
1. Busca no Oracle os agendamentos de DTA (placa, documento, pré-gate e gate-in).
2. Gate-in: para cada veículo que já entrou no pátio, registra a chegada no
   Siscomex Trânsito, confere o registro no fluxo da declaração e envia um
   e-mail com o status de cada veículo.
3. Integridade: consulta cada DTA no site da Receita e avisa a equipe quando
   houver ocorrência de integridade (violação de lacre ou de volumes).
4. Guarda históricos em CSV para não processar o mesmo documento duas vezes.
"""
import os
from datetime import datetime
from pathlib import Path

import oracledb
import pandas as pd
from dotenv import load_dotenv

from robo_transito import RoboTransito

NOME_ROTINA = "Monitor de DTA"
COLUNAS_INTEGRIDADE = ["DTA", "PLACA", "RESULTADO", "DATA", "NOME", "RECINTO ORIGEM"]

STATUS_CHEGADA = {
    "OK": '<b style="color: #28a745;">CONFIRMADA</b>',
    "NAO_LOCALIZADA": '<b style="color: #dc3545;">NÃO ENCONTRADA (VERIFICAR)</b>',
    "NAO_EXECUTADO": '<b style="color: #6c757d;">PENDENTE</b>',
}

ESTILO_TABELA = """
<style>
    table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; }
    th { background-color: #343a40; color: white; padding: 10px; text-align: left; }
    td { border: 1px solid #ddd; padding: 8px; vertical-align: top; max-width: 300px; word-wrap: break-word; }
    tr:nth-child(even) { background-color: #f9f9f9; }
</style>
"""


# ---------------------------------------------------------------- dados

def buscar_agendamentos(usuario: str, senha: str, dsn: str, caminho_sql: Path) -> pd.DataFrame:
    """Executa a consulta SQL. Colunas esperadas: PLACA, DOCUMENTO, JANELA_AGENDADA,
    DATA_PRE_GATE e DATA_GATE_IN."""
    query = caminho_sql.read_text(encoding="utf-8")
    with oracledb.connect(user=usuario, password=senha, dsn=dsn) as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(query)
            colunas = [c[0].upper() for c in cursor.description]
            df = pd.DataFrame(cursor.fetchall(), columns=colunas)
    return df.drop_duplicates()


def documentos_no_historico(caminho: Path, coluna: str) -> set[str]:
    if not caminho.exists():
        return set()
    return set(pd.read_csv(caminho, dtype=str)[coluna])


def adicionar_ao_historico(df: pd.DataFrame, caminho: Path, coluna: str) -> None:
    if caminho.exists():
        df = pd.concat([pd.read_csv(caminho, dtype=str), df.astype(str)], ignore_index=True)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    df.drop_duplicates(subset=[coluna]).to_csv(caminho, index=False)


# ---------------------------------------------------------------- e-mail

def enviar_email(destinatarios: str, assunto: str, corpo_html: str) -> None:
    import win32com.client as win32

    email = win32.Dispatch("Outlook.Application").CreateItem(0)
    email.To = destinatarios
    email.Subject = assunto
    email.HTMLBody = corpo_html
    email.Send()
    print(f"E-mail enviado: {assunto}")


def montar_email_gate_in(df: pd.DataFrame, status_por_placa: dict, assinatura: str) -> tuple[str, str]:
    tabela = df[["PLACA", "DOCUMENTO", "JANELA_AGENDADA", "DATA_PRE_GATE", "DATA_GATE_IN"]].copy()
    tabela.columns = ["PLACA", "DOCUMENTO", "JANELA AGENDADA", "DATA PRÉ-GATE", "DATA GATE IN"]
    codigos = df["PLACA"].map(status_por_placa).fillna("NAO_EXECUTADO")
    tabela["STATUS CHEGADA"] = codigos.map(STATUS_CHEGADA)

    falhou = (codigos != "OK").any()
    if falhou:
        assunto = "ATENÇÃO: Falha na chegada automática (DTA-E) - ação manual necessária"
        aviso = ('<p style="color: #dc3545; font-weight: bold;">Um ou mais veículos não tiveram a chegada '
                 "confirmada automaticamente. Verifique o Siscomex manualmente.</p>")
    else:
        assunto = "Sucesso: chegada de veículos DTA-E registrada e validada"
        aviso = ""

    corpo = f"""
    <html><body style="font-family: Arial, sans-serif;">
        <p>Prezados,</p>
        <p>Segue o status do processamento automático de chegada dos veículos DTA-E:</p>
        {aviso}
        {ESTILO_TABELA}
        {tabela.to_html(index=False, border=0, escape=False, justify="center")}
        <p style="font-size: 11px; color: #666;">* E-mail gerado automaticamente pelo robô de monitoramento de trânsito.</p>
        <p>Atenciosamente,<br><b>{assinatura}</b></p>
    </body></html>
    """
    return assunto, corpo


def montar_email_integridade(df: pd.DataFrame, assinatura: str) -> str:
    return f"""
    <p>Prezados,</p>
    <p>Segue abaixo a atualização de integridade extraída automaticamente:</p>
    {ESTILO_TABELA}
    {df.to_html(index=False, border=0, justify="center")}
    <p>Atenciosamente,<br><b>{assinatura}</b></p>
    """


# ---------------------------------------------------------------- etapas

def processar_gate_in(robo: RoboTransito, df: pd.DataFrame) -> dict:
    """Registra e valida a chegada de cada veículo. Retorna {placa: status}."""
    status = {}
    for _, linha in df.drop_duplicates(subset=["PLACA"]).iterrows():
        documento, placa = str(linha["DOCUMENTO"]).strip(), str(linha["PLACA"]).strip()
        print(f"Registrando chegada: placa {placa} / documento {documento}")
        if robo.registrar_chegada(documento, placa) or robo.chegada_consta_no_fluxo(documento):
            status[placa] = "OK"
        else:
            status[placa] = "NAO_LOCALIZADA"
    return status


def processar_integridade(robo: RoboTransito, df: pd.DataFrame) -> pd.DataFrame:
    encontrados = []
    for _, linha in df.iterrows():
        dta = str(linha["DOCUMENTO"]).strip()
        if not dta:
            continue
        dados = robo.consultar_integridade(dta)
        if dados:
            encontrados.append({"DTA": dta, "PLACA": str(linha["PLACA"]).strip(), **dados})
    return pd.DataFrame(encontrados, columns=COLUNAS_INTEGRIDADE)


def executar() -> str:
    destinatarios = os.environ["EMAIL_DESTINATARIOS"]
    assinatura = os.getenv("EMAIL_ASSINATURA", NOME_ROTINA)
    historico_gate_in = Path(os.getenv("HISTORICO_GATE_IN", "saida/gate_in_enviados.csv"))
    historico_integridade = Path(os.getenv("HISTORICO_INTEGRIDADE", "saida/emails_enviados.csv"))

    agendamentos = buscar_agendamentos(
        usuario=os.environ["ORACLE_USUARIO"],
        senha=os.environ["ORACLE_SENHA"],
        dsn=os.environ["ORACLE_DSN"],
        caminho_sql=Path(os.getenv("ARQUIVO_SQL", "consulta_agendamentos.sql")),
    )
    if agendamentos.empty:
        return "Nenhum agendamento de DTA encontrado."

    agendamentos["DOCUMENTO"] = agendamentos["DOCUMENTO"].astype(str)
    resumo = []
    robo = RoboTransito()
    try:
        robo.logar_com_certificado()

        ja_avisados = documentos_no_historico(historico_gate_in, "DOCUMENTO")
        gate_in = agendamentos[
            agendamentos["DATA_GATE_IN"].notna() & ~agendamentos["DOCUMENTO"].isin(ja_avisados)
        ]
        if not gate_in.empty:
            print(f"Gate-in: {len(gate_in)} documentos para registrar.")
            status = processar_gate_in(robo, gate_in)
            assunto, corpo = montar_email_gate_in(gate_in, status, assinatura)
            enviar_email(destinatarios, assunto, corpo)
            adicionar_ao_historico(gate_in[["DOCUMENTO"]], historico_gate_in, "DOCUMENTO")
            resumo.append(f"Gate-in: {len(gate_in)} documentos")

        ja_comunicados = documentos_no_historico(historico_integridade, "DTA")
        pendentes = agendamentos[~agendamentos["DOCUMENTO"].isin(ja_comunicados)]
        if not pendentes.empty:
            print(f"Integridade: {len(pendentes)} DTAs para consultar.")
            ocorrencias = processar_integridade(robo, pendentes)
            if not ocorrencias.empty:
                enviar_email(destinatarios, "Atualização de Integridade - DTA",
                             montar_email_integridade(ocorrencias, assinatura))
                adicionar_ao_historico(ocorrencias, historico_integridade, "DTA")
                resumo.append(f"Integridade: {len(ocorrencias)} ocorrências")
    finally:
        robo.fechar()

    return "Processamento concluído. " + " | ".join(resumo) if resumo else "Finalizado sem envio de e-mail."


def main() -> None:
    load_dotenv()
    try:
        detalhe, status = executar(), "Sucesso"
    except Exception as e:
        detalhe, status = str(e), "Falha"
    print(f"{status}: {detalhe}")

    # O robô roda a cada poucos minutos; o log por e-mail sai só no início de
    # cada hora, quando algo foi processado ou quando houve falha.
    destinatarios_log = os.getenv("EMAIL_DESTINATARIOS_LOG")
    if destinatarios_log and (datetime.now().minute <= 5 or "concluído" in detalhe or status == "Falha"):
        tabela = pd.DataFrame({
            "Rotina": [NOME_ROTINA],
            "Hora da Execução": [datetime.now().strftime("%d-%m-%Y %H:%M:%S")],
            "Status": [status],
            "Execução/Erro": [detalhe],
        })
        enviar_email(destinatarios_log, f"Log de automação - {NOME_ROTINA} - {status}",
                     f"<p>Relatório de execução do robô:</p>{tabela.to_html(index=False, border=1)}")


if __name__ == "__main__":
    main()
