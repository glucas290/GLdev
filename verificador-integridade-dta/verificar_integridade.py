"""
Verifica a integridade de cargas em trânsito aduaneiro (DTA) agendadas para o
armazém e avisa a equipe por e-mail.

Fluxo:
1. Busca no banco Oracle os agendamentos do dia com DTA (placa + número).
2. Ignora as DTAs que já foram comunicadas (histórico em CSV).
3. Consulta cada DTA no site público da Receita Federal.
4. Se houver registro de INTEGRIDADE, monta uma tabela HTML e envia pelo Outlook.
5. Registra as DTAs comunicadas no histórico.
"""
import os
import time
from pathlib import Path

import oracledb
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

URL_CONSULTA_DTA = "https://www4c.receita.fazenda.gov.br/g36154/jsp/consFluxoDecl.jsp"
XPATH_CAMPO_DTA = '//*[@id="casca"]/table/tbody/tr[2]/td[3]/form/input[4]'
XPATH_BOTAO_CONSULTAR = '//*[@id="casca"]/table/tbody/tr[2]/td[3]/form/input[5]'
XPATH_RESULTADO = "/html/body/fieldset[13]/table/tbody/tr[1]/td"
XPATH_DATA = "/html/body/fieldset[13]/table/tbody/tr[2]/td"
XPATH_RESPONSAVEL = "/html/body/fieldset[13]/table/tbody/tr[3]/td"
XPATH_RECINTO_ORIGEM = "/html/body/fieldset[2]/table/tbody/tr[2]/td"

COLUNAS_RESULTADO = ["DTA", "PLACA", "RESULTADO", "DATA", "NOME", "RECINTO ORIGEM"]

ESTILO_TABELA = """
<style>
    table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px; }
    th { background-color: #4CAF50; color: white; padding: 8px; text-align: left; }
    td { border: 1px solid #ddd; padding: 8px; vertical-align: top; max-width: 300px; word-wrap: break-word; }
    tr:nth-child(even) { background-color: #f2f2f2; }
</style>
"""


def buscar_agendamentos(usuario: str, senha: str, dsn: str, caminho_sql: Path) -> pd.DataFrame:
    """Executa a consulta SQL e devolve as colunas PLACA e DOCUMENTO."""
    query = caminho_sql.read_text(encoding="utf-8")
    with oracledb.connect(user=usuario, password=senha, dsn=dsn) as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(query)
            colunas = [c[0].upper() for c in cursor.description]
            df = pd.DataFrame(cursor.fetchall(), columns=colunas)
    return df.drop_duplicates()


def filtrar_ja_enviados(df: pd.DataFrame, historico: Path) -> pd.DataFrame:
    if not historico.exists():
        return df.copy()
    enviados = set(pd.read_csv(historico, dtype=str)["DTA"])
    return df[~df["DOCUMENTO"].astype(str).isin(enviados)].copy()


def iniciar_navegador() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def _texto(navegador, xpath: str) -> str:
    return navegador.find_element(By.XPATH, xpath).text.strip()


def consultar_dta(navegador, dta: str) -> dict | None:
    """Consulta uma DTA. Retorna os dados se houver registro de integridade."""
    navegador.get(URL_CONSULTA_DTA)
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, XPATH_CAMPO_DTA))
    ).send_keys(dta)
    navegador.find_element(By.XPATH, XPATH_BOTAO_CONSULTAR).click()

    # O site abre um alert quando a DTA não existe ou não pode ser consultada.
    try:
        navegador.switch_to.alert.accept()
        return None
    except NoAlertPresentException:
        pass

    resultado = WebDriverWait(navegador, 5).until(
        EC.presence_of_element_located((By.XPATH, XPATH_RESULTADO))
    ).text.strip()
    if "INTEGRIDADE" not in resultado.upper():
        return None

    return {
        "RESULTADO": resultado,
        "DATA": _texto(navegador, XPATH_DATA),
        "NOME": _texto(navegador, XPATH_RESPONSAVEL),
        "RECINTO ORIGEM": _texto(navegador, XPATH_RECINTO_ORIGEM),
    }


def verificar_dtas(df: pd.DataFrame) -> pd.DataFrame:
    navegador = iniciar_navegador()
    encontrados = []
    try:
        for _, linha in df.iterrows():
            dta = str(linha["DOCUMENTO"]).strip()
            placa = str(linha["PLACA"]).strip()
            if not dta:
                continue
            try:
                dados = consultar_dta(navegador, dta)
            except Exception as e:
                print(f"Não foi possível consultar a DTA {dta}: {e}")
                continue
            if dados:
                encontrados.append({"DTA": dta, "PLACA": placa, **dados})
            time.sleep(1)
    finally:
        navegador.quit()
    return pd.DataFrame(encontrados, columns=COLUNAS_RESULTADO)


def enviar_email(df: pd.DataFrame, destinatarios: str, assinatura: str) -> None:
    import win32com.client as win32

    tabela = df.to_html(index=False, border=0, justify="center")
    corpo = f"""
    <p>Prezados,</p>
    <p>Segue abaixo a atualização de integridade extraída automaticamente:</p>
    {ESTILO_TABELA}
    {tabela}
    <p>Atenciosamente,<br><b>{assinatura}</b></p>
    """
    email = win32.Dispatch("Outlook.Application").CreateItem(0)
    email.To = destinatarios
    email.Subject = "Atualização de Integridade - DTA"
    email.HTMLBody = corpo
    email.Send()
    print("E-mail enviado.")


def atualizar_historico(df: pd.DataFrame, historico: Path) -> None:
    if historico.exists():
        df = pd.concat([pd.read_csv(historico, dtype=str), df], ignore_index=True)
    historico.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(historico, index=False)
    print(f"Histórico atualizado em {historico}")


def main() -> None:
    load_dotenv()
    historico = Path(os.getenv("HISTORICO_CSV", "saida/emails_enviados.csv"))

    df_agendamentos = buscar_agendamentos(
        usuario=os.environ["ORACLE_USUARIO"],
        senha=os.environ["ORACLE_SENHA"],
        dsn=os.environ["ORACLE_DSN"],
        caminho_sql=Path(os.getenv("ARQUIVO_SQL", "consulta_agendamentos.sql")),
    )
    if df_agendamentos.empty:
        print("Nenhum agendamento com DTA encontrado.")
        return

    pendentes = filtrar_ja_enviados(df_agendamentos, historico)
    if pendentes.empty:
        print("Nenhuma DTA nova para processar.")
        return
    print(f"{len(pendentes)} DTAs novas para consultar.")

    resultado = verificar_dtas(pendentes)
    if resultado.empty:
        print("Nenhum registro de integridade encontrado.")
        return

    enviar_email(
        resultado,
        destinatarios=os.environ["EMAIL_DESTINATARIOS"],
        assinatura=os.getenv("EMAIL_ASSINATURA", "Automação de Integridade"),
    )
    atualizar_historico(resultado, historico)


if __name__ == "__main__":
    main()
