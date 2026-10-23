"""
Etapa 1 – Sincroniza as solicitações de vistoria com a lista do SharePoint.

Busca na API do sistema de agendamento as solicitações do período, classifica
cada uma (tipo de declaração, equipe, serviço), calcula o prazo de atendimento
em dias úteis e grava na lista: insere as novas e marca as canceladas.

Uso:
    python etapa1_sincronizar_agendamentos.py
    python etapa1_sincronizar_agendamentos.py --inicio 01/10/2025 --fim 31/10/2025
"""
import argparse
import os
from datetime import datetime, timedelta

import pandas as pd
import requests
from dotenv import load_dotenv

from regras import (CalendarioRJ, calcular_prazo, definir_servico, definir_setor,
                    formatar_declaracao, identificar_tipo_documento)
from sharepoint_client import SharePoint

# Coluna do DataFrame -> nome interno do campo na lista do SharePoint
CAMPOS_SHAREPOINT = {
    "ID": "Title",
    "Declaracao": "Declara_x00e7__x00e3_o",
    "Janela": "Janela",
    "DataAgendamento": "DatadoAgendamento",
    "Observacao": "Observacao",
    "DataCriacao": "Cria_x00e7__x00e3_o",
    "Conteiner": "Ctnr_Lote",
    "Usuario_Criacao": "Solicitante",
    "TipoDocumento": "TipoDocumento",
    "Setor": "Setor",
    "Servico": "Servico",
    "CanceladoSistema": "CanceladoSistema",
    "DataIndicador": "DataIndicador",
}
COLUNAS_API = ["ID", "Declaracao", "Janela", "DataAgendamento", "Observacao",
               "DataCriacao", "Conteiner", "Usuario_Criacao", "StatusDescricao"]


def obter_token() -> str:
    """OAuth2 (password grant) com autenticação básica do cliente."""
    resposta = requests.post(
        os.environ["AGENDAMENTO_TOKEN_URL"],
        auth=(os.environ["AGENDAMENTO_CLIENT_ID"], os.environ["AGENDAMENTO_CLIENT_SECRET"]),
        data={
            "grant_type": "password",
            "username": os.environ["AGENDAMENTO_USUARIO"],
            "password": os.environ["AGENDAMENTO_SENHA"],
            "scope": "offline_access",
        },
        timeout=30,
    )
    resposta.raise_for_status()
    return resposta.json()["access_token"]


def buscar_solicitacoes(token: str, inicio: str, fim: str) -> pd.DataFrame:
    resposta = requests.get(
        os.environ["AGENDAMENTO_API_URL"],
        headers={"Authorization": f"Bearer {token}"},
        params={"dataInicio": inicio, "dataTermino": fim},
        timeout=60,
    )
    resposta.raise_for_status()
    return pd.DataFrame(resposta.json().get("value", []))


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica as regras de negócio e devolve as colunas da lista."""
    df = df[COLUNAS_API].copy()
    df["ID"] = df["ID"].astype(str)
    df["TipoDocumento"] = df["Declaracao"].apply(identificar_tipo_documento)
    df["Declaracao"] = df["Declaracao"].apply(formatar_declaracao)
    df["DataAgendamento"] = pd.to_datetime(df["DataAgendamento"], errors="coerce")
    df["DataCriacao"] = pd.to_datetime(df["DataCriacao"], errors="coerce")
    df["Observacao"] = (df["Observacao"].fillna("").astype(str)
                        .str.replace(r"[\r\n]+", " ", regex=True).str.strip().str[:255])
    df["Setor"] = df.apply(lambda r: definir_setor(r["TipoDocumento"], r["Observacao"]), axis=1)
    df["Servico"] = df.apply(lambda r: definir_servico(r["Setor"], r["Observacao"]), axis=1)
    df["CanceladoSistema"] = df["StatusDescricao"].apply(
        lambda s: "Cancelado" if str(s).strip().upper() == "CANCELADO" else "")

    calendario = CalendarioRJ()
    df["DataIndicador"] = df.apply(
        lambda r: calcular_prazo(r["Servico"], r["DataCriacao"], calendario), axis=1)
    return df


def sincronizar(sp: SharePoint, lista: str, df: pd.DataFrame) -> tuple[int, int]:
    """Insere solicitações novas e marca como canceladas as que mudaram de status."""
    existentes = sp.ler_lista(lista, todos=True)
    ids_por_titulo, cancelados = {}, set()
    if not existentes.empty:
        titulos = existentes["Title"].astype(str)
        ids_por_titulo = dict(zip(titulos, existentes["Id"]))
        if "CanceladoSistema" in existentes:
            cancelados = set(titulos[existentes["CanceladoSistema"] == "Cancelado"])

    inseridos = 0
    for _, linha in df[~df["ID"].isin(ids_por_titulo)].iterrows():
        item = {}
        for coluna, campo in CAMPOS_SHAREPOINT.items():
            valor = linha[coluna]
            if isinstance(valor, (pd.Timestamp, datetime)):
                valor = None if pd.isna(valor) else valor.strftime("%Y-%m-%dT%H:%M:%S")
            elif isinstance(valor, str) and not valor.strip():
                valor = None
            item[campo] = valor
        if sp.adicionar_item(lista, item):
            inseridos += 1

    atualizados = 0
    a_cancelar = df[(df["CanceladoSistema"] == "Cancelado")
                    & df["ID"].isin(ids_por_titulo) & ~df["ID"].isin(cancelados)]
    for id_solicitacao in a_cancelar["ID"]:
        if sp.editar_item(lista, ids_por_titulo[id_solicitacao], {"CanceladoSistema": "Cancelado"}):
            atualizados += 1

    return inseridos, atualizados


def main() -> None:
    load_dotenv()
    hoje = datetime.now()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--inicio", default=(hoje - timedelta(days=7)).strftime("%d/%m/%Y"))
    parser.add_argument("--fim", default=(hoje + timedelta(days=30)).strftime("%d/%m/%Y"))
    args = parser.parse_args()

    solicitacoes = buscar_solicitacoes(obter_token(), args.inicio, args.fim)
    if solicitacoes.empty:
        print("Nenhuma solicitação retornada pela API.")
        return
    df = preparar(solicitacoes)

    sp = SharePoint(os.environ["SHAREPOINT_SITE_URL"], os.environ["SHAREPOINT_USUARIO"],
                    os.environ["SHAREPOINT_SENHA"])
    inseridos, cancelados = sincronizar(sp, os.environ["SHAREPOINT_LISTA_VISTORIAS"], df)
    print(f"{inseridos} solicitações inseridas, {cancelados} marcadas como canceladas.")


if __name__ == "__main__":
    main()
