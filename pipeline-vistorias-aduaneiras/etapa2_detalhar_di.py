"""
Etapa 2 – Detalha as vistorias de DI com dados do Siscomex e do banco.

Para cada solicitação de DI ainda não processada, consulta no Siscomex o CE
Mercante, os volumes e o peso bruto. Com o CE, busca no banco Oracle os dados
da carga (BL, navio, viagem, contêineres, cliente). Se a DI tiver vários
contêineres, a primeira linha é atualizada e as demais são criadas na lista.
"""
import os
import time
from pathlib import Path

import oracledb
import pandas as pd
from dotenv import load_dotenv

from siscomex_importacao import SiscomexImportacao
from sharepoint_client import SharePoint, converter_valor

# Campo lógico -> nome interno na lista
CAMPOS_BASE = {
    "Titulo": "Title",
    "Declaracao": "Declara_x00e7__x00e3_o",
    "Janela": "Janela",
    "DataAgendamento": "DatadoAgendamento",
    "Observacao": "Observacao",
    "DataCriacao": "Cria_x00e7__x00e3_o",
    "DataIndicador": "DataIndicador",
    "Conteiner": "Ctnr_Lote",
    "Solicitante": "Solicitante",
    "TipoDocumento": "TipoDocumento",
    "Setor": "Setor",
    "Servico": "Servico",
    "CanceladoSistema": "CanceladoSistema",
}
CAMPOS_SISCOMEX = ["Adicoes", "Cemercante", "Tipoembalagem", "Quantidade", "PesoBruto"]
CAMPOS_BANCO = {
    "MERCID": "MercId",
    "NRO_BL": "BL_TOS",
    "NAVIO": "Navio",
    "VIAGEM": "Viagem_TOS",
    "CONTEINER": "Ctnr_Lote_TOS",
    "CLASSE": "Classecarga",
    "TAMANHO": "Tamanho_TOS",
    "CLIENTE": "Cliente_TOS",
}


def buscar_carga_por_ce(ce_mercante: str) -> pd.DataFrame:
    """Uma linha por contêiner/lote da carga. O CE vai como bind variable."""
    query = Path(os.getenv("ARQUIVO_SQL", "consulta_carga_por_ce.sql")).read_text(encoding="utf-8")
    with oracledb.connect(user=os.environ["ORACLE_USUARIO"], password=os.environ["ORACLE_SENHA"],
                          dsn=os.environ["ORACLE_DSN"]) as conexao:
        with conexao.cursor() as cursor:
            cursor.execute(query, ce=ce_mercante)
            colunas = [c[0].upper() for c in cursor.description]
            return pd.DataFrame(cursor.fetchall(), columns=colunas)


def vistorias_pendentes(sp: SharePoint, lista: str) -> pd.DataFrame:
    df = sp.ler_lista(lista, todos=True)
    if df.empty:
        return df

    def coluna(nome: str) -> pd.Series:
        valores = df[nome] if nome in df else pd.Series("", index=df.index)
        return valores.fillna("").astype(str).str.strip()

    filtro = (
        (coluna("TipoDocumento") == "DI")
        & (coluna("Cemercante") == "")
        & (coluna("RPASiscomex") == "")
        & (coluna("CanceladoSistema") != "Cancelado")
    )
    return df[filtro]


def montar_item(valores: dict, tipos: dict) -> dict:
    return {campo: converter_valor(valor, tipos.get(campo, "Text")) for campo, valor in valores.items()}


def main() -> None:
    load_dotenv()
    lista = os.environ["SHAREPOINT_LISTA_VISTORIAS"]
    pausa = int(os.getenv("PAUSA_ENTRE_CONSULTAS_SEGUNDOS", "240"))
    sp = SharePoint(os.environ["SHAREPOINT_SITE_URL"], os.environ["SHAREPOINT_USUARIO"],
                    os.environ["SHAREPOINT_SENHA"])

    pendentes = vistorias_pendentes(sp, lista)
    if pendentes.empty:
        print("Nenhuma declaração para processar.")
        return

    tipos = sp.tipos_dos_campos(lista)
    siscomex = SiscomexImportacao()
    try:
        siscomex.aguardar_login_com_certificado()
        for _, item in pendentes.iterrows():
            item_id, declaracao = int(item["Id"]), item[CAMPOS_BASE["Declaracao"]]
            print(f"Processando declaração {declaracao} (item {item_id})")

            try:
                dados_di = siscomex.dados_gerais(declaracao)
            except Exception as e:
                print(f"Falha no Siscomex: {e}")
                sp.editar_item(lista, item_id, {"RPASiscomex": "Falha"})
                continue

            ce = dados_di.get("Cemercante", "").strip()
            cargas = buscar_carga_por_ce(ce) if ce else pd.DataFrame()
            valores_di = {campo: dados_di.get(campo, "") for campo in CAMPOS_SISCOMEX}

            if cargas.empty:
                valores = {**valores_di, "RPASiscomex": "OK mas sem dados do banco"}
                sp.editar_item(lista, item_id, montar_item(valores, tipos))
            else:
                for posicao, (_, carga) in enumerate(cargas.iterrows()):
                    valores_banco = {campo: carga.get(coluna, "") for coluna, campo in CAMPOS_BANCO.items()}
                    valores = {**valores_di, **valores_banco, "RPASiscomex": "OK"}
                    if posicao == 0:
                        sp.editar_item(lista, item_id, montar_item(valores, tipos))
                    else:
                        # Contêineres adicionais viram novas linhas com os dados da solicitação original.
                        base = {campo: item.get(campo, "") for campo in CAMPOS_BASE.values()}
                        sp.adicionar_item(lista, montar_item({**base, **valores}, tipos))

            # O Siscomex limita a frequência de consultas.
            time.sleep(pausa)
    finally:
        siscomex.fechar()
    print("Processo finalizado.")


if __name__ == "__main__":
    main()
