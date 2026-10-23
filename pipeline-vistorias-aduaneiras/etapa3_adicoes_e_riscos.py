"""
Etapa 3 – Detalha as adições de cada DI e sinaliza itens de risco.

Para as vistorias já detalhadas na etapa 2, lê no Siscomex todas as adições
da DI e os itens de mercadoria de cada uma. Compara as descrições com uma
lista de termos de risco mantida no SharePoint, gera uma planilha Excel por
declaração e a anexa aos itens da lista, junto com a quantidade de
referências e a marcação de item de risco.
"""
import os
import tempfile
import time
from collections import defaultdict
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from regras import normalizar
from siscomex_importacao import COLUNAS_ITEM, SiscomexImportacao, formatar_di
from sharepoint_client import SharePoint

CAMPO_DECLARACAO = "Declara_x00e7__x00e3_o"
FILTRO_PENDENTES = "RPASiscomex eq 'OK' and RPAAdicaoReferencia eq null"


def declaracoes_pendentes(sp: SharePoint, lista: str) -> dict[str, list[int]]:
    """{declaração: [ids dos itens]} – uma DI pode ter várias linhas (uma por contêiner)."""
    df = sp.ler_lista(lista, filtro=FILTRO_PENDENTES)
    agrupado = defaultdict(list)
    for _, item in df.iterrows():
        if item.get(CAMPO_DECLARACAO) and item.get("Id"):
            agrupado[item[CAMPO_DECLARACAO]].append(int(item["Id"]))
    return dict(agrupado)


def termos_de_risco(sp: SharePoint, lista: str) -> list[str]:
    df = sp.ler_lista(lista, todos=True)
    if df.empty:
        return []
    return [normalizar(t).strip() for t in df["Title"].dropna() if str(t).strip()]


def para_numero(serie: pd.Series) -> pd.Series:
    """Converte números no formato brasileiro (1.234,56)."""
    return pd.to_numeric(serie.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                         errors="coerce")


def consultar_itens_da_di(siscomex: SiscomexImportacao, declaracao: str) -> tuple[pd.DataFrame, bool]:
    siscomex.abrir_di(declaracao)
    quantidade = siscomex.quantidade_adicoes()
    print(f"DI {formatar_di(declaracao)}: {quantidade} adições")

    itens, houve_falha = [], False
    for numero in range(1, quantidade + 1):
        itens_adicao, falhou = siscomex.itens_da_adicao(numero)
        itens.extend(itens_adicao)
        houve_falha = houve_falha or falhou

    df = pd.DataFrame(itens).rename(columns=COLUNAS_ITEM)
    if not df.empty:
        df["Declaracao"] = formatar_di(declaracao)
        df["Quantidade"] = para_numero(df["Quantidade"]).astype("Int64")
        df["Valor Unitario"] = para_numero(df["Valor Unitario"]).round(2)
    return df, houve_falha


def marcar_riscos(df: pd.DataFrame, termos: list[str]) -> pd.DataFrame:
    def eh_risco(descricao) -> str:
        texto = normalizar(descricao) if pd.notna(descricao) else ""
        return "SIM" if any(termo in texto for termo in termos) else "NÃO"

    df["Item de Risco?"] = df["CodigoDescricao"].apply(eh_risco)
    return df


def main() -> None:
    load_dotenv()
    lista = os.environ["SHAREPOINT_LISTA_VISTORIAS"]
    pausa = int(os.getenv("PAUSA_ENTRE_CONSULTAS_SEGUNDOS", "240"))
    sp = SharePoint(os.environ["SHAREPOINT_SITE_URL"], os.environ["SHAREPOINT_USUARIO"],
                    os.environ["SHAREPOINT_SENHA"])

    pendentes = declaracoes_pendentes(sp, lista)
    if not pendentes:
        print("Nenhuma declaração para processar.")
        return
    print(f"{len(pendentes)} declarações para detalhar.")
    termos = termos_de_risco(sp, os.environ["SHAREPOINT_LISTA_TERMOS_RISCO"])

    siscomex = SiscomexImportacao()
    try:
        siscomex.aguardar_login_com_certificado()
        for declaracao, ids in pendentes.items():
            try:
                itens, houve_falha = consultar_itens_da_di(siscomex, declaracao)
                if itens.empty:
                    for item_id in ids:
                        sp.editar_item(lista, item_id, {"RPAAdicaoReferencia": "Falha", "ItemdeRisco": "NÃO"})
                    continue

                itens = marcar_riscos(itens, termos)
                tem_risco = "SIM" if (itens["Item de Risco?"] == "SIM").any() else "NÃO"
                status = "Possível falha na busca de algum item" if houve_falha else "OK"

                planilha = Path(tempfile.gettempdir()) / f"Adicoes_e_Referencias_{formatar_di(declaracao)}.xlsx"
                itens.to_excel(planilha, index=False)
                for item_id in ids:
                    if not sp.anexar_arquivo(lista, item_id, str(planilha)):
                        status = "OK mas falha ao anexar arquivo"
                planilha.unlink(missing_ok=True)

                for item_id in ids:
                    sp.editar_item(lista, item_id, {
                        "Referencias": len(itens),
                        "ItemdeRisco": tem_risco,
                        "RPAAdicaoReferencia": status,
                    })
            except Exception as e:
                print(f"Erro ao processar a declaração {declaracao}: {e}")
                for item_id in ids:
                    sp.editar_item(lista, item_id, {"RPAAdicaoReferencia": "Falha", "ItemdeRisco": "NÃO"})

            time.sleep(pausa)
    finally:
        siscomex.fechar()
    print("Processo finalizado.")


if __name__ == "__main__":
    main()
