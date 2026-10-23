"""
Utilitário para integrar automações com o SharePoint Online: listas, anexos,
upload de arquivos e leitura de planilhas Excel guardadas no site.
"""
import io
import os

import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.attachments.creation_information import AttachmentCreationInformation
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


def imprimir_progresso(itens) -> None:
    print(f"{len(itens)} itens carregados...")


def converter_valor(valor, tipo_campo: str):
    """Converte um valor Python para o formato aceito pelo tipo de campo do SharePoint."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)) or valor is pd.NaT:
        return None
    if isinstance(valor, str) and not valor.strip():
        return None
    try:
        if tipo_campo == "Number":
            if isinstance(valor, str):
                valor = valor.replace(".", "").replace(",", ".")
            return float(valor)
        if tipo_campo == "DateTime":
            return pd.Timestamp(valor).strftime("%Y-%m-%dT%H:%M:%S")
        if tipo_campo == "Boolean":
            return bool(valor)
        return str(valor)
    except (ValueError, TypeError):
        return None


class SharePoint:
    """Cliente do SharePoint Online com autenticação por usuário e senha."""

    def __init__(self, site_url: str, usuario: str, senha: str):
        self.site_url = site_url
        self.usuario = usuario
        self.senha = senha
        self.ctx = self._conectar()

    def _conectar(self) -> ClientContext | None:
        try:
            autenticacao = AuthenticationContext(self.site_url)
            if autenticacao.acquire_token_for_user(self.usuario, self.senha):
                print("Conexão com o SharePoint estabelecida.")
                return ClientContext(self.site_url, autenticacao)
            print(f"Erro de autenticação: {autenticacao.get_last_error()}")
        except Exception as e:
            print(f"Erro ao conectar: {e}")
        return None

    def enviar_arquivo(self, caminho_local: str, pasta_remota: str) -> bool:
        """Faz upload de um arquivo para uma pasta (URL relativa ao servidor)."""
        if not self.ctx:
            return False
        try:
            pasta = self.ctx.web.get_folder_by_server_relative_url(pasta_remota)
            nome = os.path.basename(caminho_local)
            with open(caminho_local, "rb") as arquivo:
                pasta.upload_file(nome, arquivo.read()).execute_query()
            print(f"Arquivo '{nome}' enviado para '{pasta_remota}'.")
            return True
        except Exception as e:
            print(f"Erro ao enviar o arquivo: {e}")
            return False

    def adicionar_item(self, lista: str, dados: dict) -> bool:
        if not self.ctx:
            return False
        try:
            item = self.ctx.web.lists.get_by_title(lista).add_item(dados).execute_query()
            print(f"Item {item.properties['Id']} adicionado à lista '{lista}'.")
            return True
        except Exception as e:
            print(f"Erro ao adicionar item: {e}")
            return False

    def ler_lista(self, lista: str, filtro: str | None = None, limite: int = 0,
                  todos: bool = False, tamanho_pagina: int = 2000) -> pd.DataFrame:
        """Lê uma lista e devolve um DataFrame.

        - filtro: expressão OData (ex.: "ID gt 48000"). Recomendado para listas grandes.
        - todos=True: pagina a leitura para passar do limite de 5000 itens.
        - limite: devolve só os N primeiros itens (ignorado com todos=True).
        """
        if not self.ctx:
            return pd.DataFrame()
        try:
            alvo = self.ctx.web.lists.get_by_title(lista)
            if filtro:
                itens = alvo.items.filter(filtro)
            elif todos:
                itens = alvo.items.paged(tamanho_pagina, page_loaded=imprimir_progresso)
            else:
                itens = alvo.items
            if limite > 0 and not todos:
                itens = itens.top(limite)

            self.ctx.load(itens)
            self.ctx.execute_query()
            df = pd.DataFrame([item.to_json() for item in itens])
            print(f"{len(df)} registros lidos da lista '{lista}'.")
            return df
        except Exception as e:
            print(f"Erro ao ler a lista: {e}")
            return pd.DataFrame()

    def ler_planilha(self, url_arquivo: str, aba=0) -> pd.DataFrame:
        """Lê um arquivo Excel do SharePoint como DataFrame."""
        if not self.ctx:
            return pd.DataFrame()
        try:
            resposta = File.open_binary(self.ctx, url_arquivo)
            resposta.raise_for_status()
            return pd.read_excel(io.BytesIO(resposta.content), sheet_name=aba, engine="openpyxl")
        except Exception as e:
            print(f"Erro ao ler a planilha: {e}")
            return pd.DataFrame()

    def editar_item(self, lista: str, item_id: int, dados: dict) -> bool:
        if not self.ctx:
            return False
        try:
            item = self.ctx.web.lists.get_by_title(lista).get_item_by_id(item_id)
            for campo, valor in dados.items():
                item.set_property(campo, valor)
            item.update()
            self.ctx.execute_query()
            print(f"Item {item_id} da lista '{lista}' atualizado.")
            return True
        except Exception as e:
            print(f"Erro ao editar o item {item_id}: {e}")
            return False

    def tipos_dos_campos(self, lista: str) -> dict:
        """Mapa {nome interno do campo: tipo} da lista (Text, Number, DateTime...)."""
        campos = self.ctx.web.lists.get_by_title(lista).fields.get().execute_query()
        return {campo.internal_name: campo.type_as_string for campo in campos}

    def anexar_arquivo(self, lista: str, item_id: int, caminho_local: str) -> bool:
        if not self.ctx:
            return False
        if not os.path.exists(caminho_local):
            print(f"Arquivo '{caminho_local}' não encontrado.")
            return False
        try:
            item = self.ctx.web.lists.get_by_title(lista).get_item_by_id(item_id)
            nome = os.path.basename(caminho_local)
            with open(caminho_local, "rb") as arquivo:
                anexo = AttachmentCreationInformation(filename=nome, content=arquivo.read())
            item.attachment_files.add(anexo)
            self.ctx.execute_query()
            print(f"Anexo '{nome}' adicionado ao item {item_id}.")
            return True
        except Exception as e:
            print(f"Erro ao anexar arquivo ao item {item_id}: {e}")
            return False
