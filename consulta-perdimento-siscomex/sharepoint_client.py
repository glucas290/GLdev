"""Cliente mínimo para ler e editar listas do SharePoint Online."""
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext


class SharePointClient:
    def __init__(self, site_url: str, usuario: str, senha: str, lista_url: str):
        self.site_url = site_url
        self.usuario = usuario
        self.senha = senha
        self.lista_url = lista_url
        self.ctx = self._autenticar()

    def _autenticar(self) -> ClientContext:
        autenticacao = AuthenticationContext(self.site_url)
        if not autenticacao.acquire_token_for_user(self.usuario, self.senha):
            raise RuntimeError("Falha na autenticação no SharePoint.")
        return ClientContext(self.site_url, autenticacao)

    def ler_itens(self) -> pd.DataFrame:
        """Lê todos os itens da lista e devolve um DataFrame."""
        lista = self.ctx.web.get_list(self.lista_url)
        itens = lista.items.get_all()
        self.ctx.execute_query()
        return pd.DataFrame([dict(item.properties) for item in itens])

    def editar_item(self, item_id: int, valores: dict) -> None:
        lista = self.ctx.web.get_list(self.lista_url)
        item = lista.get_item_by_id(item_id)
        for campo, valor in valores.items():
            item.set_property(campo, valor)
        item.update()
        self.ctx.execute_query()

    def adicionar_item(self, valores: dict) -> None:
        lista = self.ctx.web.get_list(self.lista_url)
        lista.add_item(valores)
        self.ctx.execute_query()
