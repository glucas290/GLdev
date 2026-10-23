"""Robô Selenium para consultas no Siscomex Importação Web (DI e adições)."""
import time

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

URL_PORTAL = "https://www1.siscomex.receita.fazenda.gov.br/"
URL_CONSULTA_DI = "https://www1c.siscomex.receita.fazenda.gov.br/importacaoweb-7/ConsultarDIMenu.do"
URL_ADICAO = "https://www1c.siscomex.receita.fazenda.gov.br/importacaoweb-7/ConsultarDIAdicao.do?numAdicao={:03d}"

XPATH_CAMPO_DI = '//*[@id="nrDeclaracao"]'
XPATH_BOTAO_CONSULTAR = '//*[@id="TABLE_4"]/tbody/tr/td/input[1]'
XPATH_QTD_ADICOES = '//*[@id="TABLE_3"]/tbody/tr/td[2]'
XPATH_ITENS_ADICAO = '//*[@id="id_detMercadoria"]/tbody[2]'

XPATHS_DADOS_GERAIS = {
    "Adicoes": XPATH_QTD_ADICOES,
    "Cemercante": '//*[@id="idTableConhecimentoCarga"]/tbody/tr[3]/td[2]/label',
    "Tipoembalagem": '//*[@id="idTableVolumes"]/tbody/tr[2]/td[2]/label',
    "Quantidade": '//*[@id="idTableVolumes"]/tbody/tr[2]/td[1]/label',
}
# A posição do peso bruto muda quando a DI tem agente de carga informado.
XPATH_LINHA_AGENTE = '//*[@id="idTableDadosCarga"]/tbody/tr[3]/td[1]/label'
XPATH_PESO_COM_AGENTE = '//*[@id="idTableDadosCarga"]/tbody/tr[4]/td[2]/label'
XPATH_PESO_SEM_AGENTE = '//*[@id="idTableDadosCarga"]/tbody/tr[3]/td[2]/label'

COLUNAS_ITEM = {
    "Coluna1": "Itens",
    "Coluna2": "Quantidade",
    "Coluna3": "Unidade Medida",
    "Coluna4": "Valor Unitario",
    "Coluna5": "Detalhe",
}


def formatar_di(numero) -> str:
    return "".join(filter(str.isdigit, str(numero))).zfill(10)


class SiscomexImportacao:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-notifications")
        self.navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def fechar(self) -> None:
        self.navegador.quit()

    def _esperar(self, xpath: str, tempo: int = 10):
        return WebDriverWait(self.navegador, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))

    def _texto_ou_vazio(self, xpath: str) -> str:
        try:
            return self.navegador.find_element(By.XPATH, xpath).text.strip()
        except Exception:
            return ""

    def aguardar_login_com_certificado(self) -> None:
        """O Siscomex pede a escolha do certificado numa janela do sistema
        operacional, que o Selenium não controla. O operador seleciona o
        certificado e confirma no terminal."""
        self.navegador.get(URL_PORTAL)
        input("Selecione o certificado digital no navegador e pressione Enter para continuar...")
        self.navegador.get(URL_CONSULTA_DI)
        self._esperar(XPATH_CAMPO_DI, tempo=30)

    def abrir_di(self, numero) -> None:
        self.navegador.get(URL_CONSULTA_DI)
        campo = self._esperar(XPATH_CAMPO_DI)
        campo.clear()
        # O campo tem máscara em JavaScript: digitar caractere a caractere evita perder dígitos.
        for caractere in formatar_di(numero):
            campo.send_keys(caractere)
            time.sleep(0.05)
        self.navegador.find_element(By.XPATH, XPATH_BOTAO_CONSULTAR).click()
        time.sleep(3)

    def dados_gerais(self, numero) -> dict:
        """Adições, CE Mercante, embalagem, quantidade de volumes e peso bruto da DI."""
        self.abrir_di(numero)
        dados = {chave: self._texto_ou_vazio(xpath) for chave, xpath in XPATHS_DADOS_GERAIS.items()}
        tem_agente = "Agente" in self._texto_ou_vazio(XPATH_LINHA_AGENTE)
        dados["PesoBruto"] = self._texto_ou_vazio(XPATH_PESO_COM_AGENTE if tem_agente else XPATH_PESO_SEM_AGENTE)
        return dados

    def quantidade_adicoes(self) -> int:
        try:
            texto = self._esperar(XPATH_QTD_ADICOES).text.strip()
        except TimeoutException:
            return 0
        try:
            return int(texto.split(":")[-1].strip())
        except ValueError:
            return 0

    def itens_da_adicao(self, numero_adicao: int) -> tuple[list[dict], bool]:
        """Lê os itens de mercadoria de uma adição. Retorna (itens, houve_falha)."""
        url = URL_ADICAO.format(numero_adicao)
        self.navegador.get(url)
        time.sleep(2)
        if self.navegador.current_url != url:
            print(f"Redirecionado: a adição {numero_adicao} não pôde ser aberta.")
            return [], True

        itens = []
        try:
            linhas = self._esperar(XPATH_ITENS_ADICAO).find_elements(By.TAG_NAME, "tr")
        except TimeoutException:
            return [], True

        # Cada item ocupa duas linhas: a de dados e a de detalhe, aberta por um botão.
        for indice_item, linha in enumerate(range(1, len(linhas) + 1, 2), start=1):
            try:
                self.navegador.find_element(By.XPATH, f"{XPATH_ITENS_ADICAO}/tr[{linha}]/td[5]/input").click()
                time.sleep(0.5)
                celulas = self.navegador.find_element(By.XPATH, f"{XPATH_ITENS_ADICAO}/tr[{linha}]")
                textos = [c.text.strip() for c in celulas.find_elements(By.TAG_NAME, "td")]
                try:
                    descricao = self._esperar(f'//*[@id="{indice_item}"]/td/label', tempo=5).text.strip()
                except TimeoutException:
                    descricao = ""

                item = {f"Coluna{i}": texto for i, texto in enumerate(textos, start=1)}
                item["CodigoDescricao"] = descricao
                item["Adicao"] = f"{numero_adicao:03d}"
                itens.append(item)
            except Exception as e:
                print(f"Erro no item {indice_item} da adição {numero_adicao}: {e}")
                return itens, True
        return itens, False
