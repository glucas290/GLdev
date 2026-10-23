"""Robô de consulta aos sistemas da Receita Federal (Siscarga e Siscomex)."""
import re
import time

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

URL_LOGIN_SISCARGA = "http://www4.receita.fazenda.gov.br/g33159/jsp/logon.jsp"
URL_SISCOMEX_WEB = "https://www1.siscomex.receita.fazenda.gov.br/"
URL_CONSULTA_CE = "https://www4c.receita.fazenda.gov.br/carga-web/ConsultarDadosBasicosCEMercanteMenu.do"
URL_CONSULTA_ENTREGA_DI = "https://www4c.receita.fazenda.gov.br/carga-web/ConsultarEntregaDeCargaMenu.do"
URL_CONSULTA_DI = "https://www1c.siscomex.receita.fazenda.gov.br/importacaoweb-7/ConsultarDIMenu.do"
URL_DETALHE_ICMS_DI = "https://www1c.siscomex.receita.fazenda.gov.br/importacaoweb-7/ConsultarDIPorId.do?popupIcms=true"

# Login com certificado digital: seleciona a opção "certificado" e clica em entrar.
XPATH_OPCAO_CERTIFICADO = '//*[@id="cpfsenha"]/table[1]/tbody/tr[2]/td/table/tbody/tr[8]/td/select/option[4]'
XPATH_ENTRAR_CERTIFICADO = '//*[@id="cpfsenha"]/table[1]/tbody/tr[2]/td/table/tbody/tr[4]/td[1]/table/tbody/tr[4]/td[2]/div/a/img'
XPATH_ENTRAR_SISCOMEX_WEB = '//*[@id="cpfsenha"]/table[1]/tbody/tr[2]/td/table/tbody/tr[4]/td/table/tbody/tr[4]/td[2]/div/a/img'

BASE_CE = "/html/body/div[2]/div/div[2]/div[2]/div[1]/div"
BASE_DI = "/html/body/div[2]/div/div[2]/div/div[1]/div/form"

# Palavras-chave que indicam cada pendência nas mensagens da DI.
PENDENCIAS = {
    "SCMX_PEND_ICMS": ["icms não declarado", "icms não realizado", "nao foi solicitado pagamento", "foi solicitado pagamento/exonera"],
    "SCMX_PEND_AFRRM": ["evento de afrmm nao realizado", "evento de afrmm"],
    "SCMX_PEND_FRETEARMADOR": ["retenção da carga por parte do armador"],
    "SCMX_PEND_TUM": ["pagamento da tum nao realizado", "pagamento da tum"],
}

# Avisos exibidos quando a DI ainda não foi desembaraçada.
AVISOS_NAO_DESEMBARACADA = ["não está desembaraçada", "DV da DI", "ara este tipo de DI"]


def somente_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto)


def possui_pendencia(mensagens: list[str], coluna: str) -> str:
    palavras = PENDENCIAS[coluna]
    for mensagem in mensagens:
        if any(p in mensagem.lower() for p in palavras):
            return "Sim"
    return "Não"


class RoboSiscomex:
    def __init__(self, timeout: int = 10):
        self.timeout = timeout
        self.navegador = None

    def iniciar_navegador(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--start-maximized")
        options.add_argument("--lang=pt-BR")
        self.navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    def fechar(self) -> None:
        if self.navegador:
            self.navegador.quit()
            self.navegador = None

    def esperar_elemento(self, xpath: str, tempo: int | None = None):
        espera = WebDriverWait(self.navegador, tempo or self.timeout)
        espera.until(EC.presence_of_element_located((By.XPATH, xpath)))
        return espera.until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def _textos(self, xpath: str) -> list[str]:
        return [e.text for e in self.navegador.find_elements(By.XPATH, xpath)]

    def recarregar_apos_erro(self) -> None:
        try:
            self.navegador.refresh()
            self.navegador.get(URL_LOGIN_SISCARGA)
            time.sleep(3)
        except Exception as e:
            print(f"Erro ao recarregar a página: {e}")

    def logar_siscarga(self) -> None:
        """Login no Siscarga com o certificado digital instalado no navegador."""
        self.navegador.get(URL_LOGIN_SISCARGA)
        self.esperar_elemento(XPATH_OPCAO_CERTIFICADO).click()
        self.esperar_elemento(XPATH_ENTRAR_CERTIFICADO).click()

    def logar_siscomex_web(self) -> None:
        self.navegador.get(URL_SISCOMEX_WEB)
        self.esperar_elemento(XPATH_ENTRAR_SISCOMEX_WEB).click()

    def consultar_ce(self, ce: str) -> dict:
        """Consulta um CE Mercante: documento de despacho, BL, situação e bloqueios."""
        try:
            time.sleep(1)
            self.navegador.get(URL_CONSULTA_CE)
            self.esperar_elemento('//*[@id="principal"]/div/div[1]/div/form/fieldset/p/input').send_keys(ce)
            self.esperar_elemento('//*[@id="principal"]/div/div[1]/div/form/p/img[1]').click()

            try:
                documento = self.esperar_elemento(f"{BASE_CE}/dl[6]/dd[6]").text
            except TimeoutException:
                documento = ""

            return {
                "documento": documento,
                "bl": self.esperar_elemento(f"{BASE_CE}/dl[6]/dd[1]").text,
                "situacao": self.esperar_elemento(f"{BASE_CE}/dl[7]/dd[6]").text,
                "bloqueio_tipos": self._textos(f"{BASE_CE}/table[*]/tbody/tr[1]/td[2]"),
                "bloqueio_motivos": self._textos(f"{BASE_CE}/table[*]/tbody/tr[2]/td[2]"),
                "bloqueio_justificativas": self._textos(f"{BASE_CE}/table[*]/tbody/tr[3]/td[2]"),
                "bloqueio_datas": self._textos(f"{BASE_CE}/table[*]/tbody/tr[4]/td[2]"),
            }
        except Exception:
            self.recarregar_apos_erro()
            raise

    def consultar_di(self, di: str) -> dict:
        """Consulta a entrega de carga de uma DI: desembaraço, VMLE e pendências."""
        vazio = {"desembaracada": "Não", "entrega": "", "vmle": "", "mensagens": []}
        vazio.update({coluna: "" for coluna in PENDENCIAS})
        try:
            time.sleep(1)
            self.navegador.get(URL_CONSULTA_ENTREGA_DI)
            self.esperar_elemento(f"{BASE_DI}/fieldset/p/input[3]").send_keys(di)
            self.esperar_elemento('//*[@id="principal"]/div/div[1]/div/form/p/img[1]').click()

            try:
                aviso = self.esperar_elemento(f"{BASE_DI}/ul/li", tempo=5).text
            except TimeoutException:
                aviso = ""

            if aviso:
                if any(a in aviso for a in AVISOS_NAO_DESEMBARACADA):
                    return vazio
                raise RuntimeError(f"Aviso inesperado ao consultar a DI {di}: {aviso}")

            mensagens = self._textos(f"{BASE_DI}/fieldset/p[*]")
            resultado = {
                "desembaracada": "Sim",
                "entrega": self.esperar_elemento(f"{BASE_DI}/fieldset/p[1]").text,
                "vmle": self.esperar_elemento(f"{BASE_DI}/fieldset/dl/dd[2]").text,
                "mensagens": mensagens,
            }
            resultado.update({coluna: possui_pendencia(mensagens, coluna) for coluna in PENDENCIAS})
            return resultado
        except Exception:
            self.recarregar_apos_erro()
            raise

    def consultar_uf_destino_di(self, di: str) -> str:
        """Consulta no Siscomex Web a UF de destino informada na DI (aba de ICMS)."""
        time.sleep(1)
        self.navegador.get(URL_CONSULTA_DI)
        self.esperar_elemento('//*[@id="nrDeclaracao"]').send_keys(di)
        self.esperar_elemento('//*[@id="TABLE_4"]/tbody/tr/td/input[1]').click()
        time.sleep(2)
        self.navegador.get(URL_DETALHE_ICMS_DI)
        time.sleep(2)
        xpath_uf = '//*[@id="TABLE_1"]/tbody/tr[5]/td[2]'
        try:
            uf = self.esperar_elemento(xpath_uf).text
        except TimeoutException:
            self.navegador.refresh()
            time.sleep(2)
            uf = self.esperar_elemento(xpath_uf).text
        return uf.strip()
