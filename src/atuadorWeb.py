from selenium import webdriver   
from selenium.webdriver.remote.webelement import WebElement   
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pyautogui import press
from time import sleep


class Interagente:
    """
    Classe para interagir com páginas web usando Selenium WebDriver.
    Permite navegar em páginas, preencher campos, esperar por elementos, e outras interações.
    """

    def __init__(self):
        pass

    def abrir_pagina_web(self, link):
        servico = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=servico)
        self.driver.get(link)
        self.driver.maximize_window()


    def interagir_pagina_web(self, xpath, acao, texto="", limitar_espera=False, limitar_retorno=False):
        """
        Realiza ações em um elemento da página web, como clicar, escrever ou retornar o próprio elemento.

        Parâmetros:
            xpath (str): O XPath do elemento a ser encontrado.
            acao (str): A ação a ser executada. Pode ser "Clicar", "Escrever", "Retornar elemento" ou "Esperar".
            texto (str, opcional): Texto a ser inserido caso a ação seja "Escrever".
            limitar_espera (bool, opcional): Limita o tempo de espera para encontrar o elemento.
            limitar_retorno (bool, opcional): Limita o número de tentativas de encontrar o elemento.

        Retorna:
            WebElement ou None: Retorna o elemento se a ação for "Retornar elemento", caso contrário retorna None.
        """
        
        aux = 0
        while True:
            try:
                elemento = self.driver.find_element(By.XPATH, xpath)
                match acao:
                    case "Clicar":
                        elemento.click()
                    case "Escrever":
                        elemento.clear()
                        sleep(0.8)
                        elemento.send_keys(texto)
                    case "Retornar elemento":
                        return elemento
                    case "Esperar":
                        pass
                break
            except:
                sleep(1)
                if limitar_espera == True:
                    aux+=1
                if limitar_retorno == True:
                    aux+=7.5
                if aux == 15:
                    break


    def inserir_arquivo(self, xpath, xpath_de_espera, arquivo):
        """
        Insere um arquivo no campo de input da página e aguarda até que a ação seja concluída.

        Parâmetros:
            xpath (str): O XPath do campo onde o arquivo será inserido.
            xpath_de_espera (str): O XPath do elemento de espera após o arquivo ser inserido.
            arquivo (str): O caminho do arquivo a ser inserido.
        """
        
        self.interagir_pagina_web(xpath, acao="Escrever", texto=arquivo)
        self.interagir_pagina_web(xpath_de_espera, acao="Esperar", limitar_espera=True)


    def migrar_ao_frame(self, acao, indice=0):
        match acao:
            case "ir":
                self.driver.switch_to.frame(indice)
            case "voltar":
                self.driver.switch_to.default_content()
            case "Aceitar alerta":
                try:
                    WebDriverWait(self.driver, 15).until(EC.alert_is_present())
                    alert = self.driver.switch_to.alert
                    alert.accept()
                    return True
                except:
                    return False


    def interagir_javaScript(self, venctos_convertidos, indice, id):
        """
        Executa um script JavaScript para preencher um campo de data na página web.

        Parâmetros:
            venctos_convertidos (list): Lista de datas a serem inseridas.
            indice (int): Índice da data a ser inserida.
            id (str): O ID do campo de data na página.
        """
        
        date_field = self.driver.find_element(By.ID, id)
        self.driver.execute_script("arguments[0].style.display = 'block';", date_field)
        self.driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input'));
            arguments[0].dispatchEvent(new Event('change'));
        """, date_field, venctos_convertidos[indice])


    def verificar_instabilidade(self, verificar):
        """
        Verifica se há instabilidade na página, verificando se os arquivos foram corretamente inseridos.
        """
        
        arquivo1_inserido = self.interagir_pagina_web(xpath="/html/body/main/div/div/div/div/form/div/div/div/div[1]/div/div[2]/div[1]/p/span/a", acao="Retornar elemento", limitar_retorno=True)
        arquivo2_inserido = self.interagir_pagina_web(xpath="/html/body/main/div/div/div/div/form/div/div/div/div[1]/div/div[2]/div[2]/p/span/a", acao="Retornar elemento", limitar_retorno=True)
        arquivo3_inserido = self.interagir_pagina_web(xpath="/html/body/main/div/div/div/div/form/div/div/div/div[1]/div/div[2]/div[3]/p/span/a", acao="Retornar elemento", limitar_retorno=True)
        if verificar == "Um de cada vez":   
            if any(isinstance(elemento, WebElement) for elemento in [arquivo1_inserido, arquivo2_inserido, arquivo3_inserido]):
                return False
            press("f5")
            return True
        else:
            if all(isinstance(elemento, WebElement) for elemento in [arquivo1_inserido, arquivo2_inserido, arquivo3_inserido]):
                return False
            press("f5")
            return True


    def fechar_driver(self):
        self.driver.quit()
    
