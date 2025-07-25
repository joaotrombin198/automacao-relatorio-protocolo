import time
import os
import subprocess
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv

while True:
    try:
        load_dotenv()
        usuario = os.getenv("USUARIO_SGUSUITE")
        senha = os.getenv("SENHA_SGUSUITE")

        # Caminho da pasta onde o script está rodando
        pasta_script = os.path.join(str(Path().resolve()), "relatorios")
        os.makedirs(pasta_script, exist_ok=True)

        # Preferências do Chrome para definir pasta de download
        prefs = {
            "download.default_directory": pasta_script,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_experimental_option("prefs", prefs)

        print("[INFO] Iniciando navegador e acessando o site...")
        driver = webdriver.Chrome(service=Service("/usr/bin/chromedriver"), options=options)
        wait = WebDriverWait(driver, 30)
        actions = ActionChains(driver)

        driver.get("https://sgusuite-prd.sgusuite.com.br/login")
        time.sleep(2)

        # LOGIN
        print("[INFO] Fazendo login...")
        xpath_email = "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[1]/div/div/p/input"
        xpath_senha = "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[2]/div[1]/div/div/p/input"
        xpath_botao_login = "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[3]/button"

        tentativas_login = 3
        for tentativa in range(1, tentativas_login + 1):
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, xpath_email))).send_keys(usuario)
                wait.until(EC.presence_of_element_located((By.XPATH, xpath_senha))).send_keys(senha)
                time.sleep(1)
                wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_login))).click()
                break
            except Exception:
                print(f"[AVISO] Erro ao preencher login (tentativa {tentativa}). Dando F5...")
                driver.refresh()
                time.sleep(3)
        else:
            raise Exception("[ERRO] Não foi possível realizar o login após várias tentativas.")

        # PÁGINA PRINCIPAL
        print("[INFO] Entrando no sistema...")
        xpath_botao_sistema = "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/button"
        for tentativa in range(10):
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_sistema))).click()
                break
            except Exception:
                print(f"[AVISO] Botão de entrada no sistema não disponível (tentativa {tentativa+1}). Dando F5...")
                driver.refresh()
                time.sleep(5)
        else:
            raise Exception("[ERRO] Não foi possível acessar o sistema após várias tentativas.")

        time.sleep(2)

        # MENU → RELATÓRIOS GERADOS
        menu_xpath = '/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[3]/div/div[1]/a/div'
        relatorios_xpath = '/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[3]/div/div[2]/div/a[3]/div/div/div/div/div[1]/div'
        for tentativa in range(3):
            try:
                actions.move_to_element(driver.find_element(By.XPATH, menu_xpath)).perform()
                time.sleep(1)
                wait.until(EC.element_to_be_clickable((By.XPATH, relatorios_xpath))).click()
                break
            except Exception:
                print(f"[AVISO] Erro ao abrir menu de relatórios (tentativa {tentativa+1}). Dando F5...")
                driver.refresh()
                time.sleep(5)
        else:
            raise Exception("[ERRO] Não foi possível acessar os relatórios após várias tentativas.")

        # ESPERA A TABELA CARREGAR
        tabela_xpath = '//table[contains(@class,"table")]/tbody/tr[1]'
        for tentativa in range(3):
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
                print("[INFO] Tabela de relatórios carregada com sucesso.")
                break
            except Exception:
                print(f"[AVISO] Tabela não carregou (tentativa {tentativa+1}). Dando F5...")
                driver.refresh()
                time.sleep(5)
        else:
            raise Exception("[ERRO] Tabela de relatórios não carregou após várias tentativas.")

        # AGUARDANDO RELATÓRIO FICAR "CONCLUÍDO"
        print("[INFO] Aguardando relatório mais recente ficar 'Concluído'...")
        time.sleep(20)
        status_xpath = '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[4]/div/table/tbody/tr[1]/td[4]'
        tentativas = 0
        max_tentativas = 60

        while tentativas < max_tentativas:
            try:
                status_elem = driver.find_element(By.XPATH, status_xpath)
                status_text = status_elem.text.strip().lower()
                print(f"[INFO] Status atual: {status_text}")

                if "concluído" in status_text:
                    print("[OK] Relatório concluído!")
                    break
            except Exception as e:
                print(f"[ERRO] Tentativa {tentativas+1} ao ler status: {e}")

            print("[INFO] Atualizando tabela de relatórios...")
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
            tentativas += 1
            time.sleep(5)

        if tentativas == max_tentativas:
            print("[FALHA] Tempo máximo de espera atingido. Relatório não ficou pronto.")
            driver.quit()
            exit()

        # CLICA NO ÍCONE DE DOWNLOAD
        download_xpath = '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[4]/div/table/tbody/tr[1]/td[6]/div/a[2]/span/i'
        try:
            download_icon = wait.until(EC.element_to_be_clickable((By.XPATH, download_xpath)))
            download_icon.click()
            print("[OK] Clique no botão de download realizado com sucesso!")
        except Exception as e:
            print(f"[ERRO] Não conseguiu clicar para baixar: {e}")

        time.sleep(10)
        driver.quit()

        subprocess.run(["python", "converte-relatorio.py"])
        break

    except Exception as e:
        print(f"[ERRO GERAL] Ocorreu um erro: {e}")
        print("[INFO] Tentando novamente em 30 segundos...")
        time.sleep(30)
