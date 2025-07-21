import time

while True:
    try:

        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        import time
        import subprocess

        driver = webdriver.Chrome()
        wait = WebDriverWait(driver, 30)
        actions = ActionChains(driver)

        driver.get("https://sgusuite-hml.sgusuite.com.br/login")
        time.sleep(2)

        # LOGIN
        print("[INFO] Fazendo login...")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[1]/div/div/p/input"))).send_keys("joao.trombin@criciuma.unimedsc.com.br")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[2]/div[1]/div/div/p/input"))).send_keys("C@mpinh02134")
        time.sleep(1)
        #wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[3]/button"))).click()

        # PÁGINA PRINCIPAL
        print("[INFO] Entrando no sistema...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/button"))).click()
        time.sleep(2)
        time.sleep(5)

        # MENU → RELATÓRIOS GERADOS
        menu_xpath = '/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[4]/div/div[1]'
        actions.move_to_element(driver.find_element(By.XPATH, menu_xpath)).perform()
        time.sleep(1)

        relatorios_xpath = '/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[4]/div/div[2]/div/a[14]/div/div/div/div/div[1]/div'
        wait.until(EC.element_to_be_clickable((By.XPATH, relatorios_xpath))).click()

        time.sleep(5)

        # ESPERA A TABELA CARREGAR
        tabela_xpath = '//table[contains(@class,"table")]/tbody/tr[1]'
        wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))

        # AGUARDANDO RELATÓRIO FICAR "CONCLUÍDO"
        print("[INFO] Aguardando relatório mais recente ficar 'Concluído'...")

        status_xpath = '//table[contains(@class,"table")]/tbody/tr[1]/td[4]/div'
        tentativas = 0
        max_tentativas = 60  # 60 tentativas de 5s = 5 minutos

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

            # Tenta forçar o refresh da tabela recarregando a página de relatórios
            print("[INFO] Atualizando tabela de relatórios...")
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
            
            tentativas += 1
            time.sleep(5)

        if tentativas == max_tentativas:
            print("[FALHA] Tempo máximo de espera atingido. Relatório não ficou pronto.")
            driver.quit()
            exit()

        # CLICA NO ÍCONE DE DOWNLOAD DA PRIMEIRA LINHA
        download_xpath = '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[4]/div/table/tbody/tr[1]/td[6]/div/a[2]/span/i'

        try:
            download_icon = wait.until(EC.element_to_be_clickable((By.XPATH, download_xpath)))
            download_icon.click()
            print("[OK] Clique no botão de download realizado com sucesso!")
        except Exception as e:
            print(f"[ERRO] Não conseguiu clicar para baixar: {e}")

        # ESPERA FINAL
        time.sleep(10)
        driver.quit()

        subprocess.run(["python", "converte-relatorio.py"])
        break  # Sai do loop se tudo rodou sem erro
    except Exception as e:
        print(f"[ERRO GERAL] Ocorreu um erro: {e}")
        print("[INFO] Tentando novamente em 30 segundos...")
        time.sleep(30)