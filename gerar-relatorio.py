from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
import unicodedata
import traceback
import re
import subprocess
from datetime import datetime
import threading
import time
import sys

def main():
    driver = None
    try:

        inicio_execucao = time.time()
        stop_event = threading.Event()

        def mostrar_tempo():
            while not stop_event.is_set():
                tempo_atual = time.time() - inicio_execucao
                tempo_formatado = time.strftime('%H:%M:%S', time.gmtime(tempo_atual))

                # Salva posição do cursor
                sys.stdout.write("\0337")
                # Move para topo da tela (linha 1, coluna 1)
                sys.stdout.write("\033[1;1H")
                # Limpa a linha
                sys.stdout.write("\033[K")
                # Escreve o tempo fixo no topo
                sys.stdout.write(f"🕒 Tempo de execução: {tempo_formatado}\n")
                # Restaura posição do cursor
                sys.stdout.write("\0338")
                sys.stdout.flush()

        # Inicia thread que mostra o tempo fixo no topo
        thread_tempo = threading.Thread(target=mostrar_tempo, daemon=True)
        thread_tempo.start()



        #Função para normalizar texto
        def normaliza_texto(texto):
            texto = texto.lower()
            texto = unicodedata.normalize('NFKD', texto)
            texto = ''.join(c for c in texto if not unicodedata.combining(c))
            texto = re.sub(r'\s+', ' ', texto)
            texto = texto.replace('º', 'o').replace('°', 'o').replace('nº', 'no').replace('n°', 'no')
            texto = texto.strip()
            return texto

        #Mapeamento de aliases para os campos permitidos
        aliases = {
            "CONTA": ["conta", "contrato", "nome da conta", "nome da conta social"],
            "PROTOCOLO": ["protocolo", "protocolo n°", "protocolo nº", "numero de protocolo", "protocolo numero"],
            "IDENFICADOR": ["identificador", "idenficador"],
            "NOME DA CONTA": ["nome da conta", "nome da conta social"],
            "DATA DE ABERTURA": ["data de abertura", "data do atendimento"],
            "STATUS": ["status"],
            "STATUS GPU": ["status gpu"],
            "RESPONSAVEL": ["responsavel", "responsável"],
            "RESPONSAVEL ORIGEM": ["responsavel origem", "responsável origem"],
            "DATA DE ENCERRAMENTO": ["data de encerramento"],
            "DESCRIÇÃO DE ATENDIMENTO": ["descrição de atendimento", "descricao de atendimento", "descrição", "descricao"],
            "DESCRIÇÃO DE ENCERRAMENTO": ["descrição de encerramento", "descricao de encerramento"],
            "CATEGORIA": ["categoria"],
            "MOTIVO": ["motivo"],
            "SUBMOTIVO": ["submotivo", "sub motivo"],
            "CANAL DE ENTRADA": ["canal de entrada"],
            "GRUPO ATENDIMENTO ATUAL": ["grupo de atendimento atual", "grupo atendimento atual"],
            "GRUPO ATENDIMENTO ORIGEM": ["grupo de atendimento origem", "grupo atendimento origem"],
            "TIPO DE SOLICITAÇÃO": ["tipo de solicitação", "tipo de solicitacao"],
            "MOVIMENTAÇÕES": ["movimentações", "movimentacoes"]
        }

        #Função para verificar se o campo do site corresponde ao campo permitido
        def campo_igual(nome_base, nome_site):
            nome_site_norm = normaliza_texto(nome_site)
            aliases_list = aliases.get(nome_base.lower().upper(), [nome_base.lower()])
            for alias in aliases_list:
                if normaliza_texto(alias) == nome_site_norm:
                    return True
            return False


        campos_permitidos = list(aliases.keys())

        # Calcula a data de ontem no formato esperado (ex: 16/07/2025)
        ontem = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')


        def js_get_text(driver, xpath):
            js = """
            var xpath = arguments[0];
            var el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if(el) { return el.innerText || el.textContent; }
            return "";
            """
            return driver.execute_script(js, xpath)

        def js_click(driver, xpath):
            js = """
            var xpath = arguments[0];
            var el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if(el) { el.click(); return true; }
            return false;
            """
            return driver.execute_script(js, xpath)

        def selecionar_protocolo_js(driver):
            dropdown_xpath = '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/div[1]/div[2]/section/div[2]/div/div/p/div/div[2]'
            primeiro_item_xpath = dropdown_xpath + '/div[2]/ul/li[1]'

            print("[INFO] Verificando se já existe protocolo selecionado...")
            texto = js_get_text(driver, dropdown_xpath).strip()
            print(f"[DEBUG] Texto no dropdown protocolo: '{texto}'")

            if texto != "":
                print("[INFO] Protocolo já selecionado, pulando seleção.")
                return
            else:
                print("[INFO] Nenhum protocolo selecionado. Abrindo dropdown via JS...")
                if not js_click(driver, dropdown_xpath):
                    print("[ERRO] Não conseguiu clicar no dropdown via JS.")
                    return
                time.sleep(0.2)
                print("[INFO] Selecionando o primeiro item do dropdown via JS...")
                if not js_click(driver, primeiro_item_xpath):
                    print("[ERRO] Não conseguiu clicar no primeiro item via JS.")
                    return
                time.sleep(0.2)
                print("[INFO] Protocolo selecionado com sucesso.")


        def click_js(driver, element):
            driver.execute_script("arguments[0].click();", element)

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        print("[INFO] Iniciando navegador e acessando o site...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 20)
        actions = ActionChains(driver)

        driver.get("https://sgusuite-hml.sgusuite.com.br/login")
        time.sleep(2)

        #LOGIN
        print("[INFO] Fazendo login...")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[1]/div/div/p/input"))).send_keys("joao.trombin@criciuma.unimedsc.com.br")
        wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[2]/div[1]/div/div/p/input"))).send_keys("C@mpinh02134")
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/section/div/div/div/div[3]/div/div/form/div[3]/button"))).click()

        #PÁGINA PRINCIPAL
        print("[INFO] Entrando no sistema...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[2]/button"))).click()
        time.sleep(2)

        #MENU RELATÓRIO
        print("[INFO] Acessando menu de relatórios...")
        menu_principal = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[4]/div/div[1]/a/div")))
        actions.move_to_element(menu_principal).perform()
        time.sleep(1)

        subitem_protocolo = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div[1]/aside/div[2]/ul/li[4]/div/div[2]/div/a[1]/div/div/div/div/div[1]/div")))
        subitem_protocolo.click()
        time.sleep(1)

        print("[INFO] Esperando tela de protocolos carregar...")
        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[1]/nav')))

        #Clica no campo solicitado ao entrar na tela de protocolos
        print("[INFO] Clicando no campo inicial da tela de protocolos...")
        elemento_para_clicar = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[1]/nav/div[2]/div/div/div/div/div[1]/p/label/span[1]')))
        elemento_para_clicar.click()
        time.sleep(1)

        print("[INFO] Fechando menu lateral...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body"))).click()
        time.sleep(1)

        #FILTRA PELA DATA DE ONTEM
        print(f"[INFO] Selecionando filtros de data: {ontem}")
        icone_calendario = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/form/div[2]/div/div/span[2]')))
        icone_calendario.click()
        time.sleep(0.3)

        input_data_inicial = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/form/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/div/div/p/input')))
        input_data_inicial.clear()
        input_data_inicial.send_keys(ontem)
        input_data_inicial.send_keys(Keys.ENTER)
        time.sleep(1)

        input_data_final = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/form/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/div/p/input')))
        input_data_final.clear()
        input_data_final.send_keys(ontem)
        input_data_final.send_keys(Keys.ENTER)
        time.sleep(2)


        print("[INFO] Selecionando protocolo no dropdown (via JS)...")
        selecionar_protocolo_js(driver)

        #GERAÇÃO DE RELATÓRIO
        print("[INFO] Gerando relatório...")
        btn_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/form/div[1]/div/div[7]/button[2]')))
        btn_gerar.click()
        time.sleep(2)

        #LIMPA CAMPOS NÃO PERMITIDOS E ADICIONA OS QUE FALTAM

        #normalizar nomes
        def normalizar(nome):
            return nome.lower().replace("º", "").replace("nº", "").replace("n°", "").replace("int/ext", "").replace("-", "").strip()

        def campo_igual(c1, c2):
            return normalizar(c1) == normalizar(c2)

        # Lista dos campos desejados
        campos_permitidos = [
            "IDENTIFICADOR",
            "NOME DA CONTA",
            "DATA DE ABERTURA",
            "PROTOCOLO Nº",
            "STATUS",
            "STATUS GPU",
            "RESPONSÁVEL",
            "RESPONSÁVEL ORIGEM",
            "DATA DE ENCERRAMENTO",
            "TEMPO MÉDIO DE ATENDIMENTO (TMA)",
            "DESCRIÇÃO DO ENCERRAMENTO",
            "DESCRIÇÃO",
            "CATEGORIA",
            "MOTIVO",
            "SUBMOTIVO",
            "CANAL DE ENTRADA",
            "GRUPO ATENDIMENTO ATUAL",
            "GRUPO ATENDIMENTO ORIGEM",
            "TIPO DE SOLICITAÇÃO",
            "MOVIMENTAÇÕES"
        ]

        # Mapeia possíveis aliases
        aliases = {
            "IDENTIFICADOR": "Identificador",
            "NOME DA CONTA": "Conta",
            "DATA DE ABERTURA": "Data de abertura",
            "PROTOCOLO Nº": "Protocolo nº",
            "STATUS": "Status",
            "STATUS GPU": "Status GPU",
            "RESPONSÁVEL": "Responsável",
            "RESPONSÁVEL ORIGEM": "Responsável origem",
            "DATA DE ENCERRAMENTO": "Data de encerramento",
            "TEMPO MÉDIO DE ATENDIMENTO (TMA)": "Tempo médio atendimento (TMA)",
            "DESCRIÇÃO DO ENCERRAMENTO": "Descrição do encerramento",
            "DESCRIÇÃO": "Descrição",
            "CATEGORIA": "Categoria",
            "MOTIVO": "Motivo",
            "SUBMOTIVO": "Submotivo",
            "CANAL DE ENTRADA": "Canal de entrada",
            "GRUPO ATENDIMENTO ATUAL": "Grupo de atendimento atual",
            "GRUPO ATENDIMENTO ORIGEM": "Grupo de atendimento origem",
            "TIPO DE SOLICITAÇÃO": "Tipo de solicitação",
            "MOVIMENTAÇÕES": "Movimentações"
        }


        container_xpath = '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/div[1]/div[2]/section/div[3]'
        container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
        campos_atual_selecionado = []

        print("[INFO] Lendo campos já selecionados...")
        tags = container.find_elements(By.CLASS_NAME, 'protocol-report__tag')
        for tag in tags:
            span_texto = tag.find_element(By.CLASS_NAME, 'protocol-report__tag__text').text.strip()
            campos_atual_selecionado.append(span_texto)

        print(f"[INFO] Selecionados atualmente: {campos_atual_selecionado}")

        # Remove campos não permitidos
        for tag in tags:
            texto = tag.find_element(By.CLASS_NAME, 'protocol-report__tag__text').text.strip()
            valido = any(
                campo_igual(texto, alvo) or texto in aliases.get(alvo, [])
                for alvo in campos_permitidos
            )
            if not valido:
                print(f"[REMOVER] Campo não permitido encontrado: {texto}")
                try:
                    btn_x = tag.find_element(By.CLASS_NAME, 'protocol-report__tag_button')
                    driver.execute_script("arguments[0].click();", btn_x)
                    print(f"[REMOVER] Removido com sucesso: {texto}")
                    time.sleep(0.3)
                except Exception as e:
                    print(f"[ERRO] Não foi possível remover {texto}: {e}")

        # Atualiza lista depois da limpeza
        tags = container.find_elements(By.CLASS_NAME, 'protocol-report__tag')
        campos_atuais = []
        for tag in tags:
            texto = tag.find_element(By.CLASS_NAME, 'protocol-report__tag__text').text.strip()
            for campo_permitido in campos_permitidos:
                if campo_igual(texto, campo_permitido) or texto in aliases.get(campo_permitido, []):
                    campos_atuais.append(campo_permitido)
                    break

        # Calcula campos faltando
        faltando = [campo for campo in campos_permitidos if campo not in campos_atuais]
        print(f"[INFO] Campos faltando para adicionar: {faltando}")

        # Abre dropdown para adicionar campos
        dropdown_btn_xpath = '//div[contains(@class,"choices")]/div[contains(@class,"choices__inner")]'
        dropdown_input_xpath = '//input[contains(@class,"choices__input--cloned")]'
        first_option_xpath = '//div[contains(@class,"choices__list")]/div[contains(@class,"choices__item--choice")][1]'

        try:
            dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_btn_xpath)))
            dropdown_btn.click()
            print("[INFO] Dropdown aberto.")
            time.sleep(0.3)
        except Exception as e:
            print(f"[ERRO] Falha ao abrir dropdown: {e}")

        #Adiciona os campos faltantes
        for campo in faltando:
            try:
                input_box = wait.until(EC.presence_of_element_located((By.XPATH, dropdown_input_xpath)))
                input_box.clear()
                input_box.send_keys(campo)
                print(f"[ADICIONAR] Buscando: {campo}")
                time.sleep(0.3)

                opcoes_dropdown = wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, '//div[contains(@class,"choices__list")]/div[contains(@class,"choices__item--choice")]')))
                
                opcao_clicada = False

                for opcao in opcoes_dropdown:
                    texto_opcao = opcao.text.strip()
                    if campo_igual(campo, texto_opcao) or texto_opcao in aliases.get(campo, []):
                        print(f"[ADICIONAR] Clicando na opção: '{texto_opcao}' para campo '{campo}'")

                        
                        driver.execute_script("arguments[0].scrollIntoView(true);", opcao)
                        time.sleep(0.3)

                        try:
                            
                            actions.move_to_element(opcao).click().perform()
                        except Exception as e:
                            print(f"[AVISO] Clique via ActionChains falhou: {e}. Tentando clique via JS...")
                            try:
                                driver.execute_script("arguments[0].click();", opcao)
                            except Exception as e2:
                                print(f"[ERRO] Clique via JS também falhou: {e2}")

                        print(f"[ADICIONADO] Campo adicionado: {campo}")
                        opcao_clicada = True
                        time.sleep(0.3)
                        break

                if not opcao_clicada:
                    print(f"[IGNORADO] Nenhuma opção compatível encontrada para '{campo}'")

            except Exception as e:
                print(f"[FALHA] Não conseguiu adicionar '{campo}': {e}")
            time.sleep(0.5)

        print("[FINALIZADO] Seleção e limpeza concluídas.")
        time.sleep(0.7)

        #PREENCHE NOME DO RELATORIO
        print("[INFO] Preenchendo nome do relatório...")
        campo_nome = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/div[1]/div[2]/section/div[1]/div/div/p/input')))
        data_ontem_para_nome = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        nome_relatorio = f"Relatorio-{data_ontem_para_nome}"

        campo_nome.clear()
        campo_nome.send_keys(nome_relatorio)
        time.sleep(0.7)

        #INICIA RELATORIO
        print("[INFO] Iniciando relatório...")
        btn_iniciar = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div/div[2]/section/div[2]/div[2]/div[1]/div[2]/footer/div/div/button[1]')
        btn_iniciar.click()
        time.sleep(0.6)

        print("[SUCESSO] A automação foi realizada com sucesso!")
        stop_event.set()
        thread_tempo.join()
        driver.quit()

        try:
            subprocess.run(["python", "baixa-relatorio.py"], check=True)
            print("[OK] Script de download finalizado.")
        except subprocess.CalledProcessError as e:
            print(f"[ERRO] Falha ao rodar o segundo script: {e}")

        print("[OK] gerar-relatorio finalizado com sucesso.")

    except (WebDriverException, TimeoutException, Exception) as e:
        print("\n[ERRO] Ocorreu um erro durante a execução do gerar-relatorio.py:")
        traceback.print_exc()

        

    finally:
        if driver:
            print("[INFO] Fechando navegador e encerrando sessão.")
            driver.quit()

if __name__ == "__main__":
    main()
