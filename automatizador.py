# Precisa dar os comandos pip install selenium e pip install pywin32
import time
import getpass  # Importa a biblioteca para senhas seguras
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# --- Configuração Segura ---
EMAIL_USUARIO = "pedro.henrsilva@mrv.com.br"
# Solicita a senha de forma segura no terminal
SENHA_USUARIO = getpass.getpass(f"Digite a senha para {EMAIL_USUARIO}: ")

# --- Configuração do Driver ---
try:
    driver = webdriver.Chrome() 
    driver.get("https://podio.com/login")
    driver.maximize_window()
except Exception as e:
    print(f"Erro ao iniciar o Chrome. Verifique seu chromedriver. {e}")
    exit()

# --- Etapa 1: Clicar no Login da Microsoft ---
try:
    print("Aguardando página de login...")
    microsoft_login_xpath = "//a[@data-provider='live']"
    
    print("Procurando o botão de login da Microsoft...")
    microsoft_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, microsoft_login_xpath))
    )
    
    print("Clicando no botão da Microsoft...")
    microsoft_button.click()

except Exception as e:
    print(f"Erro ao tentar clicar no botão da Microsoft: {e}")
    driver.quit()
    exit()

# --- Etapa 1.5: Lidar com o login da Microsoft (Versão Final) ---
try:
    print("Aguardando a nova janela/aba de login da Microsoft...")
    
    # Espera o pop-up abrir (total de 2 janelas)
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

    # Pega o ID da janela pop-up
    popup_window = None
    original_window = driver.current_window_handle # Guarda a original (que vai fechar)
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            popup_window = window_handle
            break
            
    driver.switch_to.window(popup_window)
    print("Foco mudado para a janela de login da Microsoft.")
            
    print("Aguardando a tela de login da Microsoft...")
    
    # Preenche o e-mail
    email_field_microsoft = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0116")) 
    )
    print("Preenchendo e-mail da Microsoft...")
    email_field_microsoft.send_keys(EMAIL_USUARIO)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
    
    # Preenche a senha
    password_field_microsoft = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0118")) 
    )
    print("Preenchendo senha da Microsoft...")
    password_field_microsoft.send_keys(SENHA_USUARIO)
    
    # Tenta clicar no botão "Entrar" (com loop anti-stale)
    print("Procurando o botão 'Entrar'...")
    tentativas = 0
    clicado_entrar = False
    while not clicado_entrar and tentativas < 5:
        try:
            entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            entrar_button.click()
            clicado_entrar = True
            print("Botão 'Entrar' clicado.")
        except StaleElementReferenceException:
            tentativas += 1; time.sleep(0.5)
    if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

    # --- ESPERA PELO MFA MANUAL ---
    print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
    print("Aguardando aprovação do MFA no seu celular (até 180s)...")
    
    tentativas = 0
    clicado_manter = False
    while not clicado_manter and tentativas < 5:
        try:
            keep_logged_in_button = WebDriverWait(driver, 180).until( 
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            keep_logged_in_button.click() # Clica "Sim"
            clicado_manter = True
            print("MFA Aprovado! Botão 'Manter conectado' clicado.")
        except StaleElementReferenceException:
            tentativas += 1; time.sleep(0.5)
        except TimeoutException:
             print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
             clicado_manter = False; break
    if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

    print("Login da Microsoft concluído na janela pop-up.")
    
    # --- LÓGICA CORRIGIDA (Baseada na sua análise) ---
    
    # 1. Espera a janela pop-up fechar sozinha
    print("Aguardando janela pop-up fechar...")
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
    
    # 2. Pega o handle da ÚNICA janela que sobrou (a nova janela principal)
    nova_janela_principal = driver.window_handles[0]
    driver.switch_to.window(nova_janela_principal)
    print("Foco retornado para a janela principal do Podio.")
    
    # 3. Força a navegação para a home page
    print("Forçando navegação para https://podio.com/home")
    driver.get("https://podio.com/home")
    
    # 4. (CORREÇÃO) Espera a página carregar
    print("Aguardando a página carregar (máx 10s)...")
    garantia_xpath = "//h2[normalize-space()='Atividade recente']"
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, garantia_xpath))
    )
    print("Página principal carregada com sucesso.")
    
except Exception as e:
    print(f"Erro durante o login na Microsoft (Etapa 1.5): {e}")
    driver.save_screenshot("erro_etapa_1-5.png") 
    driver.quit()
    exit()

# --- Início da Navegação no Podio (Etapas 2-5) ---
try:
    # --- ETAPA 2 CORRIGIDA (COM HOVER SIMULADO) ---
    print("Etapa 2: Procurando 'Vá para uma área de trabalho'...")
    
    # 1. Encontra o elemento 'pai'
    parent_element_xpath = "//div[contains(@class, 'space-switcher-wrapper')]"
    parent_element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, parent_element_xpath))
    )
    
    # 2. Simula o mouse passando por cima dele
    print("Simulando passagem do mouse (hover)...")
    actions = ActionChains(driver)
    actions.move_to_element(parent_element).perform()
    
    # 3. (CORREÇÃO) Espera o botão/texto de dentro aparecer
    workspace_switcher_xpath = "//div[@class='default-title' and contains(text(), 'Vá para uma área de trabalho')]"
    workspace_switcher = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, workspace_switcher_xpath))
    )
    
    # 4. (CORREÇÃO) Clica
    print("Elemento visível. Clicando...")
    driver.execute_script("arguments[0].click();", workspace_switcher)
    print("Clicado em 'Vá para uma área de trabalho'.")
    # --- FIM DA ETAPA 2 ---

    # Etapa 3: Clicar em "ADM - Núcleo Contratos"
    print("Etapa 3: Aguardando a lista de áreas de trabalho...")
    adm_link_xpath = "//a[contains(text(), 'ADM - Núcleo Contratos')]"
    adm_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, adm_link_xpath))
    )
    print("Clicando em 'ADM - Núcleo Contratos'...")
    adm_link.click()

    # (CORREÇÃO) --- ETAPA DE GARANTIA (Página Núcleo Contratos) ---
    print("Aguardando a página do 'Núcleo Contratos' carregar...")
    
    garantia_xpath_etapa3 = "//h2[contains(text(), 'ADM - Núcleo Contratos')]"
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, garantia_xpath_etapa3))
    )
    print("Página 'Núcleo Contratos' carregada.")
    # --- FIM DA GARANTIA ---

    # --- ETAPA 4 (Usando data-app-id) ---
    print("Etapa 4: Procurando o app 'Mensageria'...")
    
    mensageria_app_xpath = "//li[@data-app-id='22830484']"
    
    mensageria_app = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, mensageria_app_xpath))
    )
    print("Clicando no app 'Mensageria'...")
    mensageria_app.click()
    # --- FIM DA ETAPA 4 ---

    
    # --- ETAPA 5 CORRIGIDA (Hover em todos os <li> + Anti-Stale) ---
    
    print("Etapa 5: Aguardando a página 'Mensageria' carregar...")
    
    tentativas = 0
    clicado_filtro = False
    while not clicado_filtro and tentativas < 5: # Tenta por 5 segundos
        try:
            # 1. Espera o <nav> (pai) carregar
            parent_filter_xpath = "//nav[contains(@class, 'app-view-header__config')]"
            parent_filter_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, parent_filter_xpath))
            )
            print(f"Tentativa {tentativas+1}: Container <nav> (pai) encontrado.")

            # 2. Encontra TODOS os <li> (filhos) dentro dele
            child_items_xpath = ".//li"
            child_items = parent_filter_element.find_elements(By.XPATH, child_items_xpath)
            
            if not child_items:
                raise Exception("Container <nav> encontrado, mas nenhum <li> filho foi encontrado.")
            
            print(f"Encontrados {len(child_items)} itens. Passando o mouse sobre eles...")

            # 3. Simula a ação "humana":
            actions_filter = ActionChains(driver)
            for item in child_items:
                actions_filter.move_to_element(item)
            actions_filter.perform() # Executa a sequência de "hovers"
            
            # 4. Agora que o menu está "acordado", espera o <li> "Filtros"
            target_filter_xpath = ".//li[@data-original-title='Filtros']"
            target_filter = WebDriverWait(parent_filter_element, 2).until(
                EC.element_to_be_clickable((By.XPATH, target_filter_xpath))
            )
            
            # 5. Clica
            print("Ícone 'Filtros' acordado e clicável. Clicando...")
            target_filter.click()
            
            clicado_filtro = True # Sucesso! Sai do loop.
            print("Ícone 'Filtros' clicado com sucesso.")
            
        except (StaleElementReferenceException, TimeoutException) as e:
            # 6. Se der "Stale" ou "Timeout", espera e tenta a etapa 5 inteira de novo
            print(f"Tentativa {tentativas+1} falhou (Stale ou Timeout). Página recarregando... Tentando de novo.")
            tentativas += 1
            time.sleep(1) # Espera 1 segundo para a página estabilizar
            
    # 7. Se o loop terminar sem clicar, força um erro
    if not clicado_filtro:
        raise Exception("Falha ao encontrar o ícone 'Filtros' após 5 tentativas.")

    print("Ação concluída com sucesso!")
    time.sleep(10)


except Exception as e:
    print(f"Erro during a navegação no Podio (Etapas 2-5): {e}")
    print("Verifique os seletores XPath. Um deles pode ter mudado.")
    driver.save_screenshot("erro_de_navegacao.png")

finally:
    # Garante que o navegador feche
    print("Fechando o navegador.")
    driver.quit()
