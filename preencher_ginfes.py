import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time

# --- Função Principal (Orquestrador) ---
def main():
    """Função principal que orquestra toda a automação."""
    
    # --- DADOS FIXOS ---
    cnpj_fixo = "24255199000167"
    senha_fixa = "colibri"
    constantes = {
        "codigo_servico": "8.02/859960100",
        "descricao": "Serviços prestados",
        "aliquota": "2"
    }

    dados_variaveis = coletar_dados_gui()
    
    driver = None
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.maximize_window()
        wait = WebDriverWait(driver, 15)
        driver.get("https://guarulhos.ginfes.com.br/")

        # Chamada das funções de lógica
        realizar_login(driver, wait, cnpj_fixo, senha_fixa)
        preencher_dados_tomador(wait, dados_variaveis)
        preencher_dados_servico(driver, wait, dados_variaveis, constantes)
        emitir_nfs(driver, wait)

        print("\n🎉 Processo finalizado com sucesso! A nota foi emitida.")
        
    except Exception as e:
        print(f"\n❌ Ocorreu um erro: {e}")
    finally:
        if driver:
            input("🔎 Verifique o resultado. Pressione ENTER para fechar o navegador...")
            driver.quit()

# --- Funções de Lógica da Automação ---
def coletar_dados_gui():
    """Cria uma interface gráfica para coletar os dados do cliente e serviço."""
    root = tk.Tk()
    root.withdraw()
    dados = {}
    dados['cpf'] = simpledialog.askstring("Dados do Tomador", "Digite o CPF:")
    dados['nome'] = simpledialog.askstring("Dados do Tomador", "Digite o Nome:")
    dados['cep'] = simpledialog.askstring("Dados do Tomador", "Digite o CEP:")
    dados['numero'] = simpledialog.askstring("Dados do Tomador", "Digite o Número:")
    dados['telefone'] = simpledialog.askstring("Dados do Tomador", "Digite o Telefone:")
    dados['dia'] = simpledialog.askstring("Dados do Serviço", "Digite o Dia (ex: 28):")
    dados['valor_servico'] = simpledialog.askstring("Dados do Serviço", "Digite o Valor do Serviço:")
    return dados

def realizar_login(driver, wait, cnpj, senha):
    """Executa os passos para fazer o login no sistema."""
    print("Iniciando login...")
    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    clicar_com_retry(wait, By.CLASS_NAME, "imagem1")
    
    campo_cnpj = wait.until(EC.visibility_of_element_located((By.XPATH, "//label[text()='CNPJ:']/following::input[1]")))
    campo_cnpj.send_keys(cnpj)

    campo_senha = wait.until(EC.visibility_of_element_located((By.XPATH, "//label[text()='Senha:']/following::input[1]")))
    campo_senha.send_keys(senha)

    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Entrar')]")
    print("✅ Login realizado com sucesso.")

def preencher_dados_tomador(wait, dados):
    """Preenche as informações do tomador do serviço."""
    print("Preenchendo dados do tomador...")
    clicar_com_retry(wait, By.XPATH, "//img[@class='gwt-Image' and contains(@src,'icon_nfse3.gif')]")

    preencher_por_label(wait, "CPF:", dados['cpf'])
    preencher_por_label(wait, "Nome:", dados['nome'])
    preencher_por_label(wait, "CEP:", dados['cep'])
    clicar_com_retry(wait, By.XPATH, "//img[@title='Busca endereço pelo cep']")
    
    preencher_por_label(wait, "Número:", dados['numero'])
    preencher_por_label(wait, "Telefone:", dados['telefone'])
    
    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Próximo Passo')]")
    print("✅ Dados do tomador preenchidos.")

def preencher_dados_servico(driver, wait, dados, constantes):
    """Preenche os detalhes do serviço simulando a seleção da lista, mas apenas confirmando com Enter."""
    print("Preenchendo dados do serviço...")

    # --- Seleção do Dia ---
    input_dia = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//label[text()='Dia:']/following::input[1]")
    ))
    input_dia.click()          # abre o combo
    time.sleep(0.3)
    # Pressiona Enter para confirmar o dia selecionado
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(0.5)

    # --- Código de Serviço ---
    input_codigo = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[contains(@class,'x-combo-noedit') and contains(@class,'cbTextAlign')]")
    ))
    trigger = input_codigo.find_element(By.XPATH, "./following-sibling::img | ./following-sibling::div[contains(@class,'x-form-trigger')]")
    trigger.click()  # abre a lista
    time.sleep(0.3)

    # Digita o código e confirma com Enter
    
    ActionChains(driver).send_keys(Keys.ENTER).perform()

    # --- Preenchimento dos demais campos via TAB ---
    active_element = driver.switch_to.active_element
    active_element.send_keys(Keys.TAB)

    active_element = driver.switch_to.active_element
    active_element.send_keys(constantes['aliquota'])
    active_element.send_keys(Keys.TAB)

    active_element = driver.switch_to.active_element
    active_element.send_keys(constantes['descricao'])

    for _ in range(4):
        active_element = driver.switch_to.active_element
        active_element.send_keys(Keys.TAB)

    active_element = driver.switch_to.active_element
    active_element.send_keys(dados['valor_servico'])

    print("✅ Dados do serviço preenchidos com sucesso.")


def emitir_nfs(driver, wait):
    """Clica nos botões finais para emitir a NFS-e e confirma as caixas de diálogo."""
    print("Iniciando emissão da nota...")
    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Próximo Passo >> ')]")
    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Emitir')]")
    print("🔹 Botão 'Emitir' clicado.")

    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(1)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    
    print("🔹 Tecla ENTER pressionada 2x para confirmação.")

# --- Funções de Interação com a Página (Retry) ---
def clicar_com_retry(wait, by, identificador, tentativas=5):
    """Tenta clicar em um elemento de forma robusta, com várias tentativas."""
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, identificador)))
            elemento.click()
            return
        except (StaleElementReferenceException, TimeoutException):
            time.sleep(1)
    raise Exception(f"Não foi possível clicar no elemento {identificador} após {tentativas} tentativas.")

def preencher_por_label(wait, label_text, valor, tentativas=5):
    """Preenche um campo de input localizando-o a partir de seu label."""
    for _ in range(tentativas):
        try:
            label = wait.until(EC.presence_of_element_located((By.XPATH, f"//label[text()='{label_text}']")))
            input_id = label.get_attribute("for")
            input_element = wait.until(EC.visibility_of_element_located((By.ID, input_id)))
            input_element.clear()
            input_element.send_keys(valor)
            return
        except (StaleElementReferenceException, TimeoutException):
            time.sleep(1)
    raise Exception(f"Não foi possível preencher o campo com label '{label_text}' após {tentativas} tentativas.")

# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    main()
