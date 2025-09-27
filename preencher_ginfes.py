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
import pyautogui
import time
import os

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

        # --- Salvar PDF da Nota ---
        salvar_pdf(driver, wait, dados_variaveis)

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
    print("Preenchendo dados do serviço...")

   # ---------------- Seleção de Dia ----------------
    clicar_com_retry(wait, By.XPATH, "//label[text()='Dia:']/following::input[1]")  # abre o combo
    item_dia = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'x-combo-list-item') and text()='{dados['dia']}']"))
    )
    item_dia.click()
    time.sleep(0.3)


    # Código de Serviço
    input_codigo = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[contains(@class,'x-combo-noedit') and contains(@class,'cbTextAlign')]")
    ))
    trigger = input_codigo.find_element(By.XPATH, "./following-sibling::img | ./following-sibling::div[contains(@class,'x-form-trigger')]")
    trigger.click()
    time.sleep(0.3)

    ActionChains(driver).send_keys(Keys.ENTER).perform()

    # Preenche os demais campos
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
    print("Iniciando emissão da nota...")
    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Próximo Passo >> ')]")
    clicar_com_retry(wait, By.XPATH, "//button[contains(text(),'Emitir')]")
    print("🔹 Botão 'Emitir' clicado.")

    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(1)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    
    print("🔹 Tecla ENTER pressionada 2x para confirmação.")



def salvar_pdf(driver, wait, dados):
    print("Iniciando exportação do PDF...")

    # --- Caminho para salvar o PDF ---
    caminho_pasta = r"C:\Users\Colibri\Desktop\Notas\Notas Set"
    os.makedirs(caminho_pasta, exist_ok=True)

    # Gera o nome final do arquivo
    nome_arquivo = f"{dados['nome']} - {dados['cpf']}.pdf"
    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        nome_arquivo = nome_arquivo.replace(char, '-')
    caminho_pdf = os.path.join(caminho_pasta, nome_arquivo)

    try:
        # Garante que estamos na última aba
        driver.switch_to.window(driver.window_handles[-1])
        print(f"🔎 URL atual: {driver.current_url}")

        # Localiza e clica no botão "Exportar PDF"
        botao_pdf = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[@class='titulo' and contains(text(),'Exportar PDF')]")
        ))
        botao_pdf.click()
        print("✅ Botão 'Exportar PDF' clicado.")

        # Espera abrir a nova aba com o PDF
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])
        print("🔹 PDF carregado na nova aba.")

        # --- Salvar PDF usando pyautogui ---
        time.sleep(2)  # espera o PDF renderizar
        pyautogui.hotkey('ctrl', 's')  # abre a janela "Salvar como"
        time.sleep(1)
        pyautogui.typewrite(caminho_pdf)
        pyautogui.press('enter')
        time.sleep(3)
        print(f"✅ PDF salvo em: {caminho_pdf}")

    except Exception as e:
        print(f"❌ Erro ao exportar PDF: {e}")



# --- Funções auxiliares ---
def clicar_com_retry(wait, by, identificador, tentativas=5):
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, identificador)))
            elemento.click()
            return
        except (StaleElementReferenceException, TimeoutException):
            time.sleep(1)
    raise Exception(f"Não foi possível clicar no elemento {identificador} após {tentativas} tentativas.")

def preencher_por_label(wait, label_text, valor, tentativas=5):
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

# --- Ponto de Entrada ---
if __name__ == "__main__":
    main()
