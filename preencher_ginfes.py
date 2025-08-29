import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import time

# ---------------- Funções de retry ----------------
def clicar_com_retry(driver, by, identificador, tentativas=5):
    wait = WebDriverWait(driver, 10)
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, identificador)))
            elemento.click()
            return True
        except StaleElementReferenceException:
            time.sleep(1)
    raise Exception(f"Não foi possível clicar no elemento {identificador} após {tentativas} tentativas.")

def preencher_com_retry(driver, by, identificador, valor, tentativas=5):
    wait = WebDriverWait(driver, 10)
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.visibility_of_element_located((by, identificador)))
            elemento.clear()
            elemento.send_keys(valor)
            return True
        except StaleElementReferenceException:
            time.sleep(1)
    raise Exception(f"Não foi possível preencher o elemento {identificador} após {tentativas} tentativas.")

def preencher_por_label(driver, label_text, valor, tentativas=5):
    wait = WebDriverWait(driver, 10)
    for _ in range(tentativas):
        try:
            label = wait.until(EC.presence_of_element_located((By.XPATH, f"//label[text()='{label_text}']")))
            input_id = label.get_attribute("for")
            input_element = wait.until(EC.visibility_of_element_located((By.ID, input_id)))
            input_element.clear()
            input_element.send_keys(valor)
            return True
        except StaleElementReferenceException:
            time.sleep(1)
    raise Exception(f"Não foi possível preencher o campo com label '{label_text}' após {tentativas} tentativas.")

# ---------------- Coleta de dados via GUI ----------------
def coletar_dados_gui():
    root = tk.Tk()
    root.withdraw()
    dados = {}
    dados['cpf'] = simpledialog.askstring("Input", "Digite o CPF:")
    dados['nome'] = simpledialog.askstring("Input", "Digite o Nome:")
    dados['cep'] = simpledialog.askstring("Input", "Digite o CEP:")
    dados['numero'] = simpledialog.askstring("Input", "Digite o Número:")
    dados['telefone'] = simpledialog.askstring("Input", "Digite o Telefone:")
    dados['dia'] = simpledialog.askstring("Input", "Digite o Dia desejado (ex: 28):")
    dados['valor_servico'] = simpledialog.askstring("Input", "Digite o Valor do Serviço:")
    return dados

# ---------------- Função principal ----------------
def main():
    dados = coletar_dados_gui()

    codigo_servico = "8.02/859960100"
    descricao = "Serviços prestados"
    aliquota = "2"

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get("https://guarulhos.ginfes.com.br/")

    # ---------------- Login ----------------
    clicar_com_retry(driver, By.CLASS_NAME, "imagem1")
    time.sleep(2)
    preencher_com_retry(driver, By.ID, "ext-gen29", "24255199000167")
    preencher_com_retry(driver, By.ID, "ext-gen33", "colibri")
    clicar_com_retry(driver, By.ID, "ext-gen51")
    time.sleep(3)

    # ---------------- Inicia NFSe ----------------
    imagem_nfse = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//img[@class='gwt-Image' and contains(@src,'icon_nfse3.gif')]"))
    )
    imagem_nfse.click()
    time.sleep(2)

    # ---------------- Preenchimento via labels ----------------
    preencher_por_label(driver, "CPF:", dados['cpf'])
    preencher_por_label(driver, "Nome:", dados['nome'])
    preencher_por_label(driver, "CEP:", dados['cep'])

    clicar_com_retry(driver, By.XPATH, "//img[@title='Busca endereço pelo cep']")
    time.sleep(1)

    preencher_por_label(driver, "Número:", dados['numero'])
    preencher_por_label(driver, "Telefone:", dados['telefone'])

    clicar_com_retry(driver, By.XPATH, "//button[contains(text(),'Próximo Passo')]")
    time.sleep(2)

    # ---------------- Seleção de Dia ----------------
    clicar_com_retry(driver, By.XPATH, "//label[text()='Dia:']/following::input[1]")
    item_dia = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//div[contains(@class,'x-combo-list-item') and text()='{dados['dia']}']"))
    )
    item_dia.click()

    # ---------------- Seleção de Código de Serviço ----------------
    input_combo = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//input[contains(@class,'x-combo-noedit') and contains(@class,'cbTextAlign')]"))
)
    trigger = input_combo.find_element(By.XPATH, "./following-sibling::img | ./following-sibling::div[contains(@class,'x-form-trigger')]")
    trigger.click()

    item_servico = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, f"//div[contains(@class,'x-combo-list-item') and text()='{codigo_servico}']"))
)
    item_servico.click()

    time.sleep(10)

    # ---------------- Descrição e Valor ----------------
# ---------------- Preencher Descrição ----------------
    textarea_field = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//textarea[@class='x-form-textarea x-form-field']"))
)
    textarea_field.clear()
    textarea_field.send_keys(descricao)

    preencher_por_label(driver, "Valor do Serviço:", dados['valor_servico'])
    print("✅ Valor preenchidos")

    # ---------------- Finaliza e envia ----------------
    clicar_com_retry(driver, By.XPATH, "//button[contains(text(),'Próximo Passo')]")
    print("✅ Formulário preenchido com sucesso (verifique antes de enviar).")



# ---------------- Execução ----------------
if __name__ == "__main__":
    main()
