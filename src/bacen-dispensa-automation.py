# -*- coding: utf-8 -*-
"""
Created on Mon May 11 15:05:26 2026

@author: ov0006
"""

import os
import re
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC




chrome_driver_path = r"C:\Users\chromedriver.exe"

caminho_planilha = "samples/dispensa_input.xlsx"

saida_resultado = "outputs/dispensa_results.xlsx"

url_dispensa = "https://scr.bcb.gov.br/scr/pesquisar/entidade/DenDispensaEnvio"

pasta_screenshots = "outputs/screenshots"
os.makedirs(pasta_screenshots, exist_ok=True)




def limpar_cnpj(valor):
    if pd.isna(valor):
        return ""
    return re.sub(r"\D", "", str(valor)).zfill(8)


def normalizar_database(valor):

    if pd.isna(valor):
        return "", ""

    if isinstance(valor, pd.Timestamp):
        texto = valor.strftime("%m/%Y")
        value = valor.strftime("%Y%m")
        return texto, value

    valor = str(valor).strip()

    if re.fullmatch(r"\d{6}", valor):
        value = valor
        texto = f"{valor[4:6]}/{valor[:4]}"
        return texto, value

    if re.fullmatch(r"\d{4}-\d{2}", valor):
        ano, mes = valor.split("-")
        texto = f"{mes}/{ano}"
        value = f"{ano}{mes}"
        return texto, value

    if re.fullmatch(r"\d{2}/\d{4}", valor):
        mes, ano = valor.split("/")
        texto = f"{mes}/{ano}"
        value = f"{ano}{mes}"
        return texto, value

    raise ValueError(f"Data-base inválida: {valor}")


def clicar_com_js_se_precisar(driver, elemento):
    try:
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)


def selecionar_option_por_cnpj(select_element, cnpj_raiz):
    select = Select(select_element)

    for option in select.options:
        texto = option.text.strip()
        value = option.get_attribute("value").strip()

        if cnpj_raiz in re.sub(r"\D", "", texto) or cnpj_raiz == value:
            select.select_by_value(value)
            return True, texto

    return False, ""


def selecionar_database(select_element, texto_mm_aaaa, value_aaaamm):
    select = Select(select_element)

    for option in select.options:
        texto = option.text.strip()
        value = option.get_attribute("value").strip()

        if texto == texto_mm_aaaa or value == value_aaaamm:
            select.select_by_value(value)
            return True

    return False


def localizar_select_database(driver, label_texto):
    """
    Localiza o select associado ao label da Data-Base Início/Fim.
    Evita depender de ID dinâmico do Wicket.
    """
    xpath = (
        f"//label[contains(normalize-space(), '{label_texto}')]"
        "/ancestor::td/following-sibling::td[1]//select"
    )

    return WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )




if not os.path.exists(chrome_driver_path):
    raise FileNotFoundError(f"ChromeDriver não encontrado: {chrome_driver_path}")

df = pd.read_excel(caminho_planilha)

colunas_obrigatorias = [
    "Instituição Financeira",
    "Data-Base Início",
    "Data-Base Fim"
]

for coluna in colunas_obrigatorias:
    if coluna not in df.columns:
        raise ValueError(f"Coluna obrigatória não encontrada na planilha: {coluna}")



options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")

service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)

resultados = []

try:
    driver.get(url_dispensa)

    input("Faça o login manualmente no SCR. Depois, pressione ENTER aqui para continuar...")

    for index, row in df.iterrows():
        cnpj_raiz = limpar_cnpj(row["Instituição Financeira"])

        db_inicio_texto, db_inicio_value = normalizar_database(row["Data-Base Início"])
        db_fim_texto, db_fim_value = normalizar_database(row["Data-Base Fim"])

        print(f"\nProcessando linha {index + 1}: {cnpj_raiz} | {db_inicio_texto} até {db_fim_texto}")

        status = ""
        detalhe = ""

        try:
            driver.get(url_dispensa)

            botao_incluir = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='Incluir']"))
            )
            clicar_com_js_se_precisar(driver, botao_incluir)

            select_if = wait.until(
                EC.presence_of_element_located((By.ID, "instituicaoFinanceira_6"))
            )

            encontrou_if, texto_if = selecionar_option_por_cnpj(select_if, cnpj_raiz)

            if not encontrou_if:
                status = "Erro"
                detalhe = "Instituição Financeira não encontrada na lista suspensa"
                print(detalhe)
                resultados.append({
                    "Instituição Financeira": cnpj_raiz,
                    "Data-Base Início": db_inicio_texto,
                    "Data-Base Fim": db_fim_texto,
                    "Status": status,
                    "Detalhe": detalhe
                })
                continue

            select_inicio = localizar_select_database(driver, "Data-Base Início")
            ok_inicio = selecionar_database(select_inicio, db_inicio_texto, db_inicio_value)

            if not ok_inicio:
                raise ValueError(f"Data-Base Início não encontrada na lista: {db_inicio_texto}")

            select_fim = localizar_select_database(driver, "Data-Base Fim")
            ok_fim = selecionar_database(select_fim, db_fim_texto, db_fim_value)

            if not ok_fim:
                raise ValueError(f"Data-Base Fim não encontrada na lista: {db_fim_texto}")



            botao_salvar = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@type='submit' and @value='Salvar']")
                )
            )
            
            driver.execute_script("arguments[0].scrollIntoView(true);", botao_salvar)
            
            time.sleep(1)
            
            driver.execute_script("arguments[0].click();", botao_salvar)




            time.sleep(2)
            
            botao_confirmar = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@type='submit' and @value='Confirmar']")
                )
            )
            
            driver.execute_script(
                "arguments[0].scrollIntoView(true);",
                botao_confirmar
            )
            
            time.sleep(1)
            
            driver.execute_script(
                "arguments[0].click();",
                botao_confirmar
            )
            
            time.sleep(2)

            
            

            time.sleep(2)

            status = "Processado"
            detalhe = f"Dispensa salva para {texto_if}"
            print(detalhe)

        except Exception as e:
            status = "Erro"
            detalhe = str(e)

            nome_print = f"erro_linha_{index + 1}_{cnpj_raiz}.png"
            caminho_print = os.path.join(pasta_screenshots, nome_print)
            driver.save_screenshot(caminho_print)

            print(f"Erro na linha {index + 1}: {detalhe}")
            print(f"Screenshot salvo em: {caminho_print}")

        resultados.append({
            "Instituição Financeira": cnpj_raiz,
            "Data-Base Início": db_inicio_texto,
            "Data-Base Fim": db_fim_texto,
            "Status": status,
            "Detalhe": detalhe
        })

finally:
    pd.DataFrame(resultados).to_excel(saida_resultado, index=False)
    print(f"\nResultado salvo em: {saida_resultado}")
    driver.quit()