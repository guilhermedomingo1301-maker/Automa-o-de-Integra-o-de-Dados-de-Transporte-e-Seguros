from datetime import datetime
import glob
import os
import time
import zipfile
from dotenv import load_dotenv
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# CONFIGURAÇÃO DE DATAS E VARIÁVEIS DE AMBIENTE

hoje = datetime.now()
data_inicial = hoje.strftime("%d/%m/%Y")
data_final = hoje.strftime("%d/%m/%Y")
data_solicitacao = hoje.strftime("%d/%m/%Y")

DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
arquivo_historico = os.path.join(DIRETORIO_SCRIPT, "protocolos_enviados.txt")

load_dotenv()
usuario_email = os.getenv("user", "")
usuario_senha = os.getenv("senha", "")
url_script = os.getenv("url_script_googlesheet", "")


# AUTOMAÇÃO SELENIUM

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(options=options)

try:
    driver.get("https://postonlineplataforma.atmtec.com.br/php/home.php")
    driver.maximize_window()
    aguardar = WebDriverWait(driver, 10)

    campo_usuario = aguardar.until(
        EC.presence_of_element_located((By.ID, "email"))
    )
    campo_usuario.clear()
    campo_usuario.send_keys(usuario_email)

    campo_senha = aguardar.until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    campo_senha.clear()
    campo_senha.send_keys(usuario_senha)
    campo_senha.send_keys(Keys.RETURN)
    time.sleep(3)

    relatorio_botão = driver.find_element(By.ID, "LinkAbas-2")
    relatorio_botão.click()
    time.sleep(5)

    solicitar_botão = driver.find_element(By.ID, "BtnShowRel")
    solicitar_botão.click()
    time.sleep(5)

    relatorio_excel = driver.find_element(By.ID, "Rel016")
    relatorio_excel.click()

    grupos_botao = driver.find_element(By.ID, "lbl_grupos")
    grupos_botao.click()
    time.sleep(5)

    combo_grup = aguardar.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#comboIGGroup input.ui-igcombo-field")
        )
    )
    combo_grup.click()
    combo_grup.clear()
    combo_grup.send_keys("TRANSPORTE NORMAL")
    time.sleep(5)

    data_inicio = driver.find_element(By.ID, "RelDtInicial")
    data_inicio.click()
    data_inicio.send_keys(data_inicial)

    data_fim = driver.find_element(By.ID, "RelDtFinal")
    data_fim.click()
    data_fim.send_keys(data_final)
    data_fim.send_keys(Keys.RETURN)
    time.sleep(2)

    solicitar_rel_bot = driver.find_element(By.ID, "SolistRel")
    solicitar_rel_bot.click()
    time.sleep(5)

    bot_atualizar = driver.find_element(By.ID, "BtnAtzGrid")
    bot_atualizar.click()

    time.sleep(180)  # Espera a geração do relatório

    bot_atualizar = driver.find_element(By.ID, "BtnAtzGrid")
    bot_atualizar.click()
    time.sleep(5)

    linha_relatorio_hoje = aguardar.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//tr[td[contains(text(), '{data_solicitacao}')] and"
            " td[text()='Sim']][1]",
        ))
    )
    linha_relatorio_hoje.click()
    time.sleep(5)

    btn_baixar = driver.find_element(By.ID, "BtnDownRel")
    btn_baixar.click()
    time.sleep(5)

    remove_rel = driver.find_element(By.ID, "SelTodos")
    remove_rel.click()
    time.sleep(3)
    remove_rel_sim = driver.find_element(By.ID, "jqi_state0_buttonSim")
    remove_rel_sim.click()
    time.sleep(5)

finally:
    driver.quit()


# PROCESSAMENTO DO ARQUIVO ZIP

pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
pasta_processados = os.path.join(pasta_downloads, "zips_processados")
os.makedirs(pasta_processados, exist_ok=True)

arquivos_zip = glob.glob(os.path.join(pasta_downloads, "*.zip"))
zips_hoje = [
    f
    for f in arquivos_zip
    if datetime.fromtimestamp(os.path.getmtime(f)).date() == hoje.date()
]

if not zips_hoje:
    print("Aviso: Nenhum arquivo .zip novo foi encontrado na pasta Downloads.")
    zip_de_hoje = None
else:
    zip_de_hoje = max(zips_hoje, key=os.path.getmtime)
    print(f"ZIP identificado para processamento: {zip_de_hoje}")

if zip_de_hoje:
    pasta_extracao = os.path.join(pasta_downloads, "temp_relatorios")
    os.makedirs(pasta_extracao, exist_ok=True)

    with zipfile.ZipFile(zip_de_hoje, "r") as zip_ref:
        zip_ref.extractall(pasta_extracao)

    arquivos_extraidos = glob.glob(os.path.join(pasta_extracao, "*.*"))
    arquivo_dados = max(arquivos_extraidos, key=os.path.getmtime)

    if arquivo_dados.endswith(".csv"):
        df_temp = pd.read_csv(arquivo_dados, header=None)
    else:
        df_temp = pd.read_excel(arquivo_dados, header=None)

    linha_cabecalho = None
    for idx, row in df_temp.iterrows():
        if any(
                "Caixa Postal" in str(val) or "Protocolo de averbação" in str(val)
                for val in row.tolist()
        ):
            linha_cabecalho = idx
            break

    header_index = linha_cabecalho if linha_cabecalho is not None else 3
    if arquivo_dados.endswith(".csv"):
        df_hoje = pd.read_csv(arquivo_dados, header=header_index)
    else:
        df_hoje = pd.read_excel(arquivo_dados, header=header_index)

    df_hoje.columns = [str(col).strip() for col in df_hoje.columns]
    df_hoje = df_hoje.dropna(how="all")

    if not df_hoje.empty and len(df_hoje.columns) > 0:
        primeira_coluna_nome = df_hoje.columns[0]
        df_hoje = df_hoje[
            df_hoje[primeira_coluna_nome].astype(str).str.strip()
            != str(primeira_coluna_nome).strip()
            ]
        df_hoje = df_hoje[
            df_hoje[primeira_coluna_nome].astype(str).str.strip() != "Caixa Postal"
            ]

    colunas_que_eu_quero = [
        "Caixa Postal",
        "Razão Cliente",
        "CNPJ Segurado",
        "Série",
        "Número",
        "Tipo de Documento",
        "Ramo",
        "Origem",
        "Destino",
        "Valor Total",
        "Placa",
        "Emissão",
        "Mun Origem Desc",
        "Mun Destino Desc",
        "Protocolo de averbação",
    ]

    for col in colunas_que_eu_quero:
        if col not in df_hoje.columns:
            df_hoje[col] = ""
    df_hoje = df_hoje[colunas_que_eu_quero]

    # TRATAMENTO ANTI-DUPLICIDADE E GESTÃO DE ALTERAÇÕES (EX: 0,01)

    coluna_protocolo = "Protocolo de averbação"
    coluna_valor = "Valor Total"


    def limpar_valor_monetario(valor):
        if pd.isna(valor) or valor == "":
            return 0.0

        if isinstance(valor, (int, float)):
            return float(valor)

        v = str(valor).replace("R$", "").strip()

        if "," in v:
            v = v.replace(".", "")
            v = v.replace(",", ".")

        try:
            return float(v)
        except ValueError:
            return 0.0


    if coluna_protocolo in df_hoje.columns:

        df_hoje[coluna_protocolo] = (
            df_hoje[coluna_protocolo]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )


        if coluna_valor in df_hoje.columns:
            df_hoje[coluna_valor] = df_hoje[coluna_valor].apply(limpar_valor_monetario)


        df_hoje = df_hoje.drop_duplicates(subset=[coluna_protocolo], keep="last")

    df_hoje = df_hoje.fillna("")


    # ATUALIZAÇÃO SEGURA DO POWER BI (Arquivo Local Consolidado)

    caminho_power_bi = os.path.join(
        DIRETORIO_SCRIPT, "relatorio_atualizado.xlsx"
    )
    df_hoje.to_excel(caminho_power_bi, index=False)
    print(
        f"Base consolidada salva para o Power BI (sem duplicatas):"
        f" {caminho_power_bi}"
    )

    nome_arquivo_zip = os.path.basename(zip_de_hoje)
    caminho_destino = os.path.join(pasta_processados, nome_arquivo_zip)
    if os.path.exists(caminho_destino):
        os.remove(caminho_destino)
    os.rename(zip_de_hoje, caminho_destino)


    # FILTRAGEM INTELIGENTE: APENAS INÉDITOS OU MODIFICADOS PARA O SHEETS

    protocolos_ja_enviados = set()
    if os.path.exists(arquivo_historico):
        with open(arquivo_historico, "r", encoding="utf-8") as f:
            protocolos_ja_enviados = set(line.strip() for line in f if line.strip())


    df_novos = df_hoje[~df_hoje[coluna_protocolo].isin(protocolos_ja_enviados)]
    df_modificados = df_hoje[df_hoje[coluna_protocolo].isin(protocolos_ja_enviados)]


    df_cancelados = df_modificados[df_modificados[coluna_valor] == 0.01]

    print(
        f"Novos protocolos inéditos para adicionar: {len(df_novos)} | Protocolos"
        f" já existentes com alteração para 0,01: {len(df_cancelados)}"
    )


    # ENVIO PARA O GOOGLE SHEETS (ADICIONAR NOVOS E SUBSCREVER OS ERRADOS)

    URL_WEB_APP_GOOGLE = (url_script)


    if not df_novos.empty:
        linhas_novas = df_novos.values.tolist()
        TAMANHO_LOTE = 500

        for i in range(0, len(linhas_novas), TAMANHO_LOTE):
            lote = linhas_novas[i: i + TAMANHO_LOTE]
            protocolos_lote = df_novos[coluna_protocolo].iloc[
                i: i + TAMANHO_LOTE
            ].tolist()

            payload = {"opcao": "anexar", "rows": lote}
            try:
                response = requests.post(URL_WEB_APP_GOOGLE, json=payload, timeout=60)
                if response.status_code == 200:
                    print(f"Lote novo enviado com sucesso! ({len(lote)} linhas)")
                    with open(arquivo_historico, "a", encoding="utf-8") as f:
                        for p in protocolos_lote:
                            if str(p).strip():
                                f.write(f"{str(p).strip()}\n")
            except Exception as e:
                print(f"Erro ao enviar lote novo: {e}")
            time.sleep(1)


    if not df_cancelados.empty:
        linhas_canceladas = df_cancelados.values.tolist()
        payload_atualizacao = {"opcao": "atualizar_linhas", "rows": linhas_canceladas}

        try:
            print(
                f"Enviando solicitação para subscrever/corrigir"
                f" {len(linhas_canceladas)} linhas alteradas para 0,01..."
            )
            response = requests.post(
                URL_WEB_APP_GOOGLE, json=payload_atualizacao, timeout=60
            )
            if response.status_code == 200:
                print("Linhas canceladas/alteradas atualizadas com sucesso na planilha!")
        except Exception as e:
            print(f"Erro ao atualizar linhas canceladas: {e}")


    # LIMPEZA DE DUPLICATAS ANTIGAS NA NUVEM

    print("Enviando comando para o Google Sheets varrer e limpar duplicatas antigas...")
    payload_limpeza = {"opcao": "remover_duplicatas"}
    try:
        resp_limpeza = requests.post(URL_WEB_APP_GOOGLE, json=payload_limpeza, timeout=60)
        if resp_limpeza.status_code == 200:
            print("Limpeza automática de duplicatas concluída com sucesso na nuvem!")
    except Exception as e:
        print(f"Erro ao solicitar limpeza de duplicatas: {e}")

    print("Processo de automação e validação concluído com sucesso!")
else:
    print("Nenhum dado encontrado para processamento.")