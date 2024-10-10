from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
import time
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
import pandas as pd
import os

def dir_name_archive(nome_arquivo):
    #consegue o nome e o caminho do arquivo
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    caminho_completo = os.path.join(desktop, nome_arquivo)
    return caminho_completo

def get_previous_month_dates():
    # retorna o mês anterior, trazendo o primeiro e o ultimo dia dele
    today = datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    previous_month = last_day_of_previous_month.month
    previous_year = last_day_of_previous_month.year
    first_day_previous_month = f'01/{previous_month:02d}/{previous_year}'
    last_day_previous_month = f'{last_day_of_previous_month.day:02d}/{previous_month:02d}/{previous_year}'
    return first_day_previous_month, last_day_previous_month

def baixarArquivo():
    destination_dir = r"C:\Users\matheus.barbosa\OneDrive - Brasilata SA\CatracaData"
    edge_options = Options()
    
    prefs = {
        "download.default_directory": destination_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }

    edge_options.add_experimental_option("prefs", prefs)

    service = EdgeService()
    driver = webdriver.Edge(service=service, options=edge_options)
    
    driver.get("http://11.150.6.90:81/NewAccess/logon.aspx?ReturnUrl=%2fNewAccess%2fEquipments%2fEquipmentsLst.aspx")

    try:
        username_field = driver.find_element(By.ID, 'txtUsrLogin')
        password_field = driver.find_element(By.ID, 'txtUserPassLogin')
        username_field.send_keys('teste.rh')
        password_field.send_keys('1()vern0@2030')
        password_field.send_keys(Keys.RETURN)
        time.sleep(2)

        driver.get("http://11.150.6.90:81/NewAccess/Reports/ProcessPluginReport.aspx?idPlugin=4")
        time.sleep(3)

        first_day_previous_month, last_day_previous_month = get_previous_month_dates()

        campo_primeiro_dia = driver.find_element(By.ID, 'MainContentMainMaster_MainContent_ctl00_ctl18_txtDateFrom')
        campo_primeiro_dia.send_keys(first_day_previous_month)

        campo_ultimo_dia = driver.find_element(By.ID, 'MainContentMainMaster_MainContent_ctl00_ctl18_txtDateTo')
        campo_ultimo_dia.send_keys(last_day_previous_month)
        campo_ultimo_dia.send_keys(Keys.RETURN)
        
        dropdown = Select(driver.find_element(By.ID, "MainContentMainMaster_MainContent_ctl00_ddlGenerateType"))
        dropdown.select_by_value("4")

        okClick = driver.find_element(By.ID, 'MainContentMainMaster_MainContent_ctl00_btnGenerate')
        okClick.click()

        # Espera o download ser concluído
        timeout = 60  # Tempo máximo de espera em segundos
        start_time = time.time()
        while time.time() - start_time < timeout:
            time.sleep(5)  # Espera 5 segundos antes de verificar novamente
            list_of_files = os.listdir(destination_dir)
            full_path = [os.path.join(destination_dir, file) for file in list_of_files if file.endswith('.csv')]
            if full_path:
                latest_file = max(full_path, key=os.path.getctime)
                print(f"Arquivo mais recente: {latest_file}")
                return latest_file

        raise ValueError("Nenhum arquivo CSV foi encontrado no diretório de destino após o download.")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    try:
        # Baixa o arquivo
        arquivo_baixado = baixarArquivo()
        
        # Importa o CSV usando pandas
        df = pd.read_csv(arquivo_baixado)

        # Delete as 10 primeiras linhas
        df = df.drop(index=range(10))

        # Verifica se a coluna 'Data do Evento' existe
        if 'Data do Evento' in df.columns:
            # Separa a data e a hora
            df['Data'] = pd.to_datetime(df['Data do Evento']).dt.date
            df['Hora'] = pd.to_datetime(df['Data do Evento']).dt.time
            
            # Salva o DataFrame modificado de volta em um arquivo Excel
            modified_latest_file = arquivo_baixado.replace('.csv', '_modified.xlsx')
            df.to_excel(modified_latest_file, index=False)
            print(f"Modified file saved as: {modified_latest_file}")
        else:
            print("A coluna 'Data do Evento' não foi encontrada no arquivo CSV.")
            print(df.columns)
          
        print("As funções foram finalizadas.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")