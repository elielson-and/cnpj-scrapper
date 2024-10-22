import requests
import cloudscraper
import time
import pandas as pd
import random
from bs4 import BeautifulSoup
from tkinter import Tk, Label, Button, StringVar, OptionMenu, Checkbutton, BooleanVar, Text, Scrollbar, END, messagebox
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
import os
from datetime import datetime
import sys
import colorama
from colorama import Fore, Style
import unicodedata  # Para remover acentos
import threading

colorama.init(autoreset=True)  # Inicia a biblioteca colorama para cores no terminal

# Variável de controle para testar a mensagem de timeout
showTestMessage = False  # Coloque True para testar a mensagem

# Função para exibir uma mensagem pop-up se o timeout for atingido
def show_timeout_message():
    messagebox.showwarning(
        "Aviso de Timeout",
        "O servidor parou de enviar informações. Por favor, tente realizar solicitações com intervalos menores de datas e/ou uma menor quantidade de municípios no filtro."
    )

# URL da API
url_api = 'https://api.casadosdados.com.br/v2/public/cnpj/search'

# Função para obter estados via API do IBGE
def obter_estados():
    url_estados = 'https://servicodados.ibge.gov.br/api/v1/localidades/estados'
    response = requests.get(url_estados)
    if response.status_code == 200:
        estados = response.json()
        return [(estado['sigla'], estado['nome']) for estado in estados]
    else:
        print_terminal(f"Erro ao obter estados: {response.status_code}")
        return []

# Função para obter municípios de um estado via API do IBGE
def obter_municipios(sigla_estado):
    url_municipios = f'https://servicodados.ibge.gov.br/api/v1/localidades/estados/{sigla_estado}/municipios'
    response = requests.get(url_municipios)
    if response.status_code == 200:
        municipios = response.json()
        return [municipio['nome'] for municipio in municipios]
    else:
        print_terminal(f"Erro ao obter municípios: {response.status_code}")
        return []

# Função para remover acentos e converter texto para maiúsculas
def remover_acentos(texto):
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return ''.join([char for char in nfkd_form if not unicodedata.combining(char)]).upper()

# Função para carregar municípios após selecionar o estado
def carregar_municipios():
    sigla_estado = uf_var.get()  # Obtém a sigla do estado selecionado
    if sigla_estado:
        municipios = obter_municipios(sigla_estado)
        if municipios:
            municipio_dropdown['values'] = municipios  # Popula o combobox com os municípios
            municipio_dropdown.set('')  # Limpa a seleção anterior
            municipio_opcional1_dropdown['values'] = municipios  # Popula os municípios adicionais
            municipio_opcional1_dropdown.set('')
            municipio_opcional2_dropdown['values'] = municipios
            municipio_opcional2_dropdown.set('')
            print_terminal(f"Municípios carregados para o estado: {sigla_estado}")
        else:
            print_terminal(f"Erro ao carregar os municípios para o estado {sigla_estado}.")
    else:
        print_terminal("Selecione um estado válido.")

# Função para gerar headers aleatórios
def generate_headers():
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/54.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.1 Safari/605.1.15',
        'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'
    ]
    headers = {
        'User-Agent': random.choice(user_agents),
        'Content-Type': 'application/json'
    }
    return headers

# Função para fazer o scraping (etapa 2)
def scrape_additional_data(razao_social, cnpj):
    url = f'https://casadosdados.com.br/solucao/cnpj/{razao_social.lower().replace(" ", "-")}-{cnpj}#google_vignette'
    scraper = cloudscraper.create_scraper()
    
    response = scraper.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        dados = {}
        sections = soup.find_all('div', class_='p-3')
        
        for section in sections:
            label = section.find('label')
            value_p = section.find('p')
            value_a = section.find('a')
            
            if label:
                label_text = label.text.strip().replace(":", "")
                if value_p:
                    value_text = value_p.text.strip()
                elif value_a:
                    value_text = value_a.text.strip()
                else:
                    value_text = 'Valor não encontrado'
                dados[label_text] = value_text
        return dados
    else:
        print_terminal(f"Erro ao acessar a página: {response.status_code}")
        return {}

# Função para buscar dados da API (etapa 1) e concatenar com dados de scraping (etapa 2)
def fetch_all_pages(payload):
    all_data = []
    page = 1  # Iniciar com a primeira página
    scraper = cloudscraper.create_scraper()

    timeout_timer = threading.Timer(20.0, show_timeout_message)  # Iniciar o timer de 20 segundos para timeout
    if not showTestMessage:  # Não inicie o scraping se showTestMessage for True
        timeout_timer.start()

    while True:
        payload['page'] = page  # Atualiza a página no payload
        
        # Gera headers aleatórios para cada requisição
        headers = generate_headers()
        os.system("clear")
        
        # Tempo de espera aleatório entre 4 e 10 segundos
        wait_time = random.randint(1, 3)
        
        print_terminal(f'>> PRÓXIMA REQUISIÇÃO EM {wait_time} SEGUNDOS...')
        time.sleep(wait_time)

        # Faz a requisição POST com headers aleatórios
        response = scraper.post(url_api, json=payload, headers=headers)

        # Verifica se a requisição foi bem-sucedida
        if response.status_code == 200:
            result = response.json()
            
            # Verifica se a resposta contém dados e sucesso
            if result.get('success', False) and result['data'].get('cnpj'):
                for empresa in result['data']['cnpj']:
                    # Extraímos o nome (razao_social) e o CNPJ da empresa
                    razao_social = empresa.get('razao_social', '').replace(".", "").replace(",", "")
                    cnpj = empresa.get('cnpj', '')

                    # Fazemos o scraping na etapa 2 para obter os dados adicionais
                    print_terminal(f"Processando: {razao_social}, CNPJ: {cnpj}")
                    time.sleep(1)
                    additional_data = scrape_additional_data(razao_social, cnpj)

                    # Unir os dados da API (etapa 1) com os dados do scraping (etapa 2)
                    empresa.update(additional_data)
                    all_data.append(empresa)

                # Avança para a próxima página
                page += 1
            else:
                print_terminal('\nMAXIMO DE INFORMACOES COLETADAS!\n')
                break  # Parar se não houver mais dados ou erro na resposta
        else:
            print_terminal(f'Erro {response.status_code}: {response.text}')
            break

    timeout_timer.cancel()  # Cancela o timeout se a requisição finalizar com sucesso

    return all_data

# Função para converter a data do formato brasileiro para o formato americano
def converter_data_brasileira_para_americana(data_br):
    dia, mes, ano = data_br.split('/')
    return f"{ano}-{mes}-{dia}"

# Função para imprimir mensagens no "terminal" da interface
def print_terminal(message):
    terminal_output.config(state='normal')
    terminal_output.insert(END, message + '\n')
    terminal_output.config(state='disabled')
    terminal_output.see(END)  # Rola para a última linha

# Função que será chamada ao clicar no botão de "Obter Dados"
def obter_dados():
    if showTestMessage:
        show_timeout_message()
        return

    print_terminal("\n\nCONECTANDO AO SERVIDOR...\n\n")
   
    data_abertura_br = data_abertura_var.get()  # Pega a data no formato brasileiro
    data_abertura_americana = converter_data_brasileira_para_americana(data_abertura_br)  # Converte a data para o formato americano

    # Verificar se o campo 'Aberto até' foi preenchido (opcional)
    data_ate_americana = None
    if data_ate_var.get():
        data_ate_br = data_ate_var.get()
        data_ate_americana = converter_data_brasileira_para_americana(data_ate_br)

    # Se o município não for selecionado, passar um array vazio
    municipio_selecionado = [remover_acentos(municipio_var.get())] if municipio_var.get() else []
    municipio_opcional1 = [remover_acentos(municipio_opcional1_var.get())] if municipio_opcional1_var.get() else []
    municipio_opcional2 = [remover_acentos(municipio_opcional2_var.get())] if municipio_opcional2_var.get() else []

    # Monta o payload com gte, lte e os três municípios
    municipios_para_busca = municipio_selecionado + municipio_opcional1 + municipio_opcional2

    payload = {
        "query": {
            "termo": [],
            "atividade_principal": [],
            "natureza_juridica": [],
            "uf": [uf_var.get()],
            "municipio": municipios_para_busca,  # Inclui os três municípios (principal e opcionais)
            "bairro": [],
            "situacao_cadastral": situacao_cadastral_var.get(),
            "cep": [],
            "ddd": []
        },
        "range_query": {
            "data_abertura": {
                "gte": data_abertura_americana,  # Usa a data de início
                "lte": data_ate_americana  # Usa a data de término se fornecida
            },
            "capital_social": {
                "lte": '99999999',
                "gte": '100'
            }
        },
        "extras": {
            "somente_mei": somente_mei_var.get(),
            "excluir_mei": excluir_mei_var.get(),
            "com_email": False,
            "incluir_atividade_secundaria": False,
            "com_contato_telefonico": False,
            "somente_fixo": False,
            "somente_celular": False,
            "somente_matriz": False,
            "somente_filial": False
        },
        "page": 1
    }
    
    # Chama a função de busca e scraping em uma thread separada
    threading.Thread(target=processar_dados, args=(payload,)).start()

# Função para processar os dados e gerar o Excel
def processar_dados(payload):
    empresas = fetch_all_pages(payload)
    
    # Cria a pasta 'resultados' se não existir
    if not os.path.exists('resultados'):
        os.makedirs('resultados')

    # Limpeza e formatação dos dados:
    df_empresas = pd.DataFrame(empresas)

    # 1. Remover o texto "Whatsapp" da coluna 'Telefone'
    if 'Telefone' in df_empresas.columns:
        df_empresas['Telefone'] = df_empresas['Telefone'].str.replace('Whatsapp', '').str.strip()

    # 2. Formatando a coluna 'data_abertura' para o formato brasileiro (dd/mm/yyyy)
    if 'data_abertura' in df_empresas.columns:
        df_empresas['data_abertura'] = pd.to_datetime(df_empresas['data_abertura']).dt.strftime('%d/%m/%Y')

    # Salva o arquivo Excel na pasta 'resultados'
    data_atual = datetime.now().strftime("%Y-%m-%d")
    arquivo_excel = f'resultados/empresas_completas_{data_atual}.xlsx'
    
    df_empresas.to_excel(arquivo_excel, index=False)
    os.system("clear")
    print_terminal(f'\n\n=================\n[OPERACAO FINALIZADA] Dados salvos no arquivo {arquivo_excel}\n\nVOCE PODE FECHAR ESTE TERMINAL A PARTIR DE AGORA\n=================\n\n')

# Interface gráfica (Tkinter)
root = Tk()
root.title("Consulta de Empresas")

# Obter estados
estados = obter_estados()

# Labels e dropdowns
Label(root, text="UF").grid(row=0, column=0)
uf_var = StringVar(root)
uf_var.set("Selecione")
uf_dropdown = OptionMenu(root, uf_var, *[estado[0] for estado in estados])
uf_dropdown.grid(row=0, column=1)

# Botão para carregar municípios
Button(root, text="Carregar Municípios", command=carregar_municipios).grid(row=0, column=2)

Label(root, text="Município (Obrigatório)").grid(row=1, column=0)
municipio_var = StringVar(root)
municipio_dropdown = Combobox(root, textvariable=municipio_var)
municipio_dropdown.grid(row=1, column=1)

Label(root, text="Município Opcional 1").grid(row=2, column=0)
municipio_opcional1_var = StringVar(root)
municipio_opcional1_dropdown = Combobox(root, textvariable=municipio_opcional1_var)
municipio_opcional1_dropdown.grid(row=2, column=1)

Label(root, text="Município Opcional 2").grid(row=3, column=0)
municipio_opcional2_var = StringVar(root)
municipio_opcional2_dropdown = Combobox(root, textvariable=municipio_opcional2_var)
municipio_opcional2_dropdown.grid(row=3, column=1)

Label(root, text="Situação Cadastral").grid(row=4, column=0)
situacao_cadastral_var = StringVar(root)
situacao_cadastral_dropdown = OptionMenu(root, situacao_cadastral_var, "ATIVA", "BAIXADA", "SUSPENSA", "NULA")
situacao_cadastral_var.set("ATIVA")
situacao_cadastral_dropdown.grid(row=4, column=1)

Label(root, text="Aberto de (Data Inicial)").grid(row=5, column=0)
data_abertura_var = StringVar()
data_abertura_entry = DateEntry(root, textvariable=data_abertura_var, date_pattern='dd/mm/yyyy')
data_abertura_entry.grid(row=5, column=1)

Label(root, text="Aberto até (Data Final - Opcional)").grid(row=6, column=0)
data_ate_var = StringVar()
data_ate_entry = DateEntry(root, textvariable=data_ate_var, date_pattern='dd/mm/yyyy')
data_ate_entry.grid(row=6, column=1)

somente_mei_var = BooleanVar()
Checkbutton(root, text="Somente MEI", variable=somente_mei_var).grid(row=7, column=0)

excluir_mei_var = BooleanVar()
Checkbutton(root, text="Excluir MEI", variable=excluir_mei_var).grid(row=7, column=1)

# Botão para iniciar a consulta
Button(root, text="Obter Dados", command=obter_dados).grid(row=8, column=0, columnspan=2)

# Adicionar o widget de "terminal" na interface
Label(root, text="Terminal Output").grid(row=9, column=0, columnspan=2)
terminal_output = Text(root, height=10, state='disabled')
terminal_output.grid(row=10, column=0, columnspan=2)

# Scrollbar para o "terminal"
scrollbar = Scrollbar(root, command=terminal_output.yview)
terminal_output['yscrollcommand'] = scrollbar.set
scrollbar.grid(row=10, column=2, sticky='ns')

Label(root, text="telegram: @httptrampos").grid(row=11, column=0)

root.mainloop()
