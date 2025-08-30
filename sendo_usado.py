# Imports originais
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import os
import glob
import tkinter as tk
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
import pyperclip
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import sys
import schedule
import time
from tkinter import messagebox


# Pega o diretório raiz do projeto (onde o script está rodando)
project_root = os.getcwd()

# Cria uma pasta 'downloads' dentro da raiz, se quiser organizar assim
download_dir = os.path.join(project_root, "downloads_Inova")

# Cria a pasta se não existir
os.makedirs(download_dir, exist_ok=True)

# Configurações e variáveis globais
usuario_padrao = "u04846"
senha_padrao = "@Inova260909"
URL = 'https://glpi10.pmjlle.joinville.sc.gov.br/index.php?noAUTO=1'

# Configura opções do Chrome
options = Options()
options.add_argument("--disable-sync")
options.add_argument("--disable-notifications")
options.add_argument("--disable-default-apps")
options.add_argument("--no-first-run")
options.add_argument("--no-default-browser-check")
prefs = {
    "profile.default_content_setting_values.notifications": 2,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# A variável navegador é definida dentro da função de automação
# para garantir que uma nova instância seja criada a cada execução agendada.
navegador = None

def copiar_csv_para_area_transferencia():
    arquivos_csv = glob.glob(os.path.join(download_dir, '*.csv'))
    if not arquivos_csv:
        print("Nenhum arquivo CSV encontrado.")
        return

    arquivo_csv = max(arquivos_csv, key=os.path.getctime)
    print(f"Usando CSV: {arquivo_csv}")
    df = pd.read_csv(arquivo_csv, encoding='utf-8', sep=';')
    df['ultima_att_planilha'] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    texto = df.to_csv(sep='\t', index=False)
    pyperclip.copy(texto)
    print("📋 Dados copiados")

def abrir_planilha_e_colar(link_planilha):
    global navegador
    print("🔗 Abrindo planilha...")
    navegador.get(link_planilha)
    WebDriverWait(navegador, 60).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
    sleep(10)
    print("📋 Colando dados na planilha...")
    actions = ActionChains(navegador)
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    print("✅ Dados colados.")
    sleep(15)
    print("Fechando Navegador...")
    navegador.quit()
    print("✅ Automação concluída para este horário.")

def esperar_download(download_folder, timeout=120):
    segundos_passados = 0
    while segundos_passados < timeout:
        arquivos_parciais = glob.glob(os.path.join(download_folder, '*.crdownload'))
        if arquivos_parciais:
            sleep(1)
            segundos_passados += 1
        else:
            return True
    return False

def iniciar_navegador():
    global navegador
    navegador = webdriver.Chrome(options=options)

def login(usuario, senha):
    global navegador
    navegador.get(URL)
    navegador.maximize_window()
    sleep(1)
    print('🔐 Fazendo Login...')
    navegador.find_element(By.XPATH, '//*[@id="login_name"]').send_keys(usuario + Keys.TAB)
    sleep(0.5)
    navegador.find_element(By.XPATH, '//*[@id="login_password"]').send_keys(senha + Keys.TAB)
    sleep(0.5)
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/div[1]/div/form/div/div[1]/div[6]/button').click()

def entrando():
    global navegador
    print('➡️ Navegando para página dos dados...')
    WebDriverWait(navegador, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/aside')))
    sleep(2)
    navegador.find_element(By.XPATH, '/html/body/div[2]/aside/div/div[2]/ul/li[2]/a').click()
    sleep(5)
    navegador.find_element(By.XPATH, '/html/body/div[2]/aside/div/div[2]/ul/li[2]/div/div/div/a[2]').click()
    sleep(10)

def apagar_csv():
    pasta = download_dir
    arquivos_csv = glob.glob(os.path.join(pasta, "*.csv"))
    for arquivo in arquivos_csv:
        try:
            os.remove(arquivo)
            print(f"Arquivo removido: {arquivo}")
        except Exception as e:
            print(f"Erro ao remover {arquivo}: {e}")

def fazer_download():
    global navegador
    botao_abrir_download = navegador.find_element(By.XPATH, '//*[@id="dropdown-export"]/span')
    botao_abrir_download.click()
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="massformTicket"]/div/div[1]/div/div/ul/li[8]/a').click()
    print('Aguarde Enquanto está sendo feito o Download!')
    if esperar_download(download_dir):
        print("Download finalizado.")
        sleep(5)
        copiar_csv_para_area_transferencia()
        abrir_planilha_e_colar("https://docs.google.com/spreadsheets/d/1CSPaLSCYeov30wmoiekIc-7MJKrTtedFd9WJGUtVKls/edit?gid=1305019307#gid=1305019307")
    else:
        print("Tempo de espera excedido, download pode não ter terminado.")
        navegador.quit()

def executar_automacao(usuario, senha):
    """
    Função que executa uma rodada completa da automação.
    """
    print("="*50)
    print(f"🚀 Iniciando execução agendada em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("="*50)
    try:
        iniciar_navegador()
        login(usuario, senha)
        entrando()
        apagar_csv()
        fazer_download()
    except Exception as e:
        print(f"❌ Ocorreu um erro durante a automação: {e}")
        if navegador:
            navegador.quit()

class AgendadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Agendador de Tarefas")

        self.dias_vars = {
            "segunda": tk.BooleanVar(),
            "terca": tk.BooleanVar(),
            "quarta": tk.BooleanVar(),
            "quinta": tk.BooleanVar(),
            "sexta": tk.BooleanVar(),
            "sabado": tk.BooleanVar(),
            "domingo": tk.BooleanVar()
        }

        frame_dias = tk.LabelFrame(root, text="1. Escolha os dias da semana", padx=10, pady=10)
        frame_dias.pack(padx=15, pady=10, fill="x")

        tk.Checkbutton(frame_dias, text="Segunda", variable=self.dias_vars["segunda"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Terça", variable=self.dias_vars["terca"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Quarta", variable=self.dias_vars["quarta"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Quinta", variable=self.dias_vars["quinta"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Sexta", variable=self.dias_vars["sexta"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Sábado", variable=self.dias_vars["sabado"]).pack(anchor="w")
        tk.Checkbutton(frame_dias, text="Domingo", variable=self.dias_vars["domingo"]).pack(anchor="w")

        frame_horarios = tk.LabelFrame(root, text="2. Configure os horários", padx=10, pady=10)
        frame_horarios.pack(padx=15, pady=10, fill="x")

        tk.Label(frame_horarios, text="Horários (separados por vírgula, formato HH:MM):").pack(anchor="w")
        self.horarios_entry = tk.Entry(frame_horarios, width=50)
        self.horarios_entry.pack(pady=5)
        self.horarios_entry.insert(0, "06:00, 12:00, 18:00")

        self.btn_iniciar = tk.Button(root, text="Agendar e Iniciar", command=self.iniciar_agendamento, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.btn_iniciar.pack(pady=15, padx=15, fill="x")
        
        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.pack(pady=(0,10))

    def iniciar_agendamento(self):
        dias_selecionados = [dia for dia, var in self.dias_vars.items() if var.get()]
        horarios_str = self.horarios_entry.get()

        if not dias_selecionados:
            messagebox.showerror("Erro", "Selecione pelo menos um dia da semana.")
            return

        if not horarios_str:
            messagebox.showerror("Erro", "Informe pelo menos um horário.")
            return

        horarios = [h.strip() for h in horarios_str.split(',')]
        
        for h in horarios:
            if len(h) != 5 or h[2] != ':':
                messagebox.showerror("Erro de Formato", f"O horário '{h}' parece estar em um formato inválido. Use HH:MM.")
                return

        self.root.destroy()
        
        print("✔ Configuração recebida. Iniciando agendador...")
        print(f"Dias selecionados: {', '.join(dias_selecionados)}")
        print(f"Horários configurados: {', '.join(horarios)}")
        
        agendar_e_rodar(dias_selecionados, horarios)


def agendar_e_rodar(dias, horarios):
    mapa_dias = {
        "segunda": "monday",
        "terca": "tuesday",
        "quarta": "wednesday",
        "quinta": "thursday",
        "sexta": "friday",
        "sabado": "saturday",
        "domingo": "sunday"
    }

    for dia_str in dias:
        dia_em_ingles = mapa_dias[dia_str]
        
        for horario in horarios:
            print(f"✔️ Agendando tarefa para toda {dia_str}-feira às {horario}")
            
            agendador = getattr(schedule.every(), dia_em_ingles)
            agendador.at(horario).do(executar_automacao, usuario=usuario_padrao, senha=senha_padrao)

    print("\n==========================================================")
    print("✨ AGENDADOR ATIVO. O terminal ficará em execução.")
    print("   Pressione CTRL+C para encerrar o processo.")
    print("==========================================================")

    while True:
        schedule.run_pending()
        time.sleep(1)


# Ponto de entrada do script
if __name__ == "__main__":
    root = tk.Tk()
    app = AgendadorGUI(root)
    root.mainloop()