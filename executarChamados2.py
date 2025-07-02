import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import threading
import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image, ImageTk
import tempfile
import shutil
import os

class App:
    def __init__(self, root):
        self.root = root
        root.title("Automação SUA - Abrir Chamados")
        root.geometry("700x740")

        # Logo
        try:
            caminho_logo = os.path.join(os.path.dirname(__file__), "logotipoDETI.png")
            img = Image.open(caminho_logo)
            img.thumbnail((300, 160))
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(root, image=self.logo_img).pack(pady=10)
        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")

        self.login_var = tk.StringVar(value="felipe.pinto")
        self.senha_var = tk.StringVar(value="felypekratos2")
        self.arquivo_path = tk.StringVar()
        self.navegador_var = tk.StringVar(value="Edge")

        tk.Label(root, text="Login SUA:").pack(pady=(10, 0))
        tk.Entry(root, textvariable=self.login_var, width=40).pack()

        tk.Label(root, text="Senha SUA:").pack(pady=(10, 0))
        tk.Entry(root, textvariable=self.senha_var, width=40, show="*").pack()

        tk.Label(root, text="Arquivo Excel de chamados:").pack(pady=(10, 0))
        tk.Entry(root, textvariable=self.arquivo_path, width=80).pack(padx=10)
        tk.Button(root, text="Selecionar arquivo", command=self.selecionar_arquivo).pack(pady=5)

        tk.Label(root, text="Escolha o navegador:").pack(pady=(10, 0))
        tk.OptionMenu(root, self.navegador_var, "Edge", "Chrome").pack()

        self.btn_iniciar = tk.Button(root, text="Iniciar Abertura de Chamados", command=self.iniciar_processamento)
        self.btn_iniciar.pack(pady=10)

        tk.Label(root, text="Progresso:").pack()
        self.progress = ttk.Progressbar(root, length=600, mode='determinate')
        self.progress.pack(pady=5)

        tk.Label(root, text="Log de execução:").pack()
        self.log_text = scrolledtext.ScrolledText(root, height=15, width=85, state='disabled')
        self.log_text.pack(padx=10, pady=5)

        # Informações de versão e desenvolvedor
        tk.Label(root, text="Versão 1.0 | Desenvolvedor: Felipe Souza - DETI 2025", font=("Arial", 9, "italic")).pack(pady=(10, 5))

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if caminho:
            self.arquivo_path.set(caminho)
            self.log(f"Arquivo selecionado: {caminho}")

    def iniciar_processamento(self):
        if not self.login_var.get().strip():
            messagebox.showerror("Erro", "Por favor, insira o login.")
            return
        if not self.senha_var.get().strip():
            messagebox.showerror("Erro", "Por favor, insira a senha.")
            return
        if not self.arquivo_path.get():
            messagebox.showerror("Erro", "Por favor, selecione um arquivo Excel.")
            return

        self.btn_iniciar.config(state='disabled')
        self.progress['value'] = 0
        thread = threading.Thread(target=self.processar_chamados)
        thread.start()

    def processar_chamados(self):
        navegador = None
        user_data_dir = None
        try:
            navegador_escolhido = self.navegador_var.get()

            user_data_dir = tempfile.mkdtemp(prefix="selenium_profile_")

            if navegador_escolhido == "Edge":
                options = EdgeOptions()
                options.add_argument("--start-maximized")
                options.add_argument(f"--user-data-dir={user_data_dir}")
                service = EdgeService(EdgeChromiumDriverManager().install())
                navegador = webdriver.Edge(service=service, options=options)

            elif navegador_escolhido == "Chrome":
                options = ChromeOptions()
                options.add_argument("--start-maximized")
                options.add_argument(f"--user-data-dir={user_data_dir}")
                service = ChromeService(ChromeDriverManager().install())
                navegador = webdriver.Chrome(service=service, options=options)

            self.log(f"Navegador iniciado: {navegador_escolhido}")

            navegador.get("https://sua.riobranco.ac.gov.br/index.php?noAUTO=1")
            WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'login_name')))
            navegador.find_element(By.ID, 'login_name').send_keys(self.login_var.get())
            navegador.find_element(By.ID, 'login_password').send_keys(self.senha_var.get())
            navegador.find_element(By.XPATH, '//*[@id="boxlogin"]/form/p[5]/input').click()
            time.sleep(2)

            df = pd.read_excel(self.arquivo_path.get())
            total = len(df)
            self.progress['maximum'] = total

            for index, row in df.iterrows():
                titulo = row['TITULO']
                descricao = row['DESCRICAO']
                descricao = re.sub(r"(?<!^)\s*(Problema:|Unidade:|Patrimônio:|Modelo:|Local:)", r"\n\1", str(descricao).strip())
                hora = row['DATA E HORA']
                categoria = row['CATEGORIA']
                atribuido = row['ATRIBUIDO']
                localizacao = row['LOCALIZACAO']
                unidade = row['UNIDADE']

                if pd.isna(titulo) or pd.isna(descricao) or pd.isna(hora) or pd.isna(categoria) or pd.isna(atribuido):
                    self.log(f"[{index+1}] Linha ignorada: campos obrigatórios vazios.")
                    self.progress['value'] += 1
                    continue

                try:
                    navegador.get("https://sua.riobranco.ac.gov.br/front/ticket.form.php")
                    time.sleep(3)

                    navegador.find_element(By.XPATH, '//*[@id="mainformtable4"]/tbody/tr[1]/td/input').send_keys(titulo + ' - ' + unidade)
                    time.sleep(2)

                    iframe = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//iframe[contains(@id,"ifr")]')))
                    navegador.switch_to.frame(iframe)
                    editor = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'tinymce')))
                    editor.clear()
                    editor.send_keys(descricao)
                    navegador.switch_to.default_content()

                    def preencher_dropdown(xpath_click, valor):
                        campo = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_click)))
                        campo.click()
                        inp = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]')))
                        inp.send_keys(valor)
                        time.sleep(1)
                        inp.send_keys(Keys.RETURN)

                    preencher_dropdown('//span[contains(@id,"select2-dropdown_itilcategories_id")]', categoria)
                    preencher_dropdown('//span[contains(@id,"select2-dropdown__users_id_assign")]', atribuido)
                    preencher_dropdown('//span[contains(@id,"select2-dropdown_locations_id")]', localizacao)

                    cal = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//a[@class="input-button"]')))
                    cal.click()

                    if not isinstance(hora, str):
                        hora_str = hora.strftime("%d/%m/%Y %H:%M:%S")
                    else:
                        hora_str = hora

                    campo = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//input[contains(@class,"flatpickr") and contains(@class, "active")]'))
                    )
                    campo.clear()
                    campo.send_keys(hora_str)
                    campo.send_keys(Keys.RETURN)

                    try:
                        WebDriverWait(navegador, 3).until(
                            EC.invisibility_of_element_located((By.XPATH, '//div[contains(@class,"flatpickr-calendar")]'))
                        )
                    except:
                        navegador.find_element(By.XPATH, '//body').click()

                    botao_enviar = WebDriverWait(navegador, 5).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="tabsbody"]/div[2]/button'))
                    )
                    botao_enviar.click()

                    self.log(f"[{index+1}] Chamado enviado com sucesso.")
                    time.sleep(1)

                except Exception as e:
                    self.log(f"[{index+1}] Erro no chamado: {e}")

                self.progress['value'] = index + 1

            self.log("Processamento finalizado.")
            messagebox.showinfo("Fim", "Processamento finalizado com sucesso!")

        except Exception as e:
            self.log(f"Erro geral: {e}")
            messagebox.showerror("Erro", f"Erro geral: {e}")
        finally:
            self.btn_iniciar.config(state='normal')
            if navegador:
                try:
                    navegador.quit()
                except:
                    pass
            if user_data_dir and os.path.exists(user_data_dir):
                try:
                    shutil.rmtree(user_data_dir)
                except Exception as e:
                    self.log(f"Erro ao apagar diretório temporário: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
