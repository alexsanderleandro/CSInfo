import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
import csinfo
from csinfo import main as csinfo_main
import importlib.util, sys
import subprocess


def resource_path(rel_path: str) -> str:
    """Resolve um caminho relativo ao diretório do script ou ao bundle do PyInstaller."""
    if getattr(sys, 'frozen', False):
        # PyInstaller cria um temp folder e coloca o path em _MEIPASS
        base = sys._MEIPASS
    else:
        base = os.path.dirname(__file__)
    return os.path.join(base, rel_path)

# --------- Função para configurar estilos visuais ---------
def estilo_botoes():
    style = ttk.Style()
    style.theme_use('default')
    style.configure('Rounded.TButton',
                   font=('Segoe UI', 10, 'bold'),
                   padding=8,
                   borderwidth=0,
                   relief='flat',
                   foreground='#fff',
                   background='#1976D2',
                   focusthickness=3,
                   focuscolor='none')
    style.map('Rounded.TButton',
              background=[('active', '#1565C0'), ('disabled', '#b0b0b0')])
    # Botão Exportar - verde
    style.configure('Export.TButton',
                   font=('Segoe UI', 10, 'bold'),
                   padding=8,
                   borderwidth=0,
                   relief='flat',
                   foreground='#fff',
                   background='#43A047')
    style.map('Export.TButton',
              background=[('active', '#388E3C'), ('disabled', '#b0b0b0')])
    # Botão Sair - vermelho
    style.configure('Exit.TButton',
                   font=('Segoe UI', 10, 'bold'),
                   padding=8,
                   borderwidth=0,
                   relief='flat',
                   foreground='#fff',
                   background='#E53935')
    style.map('Exit.TButton',
              background=[('active', '#B71C1C'), ('disabled', '#b0b0b0')])
    # Radiobutton circular
    style.configure('Custom.TRadiobutton',
                   indicatorcolor='#1976D2',
                   indicatordiameter=14,
                   indicatorsize=14,
                   font=('Segoe UI', 10),
                   padding=4)
    style.map('Custom.TRadiobutton',
              indicatorcolor=[('selected', '#1976D2'), ('active', '#1565C0')])


# --------- Classe principal adaptada ---------
class CSInfoApp(tk.Tk):
    def bloquear_fechar(self):
        pass  # Ignora o evento de fechar
    def desbloquear_fechar(self):
        self.protocol("WM_DELETE_WINDOW", self.quit)

    def __init__(self):
        super().__init__()
        # Incluir versão do pacote no título
        try:
            ver = getattr(csinfo, '__version__', None)
            if ver:
                self.title(f"CSInfo – Inventário de Hardware e Software  v{ver}")
            else:
                self.title("CSInfo – Inventário de Hardware e Software")
        except Exception:
            self.title("CSInfo – Inventário de Hardware e Software")
        self.geometry("900x650")
        self.configure(bg='#f4f6f9')
        self.resizable(False, False)
        estilo_botoes()
        # Tenta configurar o ícone do aplicativo (assets/ico.png)
        try:
            self.set_app_icon()
        except Exception:
            pass
        self.criar_layout()

    def set_app_icon(self):
        """Configura o ícone da aplicação. Usa .ico no Windows ou PNG via Pillow/tk.PhotoImage como fallback."""
        # Caminho relativo ao projeto
        png_path = resource_path(os.path.join('assets', 'ico.png'))
        ico_path = resource_path(os.path.join('assets', 'app.ico'))
        # Se existir um .ico preferimos usá-lo no Windows
        try:
            if os.path.exists(ico_path) and os.name == 'nt':
                self.iconbitmap(ico_path)
                return
        except Exception:
            pass

        # Se não existir .ico mas houver o PNG, tentar gerar app.ico automaticamente (Windows)
        if os.path.exists(png_path):
            # Tentar criar o .ico a partir do PNG usando Pillow, se possível
            # Regenerar o .ico se ele não existir ou se o PNG for mais novo
            regenerate_ico = False
            try:
                if not os.path.exists(ico_path):
                    regenerate_ico = True
                else:
                    if os.path.exists(png_path) and os.path.getmtime(png_path) > os.path.getmtime(ico_path):
                        regenerate_ico = True
            except Exception:
                regenerate_ico = regenerate_ico or False

            if regenerate_ico and os.name == 'nt':
                try:
                    from PIL import Image
                    # Garantir diretório de destino
                    ico_dir = os.path.dirname(ico_path)
                    if ico_dir and not os.path.exists(ico_dir):
                        os.makedirs(ico_dir, exist_ok=True)
                    img = Image.open(png_path)
                    # Converter para RGBA caso necessário
                    if img.mode not in ('RGBA', 'RGB'):
                        img = img.convert('RGBA')
                    # Salvar ícone com múltiplas resoluções para melhor compatibilidade no Windows
                    sizes = [(16, 16), (32, 32), (48, 48), (256, 256)]
                    img.save(ico_path, format='ICO', sizes=sizes)
                except Exception:
                    # Se falhar, apenas seguimos para os fallbacks
                    pass

            # Tenta usar Pillow + ImageTk para aplicar o PNG como iconphoto (cross-platform)
            try:
                from PIL import Image, ImageTk
                img = Image.open(png_path)
                photo = ImageTk.PhotoImage(img)
                # wm_iconphoto aplica ao ícone da janela em muitas plataformas
                self.wm_iconphoto(True, photo)
                # manter referência para evitar garbage collection
                self._icon_photo = photo
                # Se geramos o .ico e estamos no Windows, também aplicá-lo como iconbitmap
                try:
                    if os.path.exists(ico_path) and os.name == 'nt':
                        self.iconbitmap(ico_path)
                except Exception:
                    pass
                return
            except Exception:
                pass

        # Fallback: tentar usar PhotoImage nativo do Tkinter (suporta PNG em builds recentes)
        try:
            if os.path.exists(png_path):
                photo = tk.PhotoImage(file=png_path)
                self.wm_iconphoto(True, photo)
                self._icon_photo = photo
                return
        except Exception:
            pass

        # Último recurso: se existir o app.ico e estivermos no Windows, tentar aplicá-lo
        try:
            if os.path.exists(ico_path) and os.name == 'nt':
                self.iconbitmap(ico_path)
                return
        except Exception:
            pass

    def criar_layout(self):
        # Menu removido (Ajuda/Sobre) a pedido — nada a configurar aqui
        # Título
        self.label = tk.Label(self, text="CSInfo – Inventário de Hardware e Software", font=("Segoe UI", 16, "bold"), fg="#003366", bg="#f4f6f9")
        self.label.pack(pady=(20, 10))

        # Frame de entrada
        frame_entrada = tk.Frame(self, bg='#f4f6f9')
        frame_entrada.pack(pady=(10, 5))
        label_machine = tk.Label(frame_entrada, text="Nome da máquina:", font=("Segoe UI", 10, 'normal'), fg='#333', bg='#f4f6f9')
        label_machine.grid(row=0, column=0, padx=(0, 8), sticky='e')

        # Campo de input seguido pelo botão Iniciar (botão à direita)
        self.machine_var = tk.StringVar()
        self.machine_entry = tk.Entry(frame_entrada, textvariable=self.machine_var, font=("Segoe UI", 10), width=40, relief='solid', justify='center')
        self.machine_entry.grid(row=0, column=1, sticky='w')
        # Permitir que a coluna do input cresça se necessário
        frame_entrada.grid_columnconfigure(1, weight=1)

        # Botão Iniciar posicionado à direita do campo de input
        self.start_btn = ttk.Button(frame_entrada, text="▶ Iniciar", style='Rounded.TButton', command=self.start_process)
        self.start_btn.grid(row=0, column=2, padx=(8, 0))

        self.placeholder = "Deixe em branco para máquina local"
        self.machine_entry.insert(0, self.placeholder)
        self.machine_entry.config(fg='#888', font=('Segoe UI', 9, 'italic'), justify='center')
        def limitar_input_e_maiusculo(*args):
            valor = self.machine_var.get()
            valor_maiusculo = valor.upper()
            if len(valor_maiusculo) > 40:
                valor_maiusculo = valor_maiusculo[:40]
            if valor != valor_maiusculo:
                self.machine_var.set(valor_maiusculo)
            elif len(valor_maiusculo) > 40:
                self.machine_var.set(valor_maiusculo)
        self.machine_var.trace_add('write', limitar_input_e_maiusculo)
        def on_focus_in(event):
            if self.machine_entry.get() == self.placeholder:
                self.machine_entry.delete(0, tk.END)
                self.machine_entry.config(fg='#222', font=('Segoe UI', 10), justify='center')
        def on_focus_out(event):
            if not self.machine_entry.get():
                self.machine_entry.insert(0, self.placeholder)
                self.machine_entry.config(fg='#888', font=('Segoe UI', 9, 'italic'), justify='center')
                self.machine_var.set("")  # Garante que a variável não fique com o placeholder
            elif self.machine_entry.get() == self.placeholder:
                self.machine_var.set("")  # Garante que a variável não fique com o placeholder
        self.machine_entry.bind("<FocusIn>", on_focus_in)
        self.machine_entry.bind("<FocusOut>", on_focus_out)

    # Removido label 'Gerado por' do formulário — fica apenas nos relatórios
    # ...existing code...

    # ...o botão Iniciar agora fica dentro do frame_entrada ao lado do input...

        # Área de resultados
        frame_result = tk.Frame(self, bg='#f4f6f9')
        frame_result.pack(pady=(5, 10))
        self.info_text = ScrolledText(frame_result, width=110, height=20, font=('Consolas', 9),
                                      bg='white', fg='#222', relief='solid', borderwidth=1)
        self.info_text.pack()
        self.info_text.config(state=tk.DISABLED)  # Sempre inicia como somente leitura
        # Remove qualquer binding que permita digitação
        self.info_text.bind('<Key>', lambda e: 'break')
        self.info_text.bind('<Button-1>', lambda e: 'break')

        # Barra de progresso / mensagem
        self.progress_label = tk.Label(self, text="", font=('Segoe UI', 10, 'bold'), fg='#1976D2', bg='#f4f6f9')
        self.progress_label.pack(pady=(5, 0))

        # Progressbar visual (0% - 100%)
        self.progressbar = ttk.Progressbar(self, orient='horizontal', length=600, mode='determinate')
        self.progressbar.pack(pady=(4, 6))
        self.progress_percent_label = tk.Label(self, text="", font=('Segoe UI', 9), fg='#333', bg='#f4f6f9')
        self.progress_percent_label.pack()

        # Rodapé: Exportação
        frame_export = tk.Frame(self, bg='#f4f6f9')
        frame_export.pack(pady=(15, 5))
        label_export = tk.Label(frame_export, text="Exportar resultado:", font=('Segoe UI', 10, 'normal'), fg='#333', bg='#f4f6f9')
        label_export.grid(row=0, column=0, padx=(0, 8))
        self.export_var = tk.StringVar(value='txt')
        self.radio_txt = tk.Radiobutton(frame_export, text='TXT', variable=self.export_var, value='txt', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_pdf = tk.Radiobutton(frame_export, text='PDF', variable=self.export_var, value='pdf', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_ambos = tk.Radiobutton(frame_export, text='Ambos', variable=self.export_var, value='ambos', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_txt.grid(row=0, column=1, padx=5)
        self.radio_pdf.grid(row=0, column=2, padx=5)
        self.radio_ambos.grid(row=0, column=3, padx=5)

        # Botões Exportar e Sair
        frame_botoes = tk.Frame(self, bg='#f4f6f9')
        frame_botoes.pack(pady=(10, 20))
        self.export_btn = ttk.Button(frame_botoes, text="💾 Exportar", style='Export.TButton', command=self.exportar)
        self.export_btn.grid(row=0, column=0, padx=10)
        self.exit_btn = ttk.Button(frame_botoes, text="✖ Sair", style='Exit.TButton', command=self.quit)
        self.exit_btn.grid(row=0, column=1, padx=10)

        # Inicialmente, desabilita botões de exportação
        self.export_btn.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)

        # Rodapé centralizado com versão
        # Rodapé sem versão do app (apenas o nome da empresa)
        rodape_text = "CEOsoftware Sistemas"
        # Rodapé centralizado
        rodape = tk.Label(self, text=rodape_text, font=("Segoe UI", 8), fg="#666", bg="#f4f6f9")
        rodape.pack(side=tk.BOTTOM, pady=(0, 6))
        rodape.configure(anchor="center", justify="center")

    def show_about(self):
        try:
            ver = getattr(csinfo, '__version__', 'desconhecida')
            info = f"CSInfo — versão {ver}\n\nGerado por CEOsoftware Sistemas"
        except Exception:
            info = "CSInfo — versão desconhecida"
        messagebox.showinfo("Sobre CSInfo", info)

    # ...métodos start_process, run_csinfo, exportar permanecem iguais, apenas adaptando para os novos widgets...

    def start_process(self):
        self.protocol("WM_DELETE_WINDOW", self.bloquear_fechar)  # Bloqueia o botão X
        self.config(cursor="wait")
        self.info_text.config(cursor="wait", state=tk.DISABLED)
        self.start_btn.config(state=tk.DISABLED)
        self.machine_entry.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.exit_btn.config(state=tk.DISABLED)
        self.progress_label.config(text="", fg="black")
        self.info_text.delete(1.0, tk.END)
        # Reseta barra de progresso antes da nova análise
        try:
            self.progressbar['value'] = 0
            self.progress_percent_label.config(text='0%')
        except Exception:
            pass
        threading.Thread(target=self.run_csinfo, daemon=True).start()

    def run_csinfo(self):
        print('DEBUG: Entrou em run_csinfo')
        try:
            import platform
            valor_input = self.machine_var.get().strip()
            print('DEBUG: valor_input:', valor_input)
            if not valor_input or valor_input == self.placeholder:
                machine_name = None
            elif valor_input.upper() == platform.node().upper():
                machine_name = None
            else:
                machine_name = valor_input
            print('DEBUG: machine_name:', machine_name)
            if machine_name:
                self.progress_label.config(text=f"Acessando a máquina {machine_name}...", fg="black", font=("Helvetica", 10, "bold"))
                self.update_idletasks()
            from csinfo import check_remote_machine
            print('DEBUG: Importou check_remote_machine')
            if machine_name and not check_remote_machine(machine_name):
                self.progress_label.config(text=f"Não foi possível acessar a máquina '{machine_name}'. Ela pode estar desligada, fora da rede ou sem WinRM ativado.")
                self.start_btn.config(state=tk.NORMAL)
                self.machine_entry.config(state=tk.NORMAL)
                self.radio_txt.config(state=tk.DISABLED)
                self.radio_pdf.config(state=tk.DISABLED)
                self.radio_ambos.config(state=tk.DISABLED)
                self.export_btn.config(state=tk.DISABLED)
                self.exit_btn.config(state=tk.NORMAL)
                self.desbloquear_fechar()  # Reabilita o botão X
                print('DEBUG: Falha ao acessar máquina remota')
                return
            self.capturado = []
            self.info_text.config(state=tk.NORMAL)
            self.info_text.delete(1.0, tk.END)
            def gui_callback(percent, etapa_texto=None):
                # callback que recebe (percent, etapa_texto) vindos de csinfo.barra_progresso
                def atualizar_ui():
                    try:
                        if etapa_texto:
                            self.progress_label.config(text=etapa_texto, font=("Helvetica", 10, "bold"))
                        # Se o primeiro argumento for percentual, atualiza a barra
                        try:
                            perc = int(percent)
                            self.progressbar['value'] = perc
                            self.progress_percent_label.config(text=f"{perc}%")
                        except Exception:
                            pass
                    except Exception:
                        pass
                # Atualiza UI de forma segura na thread principal
                self.after(0, atualizar_ui)
            print('DEBUG: Chamando csinfo_main')
            # Apenas coleta os dados na análise; não gerar arquivos automaticamente
            resultado = csinfo_main(export_type=None, barra_callback=gui_callback, computer_name=machine_name)
            # Armazena resultado da análise para possível exportação posterior
            self.last_result = resultado
            print('DEBUG: Resultado csinfo_main:', resultado)
            self.progress_label.config(text="Análise finalizada!", fg="red", font=("Helvetica", 10, "bold"))
            # Ao final, preferir mostrar o relatório em memória (lines) em vez de abrir arquivo no disco
            lines = resultado.get('lines') or []
            if lines:
                conteudo = "\n".join(lines)
                self.info_text.config(state=tk.NORMAL)  # Habilita só para atualizar
                self.info_text.delete(1.0, tk.END)
                self.info_text.insert(tk.END, conteudo)
                self.info_text.see(tk.END)
                self.info_text.config(state=tk.DISABLED)  # Volta para somente leitura
                # Atualiza barra de progresso para 100% ao final
                try:
                    self.progressbar['value'] = 100
                    self.progress_percent_label.config(text='100%')
                except Exception:
                    pass
            else:
                # Fallback: se não houver linhas, tentar ler txt_path (caso exista)
                txt_path = resultado.get('txt')
                print('DEBUG: txt_path fallback:', txt_path)
                if txt_path and os.path.exists(txt_path):
                    self.info_text.config(state=tk.NORMAL)
                    self.info_text.delete(1.0, tk.END)
                    with open(txt_path, encoding='utf-8') as f:
                        conteudo = f.read()
                    self.info_text.insert(tk.END, conteudo)
                    self.info_text.config(state=tk.DISABLED)
                else:
                    self.info_text.config(state=tk.NORMAL)
                    self.info_text.delete(1.0, tk.END)
                    self.info_text.insert(tk.END, "Relatório não encontrado.")
                    self.info_text.config(state=tk.DISABLED)
            self.export_btn.config(state=tk.NORMAL)
            self.radio_txt.config(state=tk.NORMAL)
            self.radio_pdf.config(state=tk.NORMAL)
            self.radio_ambos.config(state=tk.NORMAL)
        except Exception as e:
            print('DEBUG: Exceção geral em run_csinfo:', e)
            messagebox.showerror("Erro", str(e))
            self.export_btn.config(state=tk.DISABLED)
        finally:
            self.config(cursor="arrow")
            self.info_text.config(cursor="arrow", state=tk.NORMAL)
            self.start_btn.config(state=tk.NORMAL)
            self.machine_entry.config(state=tk.NORMAL)
            self.radio_txt.config(state=tk.NORMAL)
            self.radio_pdf.config(state=tk.NORMAL)
            self.radio_ambos.config(state=tk.NORMAL)
            self.exit_btn.config(state=tk.NORMAL)
            self.desbloquear_fechar()  # Reabilita o botão X

    def update_progress(self, value, etapa_texto=None):
        def atualizar():
            if etapa_texto:
                self.progress_label.config(text=etapa_texto)
            self.update_idletasks()
        self.after(0, atualizar)

    def exportar(self):
        # Reusa resultado da análise se disponível para evitar reexecução
        # self.last_result é preenchido após run_csinfo
        self.protocol("WM_DELETE_WINDOW", self.bloquear_fechar)  # Bloqueia o botão X
        self.config(cursor="wait")
        self.info_text.config(cursor="wait", state=tk.DISABLED)
        self.start_btn.config(state=tk.DISABLED)
        self.machine_entry.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.exit_btn.config(state=tk.DISABLED)
        self.progress_label.config(text="Exportando o relatório, aguarde...", font=("Helvetica", 10, "bold"))
        self.update_idletasks()
        try:
            print('DEBUG: Entrou no bloco try da exportação')
            machine_name = self.get_computer_name_input()
            tipo = self.export_var.get()
            msg = []
            # Tentar reutilizar resultado da análise
            reused = False
            if hasattr(self, 'last_result') and self.last_result:
                last_machine = self.last_result.get('machine')
                # Comparar nomes (None ou igual ignorando case)
                if (not machine_name and (not last_machine or last_machine.lower() == csinfo.get_machine_name(None).lower())) or (last_machine and machine_name and last_machine.lower() == machine_name.lower()):
                    lines = self.last_result.get('lines') or []
                    safe_name = csinfo.safe_filename(last_machine or csinfo.get_machine_name(None))
                    base_path = os.path.join(os.getcwd(), f"Info_maquina_{safe_name}.txt")
                    if tipo in ('txt', 'ambos'):
                        try:
                            csinfo.write_report(base_path, lines)
                            msg.append(f"Arquivo TXT exportado: {base_path}")
                            reused = True
                        except Exception as e:
                            print('DEBUG: Erro ao reusar write_report:', e)
                    if tipo in ('pdf', 'ambos'):
                        pdf_path = base_path.replace('.txt', '.pdf')
                        try:
                            ok = csinfo.write_pdf_report(pdf_path, lines, last_machine or csinfo.get_machine_name(None))
                            if ok:
                                msg.append(f"Arquivo PDF exportado: {pdf_path}")
                                reused = True
                        except Exception as e:
                            print('DEBUG: Erro ao reusar write_pdf_report:', e)

            # Se não reutilizamos (nenhum resultado disponível ou máquina diferente), executar export padrão
            if not reused:
                resultado = csinfo_main(export_type=tipo, barra_callback=None, computer_name=machine_name)
                print('DEBUG: Resultado da exportação (execução):', resultado)
                if resultado.get('txt'):
                    msg.append(f"Arquivo TXT exportado: {resultado['txt']}")
                if resultado.get('pdf'):
                    msg.append(f"Arquivo PDF exportado: {resultado['pdf']}")
            messagebox.showinfo("Exportação", "\n".join(msg) if msg else "Nenhum arquivo exportado.")
            caminho = None
            if msg:
                # pegar o primeiro caminho reportado na mensagem
                primeiros = [m.split(': ',1)[1] for m in msg if ': ' in m]
                caminho = primeiros[0] if primeiros else None
            if caminho:
                print('DEBUG: Caminho do arquivo exportado:', caminho)
                abrir = messagebox.askyesno("Abrir pasta", "Deseja abrir o diretório onde o relatório foi salvo?")
                if abrir:
                    print('DEBUG: Tentando acessar os.path.dirname')
                    pasta = os.path.dirname(caminho)
                    print('DEBUG: Pasta:', pasta)
                    try:
                        subprocess.Popen(f'explorer "{pasta}"')
                    except Exception as e:
                        print('DEBUG: Erro ao abrir pasta:', e)
                        messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")
        except Exception as e:
            print('DEBUG: Exceção geral na exportação:', e)
            messagebox.showerror("Erro na exportação", str(e))
        finally:
            self.config(cursor="arrow")
            self.info_text.config(cursor="arrow", state=tk.NORMAL)
            self.progress_label.config(text="")
            self.start_btn.config(state=tk.NORMAL)
            self.machine_entry.config(state=tk.NORMAL)
            self.radio_txt.config(state=tk.NORMAL)
            self.radio_pdf.config(state=tk.NORMAL)
            self.radio_ambos.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.NORMAL)
            self.exit_btn.config(state=tk.NORMAL)
            self.desbloquear_fechar()  # Reabilita o botão X

    def get_computer_name_input(self):
        value = self.machine_var.get().strip()
        # Garante que o placeholder nunca seja considerado
        if not value or value == self.placeholder or value.lower() == self.placeholder.lower():
            return None  # None indica máquina local
        return value

    def on_listbox_double_click(self, event):
        selection = self.listbox_maquinas.curselection()
        if selection:
            valor = self.listbox_maquinas.get(selection[0])
            nome_maquina = valor.split(' - ')[0].strip()
            if nome_maquina and nome_maquina.lower() != 'carregando...' and nome_maquina.lower() != 'nenhuma máquina encontrada na rede.' and not nome_maquina.lower().startswith('erro'):
                self.machine_var.set(nome_maquina)
                self.machine_entry.config(fg='#222', font=('Segoe UI', 10), justify='center')

if __name__ == "__main__":
    app = CSInfoApp()
    app.mainloop()
