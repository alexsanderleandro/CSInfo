
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
from csinfo import main as csinfo_main

# --------- Função para configurar estilos visuais ---------
def estilo_botoes():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TButton',
        font=('Segoe UI', 10, 'bold'),
        foreground='white',
        background='#1976D2',
        borderwidth=0,
        padding=10,
        relief='flat'
    )
    style.map('TButton',
        background=[('active', '#1565C0'), ('hover', '#1565C0')],
        foreground=[('disabled', '#ccc')]
    )
    style.configure('Rounded.TButton',
        font=('Segoe UI', 10, 'bold'),
        foreground='white',
        background='#1976D2',
        borderwidth=0,
        padding=10,
        relief='flat'
    )
    style.map('Rounded.TButton',
        background=[('active', '#1565C0'), ('hover', '#1565C0')],
        foreground=[('disabled', '#ccc')]
    )


# --------- Classe principal adaptada ---------
class CSInfoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSInfo – Inventário de Hardware e Software")
        self.geometry("900x650")
        self.configure(bg='#f4f6f9')
        self.resizable(False, False)
        estilo_botoes()
        self.criar_layout()

    def criar_layout(self):
        # Título
        self.label = tk.Label(self, text="CSInfo – Inventário de Hardware e Software", font=("Segoe UI", 16, "bold"), fg="#003366", bg="#f4f6f9")
        self.label.pack(pady=(20, 10))

        # Frame de entrada
        frame_entrada = tk.Frame(self, bg='#f4f6f9')
        frame_entrada.pack(pady=(10, 5))
        label_nome = tk.Label(frame_entrada, text="Nome da máquina:", font=("Segoe UI", 10, "bold"), fg="#333", bg="#f4f6f9")
        label_nome.grid(row=0, column=0, padx=(0, 8), sticky='e')

        self.machine_var = tk.StringVar()
        self.machine_entry = tk.Entry(frame_entrada, textvariable=self.machine_var, font=('Segoe UI', 10), width=40, relief='solid', justify='center')
        self.machine_entry.grid(row=0, column=1)
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
        self.machine_entry.bind("<FocusIn>", on_focus_in)
        self.machine_entry.bind("<FocusOut>", on_focus_out)

        # Botão Iniciar
        self.start_btn = ttk.Button(self, text="▶ Iniciar", style='Rounded.TButton', command=self.start_process)
        self.start_btn.pack(pady=(10, 15))

        # Área de resultados
        frame_result = tk.Frame(self, bg='#f4f6f9')
        frame_result.pack(pady=(5, 10))
        self.info_text = ScrolledText(frame_result, width=110, height=20, font=('Consolas', 9),
                                      bg='white', fg='#222', relief='solid', borderwidth=1)
        self.info_text.pack()
        self.info_text.config(state=tk.DISABLED)

        # Barra de progresso / mensagem
        self.progress_label = tk.Label(self, text="", font=('Segoe UI', 10, 'bold'), fg='#1976D2', bg='#f4f6f9')
        self.progress_label.pack(pady=(5, 0))

        # Rodapé: Exportação
        frame_export = tk.Frame(self, bg='#f4f6f9')
        frame_export.pack(pady=(15, 5))
        label_export = tk.Label(frame_export, text="Exportar resultado:", font=('Segoe UI', 10, 'bold'), fg='#333', bg='#f4f6f9')
        label_export.grid(row=0, column=0, padx=(0, 8))
        self.export_var = tk.StringVar(value='txt')
        self.radio_txt = ttk.Radiobutton(frame_export, text='TXT', variable=self.export_var, value='txt')
        self.radio_pdf = ttk.Radiobutton(frame_export, text='PDF', variable=self.export_var, value='pdf')
        self.radio_ambos = ttk.Radiobutton(frame_export, text='Ambos', variable=self.export_var, value='ambos')
        self.radio_txt.grid(row=0, column=1, padx=5)
        self.radio_pdf.grid(row=0, column=2, padx=5)
        self.radio_ambos.grid(row=0, column=3, padx=5)

        # Botões Exportar e Sair
        frame_botoes = tk.Frame(self, bg='#f4f6f9')
        frame_botoes.pack(pady=(10, 20))
        self.export_btn = ttk.Button(frame_botoes, text="💾 Exportar", style='Rounded.TButton', command=self.exportar)
        self.export_btn.grid(row=0, column=0, padx=10)
        self.exit_btn = ttk.Button(frame_botoes, text="✖ Sair", style='Rounded.TButton', command=self.quit)
        self.exit_btn.grid(row=0, column=1, padx=10)

        # Inicialmente, desabilita botões de exportação
        self.export_btn.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)

    # ...métodos start_process, run_csinfo, exportar permanecem iguais, apenas adaptando para os novos widgets...

    def start_process(self):
        self.start_btn.config(state=tk.DISABLED)
        self.machine_entry.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.exit_btn.config(state=tk.DISABLED)
        self.progress_label.config(text="", fg="black")
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.config(state=tk.DISABLED)
        threading.Thread(target=self.run_csinfo, daemon=True).start()

    def run_csinfo(self):
        try:
            import platform
            valor_input = self.machine_var.get().strip()
            if not valor_input or valor_input == self.placeholder:
                machine_name = None
            elif valor_input.upper() == platform.node().upper():
                machine_name = None
            else:
                machine_name = valor_input
            if machine_name:
                self.progress_label.config(text=f"Acessando a máquina {machine_name}...", fg="black", font=("Helvetica", 10, "bold"))
                self.update_idletasks()
            from csinfo import check_remote_machine
            if machine_name and not check_remote_machine(machine_name):
                self.progress_label.config(text=f"Não foi possível acessar a máquina '{machine_name}'. Ela pode estar desligada, fora da rede ou sem WinRM ativado.")
                self.start_btn.config(state=tk.NORMAL)
                self.machine_entry.config(state=tk.NORMAL)
                self.radio_txt.config(state=tk.DISABLED)
                self.radio_pdf.config(state=tk.DISABLED)
                self.radio_ambos.config(state=tk.DISABLED)
                self.export_btn.config(state=tk.DISABLED)
                self.exit_btn.config(state=tk.NORMAL)
                return
            self.capturado = []
            self.info_text.config(state=tk.NORMAL)
            self.info_text.delete(1.0, tk.END)
            def gui_callback(_, etapa_texto=None):
                if etapa_texto:
                    self.progress_label.config(text=etapa_texto, font=("Helvetica", 10, "bold"))
                    self.progress_label.update_idletasks()
            resultado = csinfo_main(export_type=self.export_var.get(), barra_callback=gui_callback, computer_name=machine_name)
            self.progress_label.config(text="Análise finalizada!", fg="red", font=("Helvetica", 10, "bold"))
            # Ao final, mostra apenas o relatório completo
            txt_path = resultado.get('txt')
            if txt_path and os.path.exists(txt_path):
                self.info_text.config(state=tk.NORMAL)
                self.info_text.delete(1.0, tk.END)
                with open(txt_path, encoding='utf-8') as f:
                    conteudo = f.read()
                self.info_text.insert(tk.END, conteudo)
                self.info_text.see(tk.END)
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
            messagebox.showerror("Erro", str(e))
            self.export_btn.config(state=tk.DISABLED)
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.machine_entry.config(state=tk.NORMAL)
            self.radio_txt.config(state=tk.NORMAL)
            self.radio_pdf.config(state=tk.NORMAL)
            self.radio_ambos.config(state=tk.NORMAL)
            self.exit_btn.config(state=tk.NORMAL)

    def update_progress(self, value, etapa_texto=None):
        def atualizar():
            if etapa_texto:
                self.progress_label.config(text=etapa_texto)
            self.update_idletasks()
        self.after(0, atualizar)

    def exportar(self):
        # Desabilita todos os campos do formulário
        self.start_btn.config(state=tk.DISABLED)
        self.machine_entry.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.exit_btn.config(state=tk.DISABLED)
        self.progress_label.config(text="Exportando relatório, aguarde...", font=("Helvetica", 10, "bold"))
        self.update_idletasks()
        try:
            machine_name = self.machine_var.get().strip() or None
            tipo = self.export_var.get()
            from csinfo import main as csinfo_main
            resultado = csinfo_main(export_type=tipo, barra_callback=None, computer_name=machine_name)
            msg = []
            if resultado.get('txt'):
                msg.append(f"Arquivo TXT exportado: {resultado['txt']}")
            if resultado.get('pdf'):
                msg.append(f"Arquivo PDF exportado: {resultado['pdf']}")
            messagebox.showinfo("Exportação", "\n".join(msg) if msg else "Nenhum arquivo exportado.")
            if resultado.get('txt') or resultado.get('pdf'):
                caminho = resultado.get('txt') or resultado.get('pdf')
                # Pergunta se deseja abrir a pasta
                abrir = messagebox.askyesno("Abrir pasta", "Deseja abrir o diretório onde o relatório foi salvo?")
                if abrir:
                    import os
                    import subprocess
                    pasta = os.path.dirname(caminho)
                    try:
                        subprocess.Popen(f'explorer "{pasta}"')
                    except Exception as e:
                        messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")
        except Exception as e:
            messagebox.showerror("Erro na exportação", str(e))
        finally:
            self.progress_label.config(text="")
            # Reabilita todos os campos após exportação
            self.start_btn.config(state=tk.NORMAL)
            self.machine_entry.config(state=tk.NORMAL)
            self.radio_txt.config(state=tk.NORMAL)
            self.radio_pdf.config(state=tk.NORMAL)
            self.radio_ambos.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.NORMAL)
            self.exit_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    app = CSInfoApp()
    app.mainloop()
