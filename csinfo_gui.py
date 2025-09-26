import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
from csinfo import main as csinfo_main

class CSInfoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSInfo - Coletor de Informações")
        self.geometry("900x550")
        self.resizable(False, False)
        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self, text="CSInfo by CEOsoftware", font=("Helvetica", 14, "bold"))
        self.label.pack(pady=10)

        self.export_var = tk.StringVar(value="txt")
        frame = tk.Frame(self)
        frame.pack(pady=10)
        tk.Label(frame, text="Tipo de exportação:").pack(side=tk.LEFT)
        self.radio_txt = ttk.Radiobutton(frame, text="TXT", variable=self.export_var, value="txt")
        self.radio_txt.pack(side=tk.LEFT)
        self.radio_pdf = ttk.Radiobutton(frame, text="PDF", variable=self.export_var, value="pdf")
        self.radio_pdf.pack(side=tk.LEFT)
        self.radio_ambos = ttk.Radiobutton(frame, text="Ambos", variable=self.export_var, value="ambos")
        self.radio_ambos.pack(side=tk.LEFT)

        self.machine_var = tk.StringVar()
        def to_uppercase(*args):
            value = self.machine_var.get()
            self.machine_var.set(value.upper())
        self.machine_var.trace_add('write', to_uppercase)
        machine_frame = tk.Frame(self)
        machine_frame.pack(pady=5)
        tk.Label(machine_frame, text="Nome da máquina:").pack(side=tk.LEFT)
        self.machine_entry = tk.Entry(machine_frame, textvariable=self.machine_var, width=25)
        self.machine_entry.pack(side=tk.LEFT)

        self.start_btn = tk.Button(self, text="Iniciar", command=self.start_process, width=15)
        self.start_btn.pack(pady=10)

        self.progress_label = tk.Label(self, text="", font=("Helvetica", 10), wraplength=500, justify=tk.LEFT)
        self.progress_label.pack(pady=2)

        self.info_text = tk.Text(self, height=20, width=110, font=("Consolas", 9))
        self.info_text.pack(pady=5)
        self.info_text.config(state=tk.DISABLED)

        self.export_btn = tk.Button(self, text="Exportar", command=self.exportar, width=15)
        self.export_btn.pack(pady=10)
        self.export_btn.config(state=tk.DISABLED)

        self.exit_btn = tk.Button(self, text="Sair", command=self.quit, width=15)
        self.exit_btn.pack(pady=10)

    def start_process(self):
        self.start_btn.config(state=tk.DISABLED)
        self.machine_entry.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.exit_btn.config(state=tk.DISABLED)
        self.progress_label.config(text="")
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)
        self.info_text.config(state=tk.DISABLED)
        threading.Thread(target=self.run_csinfo, daemon=True).start()

    def run_csinfo(self):
        try:
            machine_name = self.machine_var.get().strip() or None
            if machine_name:
                self.progress_label.config(text=f"Acessando a máquina {machine_name}...")
                self.update_idletasks()
            from csinfo import check_remote_machine
            if machine_name and not check_remote_machine(machine_name):
                self.progress_label.config(text=f"Não foi possível acessar a máquina '{machine_name}'. Ela pode estar desligada, fora da rede ou sem WinRM ativado.")
                self.start_btn.config(state=tk.NORMAL)
                self.machine_entry.config(state=tk.NORMAL)
                self.radio_txt.config(state=tk.NORMAL)
                self.radio_pdf.config(state=tk.NORMAL)
                self.radio_ambos.config(state=tk.NORMAL)
                self.export_btn.config(state=tk.DISABLED)
                self.exit_btn.config(state=tk.NORMAL)
                return
            self.capturado = []
            self.info_text.config(state=tk.NORMAL)
            self.info_text.delete(1.0, tk.END)
            def gui_callback(_, etapa_texto=None):
                if etapa_texto:
                    self.progress_label.config(text=etapa_texto)
                    self.progress_label.update_idletasks()
            resultado = csinfo_main(export_type=self.export_var.get(), barra_callback=gui_callback, computer_name=machine_name)
            self.progress_label.config(text="Processo finalizado!")
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

if __name__ == "__main__":
    app = CSInfoApp()
    app.mainloop()
