import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os
import csinfo
from csinfo import main as csinfo_main
import importlib.util, sys
import subprocess
import tkinter.font as tkfont


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


# Simples tooltip para widgets (mostrar Toplevel ao passar o mouse)
class Tooltip:
    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._id = None
        self._tip = None
        widget.bind('<Enter>', self._schedule)
        widget.bind('<Leave>', self._hide)
        widget.bind('<ButtonPress>', self._hide)

    def _schedule(self, event=None):
        self._unschedule()
        self._id = self.widget.after(self.delay, self._show)

    def _unschedule(self):
        if self._id:
            try:
                self.widget.after_cancel(self._id)
            except Exception:
                pass
            self._id = None

    def _show(self):
        if self._tip or not self.widget.winfo_ismapped():
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f'+{x}+{y}')
        lbl = tk.Label(self._tip, text=self.text, justify='left', background='#ffffe0', relief='solid', borderwidth=1, font=('Segoe UI', 9))
        lbl.pack(ipadx=6, ipady=3)

    def _hide(self, event=None):
        self._unschedule()
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


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

        # Botão para abrir modal de credenciais (pequeno) à esquerda do Start
        self.creds_btn = ttk.Button(frame_entrada, text="Credenciais...", command=self.open_credentials_modal)
        self.creds_btn.grid(row=0, column=3, padx=(8,0))

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
        # Campo para apelido (alias)
        label_alias = tk.Label(frame_entrada, text="Apelido:", font=("Segoe UI", 10, 'normal'), fg='#333', bg='#f4f6f9')
        label_alias.grid(row=1, column=0, padx=(0, 8), sticky='e')
        self.alias_var = tk.StringVar()
        self.alias_entry = tk.Entry(frame_entrada, textvariable=self.alias_var, font=("Segoe UI", 10), width=24, relief='solid', justify='center')
        self.alias_entry.grid(row=1, column=1, sticky='w')
        # Forçar apelido em caixa alta e limitar tamanho (comportamento consistente ao campo máquina)
        def alias_maiusculo(*args):
            valor = self.alias_var.get() or ''
            valor_maiusculo = valor.upper()
            if len(valor_maiusculo) > 40:
                valor_maiusculo = valor_maiusculo[:40]
            if valor != valor_maiusculo:
                self.alias_var.set(valor_maiusculo)
        self.alias_var.trace_add('write', alias_maiusculo)
        # Layout principal: container com duas colunas lado a lado (resultados | histórico)
        self.main_content = tk.Frame(self, bg='#f4f6f9')
        self.main_content.pack(fill='both', expand=True, padx=10, pady=(6, 6))

        # Left frame: área de resultados (vai conter frame_result)
        self.left_frame = tk.Frame(self.main_content, bg='#f4f6f9')
        self.left_frame.pack(side='left', fill='both', expand=True)

        # Right frame: histórico de máquinas (mantém largura fixa)
        self.right_frame = tk.Frame(self.main_content, bg='#f4f6f9', width=260)
        self.right_frame.pack(side='right', fill='y', padx=(8, 0))
        self.right_frame.pack_propagate(False)  # manter largura fixa

        tk.Label(self.right_frame, text="Máquinas salvas:", bg='#f4f6f9', font=("Segoe UI", 10)).pack(anchor='w', padx=6, pady=(6, 0))
        self.machines_listbox = tk.Listbox(self.right_frame, font=("Segoe UI", 9), activestyle='dotbox')
        self.machines_listbox.pack(fill='both', expand=True, pady=(6, 6), padx=6)
        # ligar seleção simples para habilitar Remover/Editar e bloquear o campo Apelido
        self.machines_listbox.bind('<<ListboxSelect>>', self.on_history_select)
        self.machines_listbox.bind('<Double-Button-1>', self.on_history_double_click)

        btns_hist = tk.Frame(self.right_frame, bg='#f4f6f9')
        btns_hist.pack(fill='x', padx=6, pady=(0, 6))
        self.btn_save = ttk.Button(btns_hist, text='Salvar', command=self.save_machine)
        self.btn_save.pack(side='left', padx=4, pady=4)
        self.btn_edit = ttk.Button(btns_hist, text='Editar', command=self.edit_selected_machine)
        self.btn_edit.pack(side='left', padx=4, pady=4)
        self.btn_remove = ttk.Button(btns_hist, text='Remover', command=self.remove_selected_machine)
        self.btn_remove.pack(side='left', padx=4, pady=4)
        # Iniciar com Remover/Editar desabilitados até seleção
        try:
            self.btn_remove.config(state=tk.DISABLED)
            self.btn_edit.config(state=tk.DISABLED)
        except Exception:
            pass
        self.btn_open_hist = ttk.Button(btns_hist, text='Abrir', command=self.open_history_folder)
        self.btn_open_hist.pack(side='left', padx=4, pady=4)

        # Arquivo de histórico JSON (migrado para pasta do usuário)
        try:
            if os.name == 'nt':
                base_dir = os.getenv('APPDATA') or os.path.expanduser('~')
                hist_dir = os.path.join(base_dir, 'CSInfo')
            else:
                hist_dir = os.path.join(os.path.expanduser('~'), '.csinfo')
            os.makedirs(hist_dir, exist_ok=True)
            self.history_path = os.path.join(hist_dir, 'machines_history.json')
            # migrar histórico antigo do diretório do projeto, se existir e não houver no novo local
            legacy_path = os.path.join(os.getcwd(), 'machines_history.json')
            if os.path.exists(legacy_path) and not os.path.exists(self.history_path):
                try:
                    import shutil
                    shutil.copy2(legacy_path, self.history_path)
                except Exception:
                    pass
        except Exception:
            # fallback para o diretório do projeto
            self.history_path = os.path.join(os.getcwd(), 'machines_history.json')

        self.machines_history = []
        self.editing_index = None
        self.load_history()

        # Limite do histórico (número máximo de entradas)
        self.MAX_HISTORY = 100

        # Label acima do campo de informações processadas
        self.label_info = tk.Label(self.left_frame, text="Informações processadas:", font=("Segoe UI", 10), fg='#333', bg='#f4f6f9')
        self.label_info.pack(anchor='w', padx=6, pady=(6, 0))

        # Área de resultados (dentro do left_frame para manter altura simétrica)
        self.frame_result = tk.Frame(self.left_frame, bg='#f4f6f9')
        # aumentar o espaço inicial do frame_result para dar mais área ao ScrolledText
        self.frame_result.pack(fill='both', expand=True, pady=(5, 6))
        # ScrolledText preenche o frame para manter mesma altura da lista de histórico
        # aumentamos 'height' para deslocar os controles para baixo
        # aumentar para 24 linhas (mais área, mas deixando espaço para os controles abaixo)
        self.info_text = ScrolledText(self.frame_result, width=1, height=24, font=('Consolas', 9),
                                      bg='white', fg='#222', relief='solid', borderwidth=1)
        self.info_text.pack(fill='both', expand=True)
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

        # Rodapé: Exportação (colocado dentro do left_frame para manter alinhamento)
        self.frame_export = tk.Frame(self.left_frame, bg='#f4f6f9')
        self.frame_export.pack(pady=(8, 5), fill='x')
        label_export = tk.Label(self.frame_export, text="Exportar resultado:", font=('Segoe UI', 10, 'normal'), fg='#333', bg='#f4f6f9')
        label_export.grid(row=0, column=0, padx=(0, 8))
        self.export_var = tk.StringVar(value='txt')
        self.radio_txt = tk.Radiobutton(self.frame_export, text='TXT', variable=self.export_var, value='txt', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_pdf = tk.Radiobutton(self.frame_export, text='PDF', variable=self.export_var, value='pdf', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_ambos = tk.Radiobutton(self.frame_export, text='Ambos', variable=self.export_var, value='ambos', font=('Segoe UI', 10), bg='#f4f6f9', activebackground='#e3e6ea', selectcolor='white', highlightthickness=0)
        self.radio_txt.grid(row=0, column=1, padx=5)
        self.radio_pdf.grid(row=0, column=2, padx=5)
        self.radio_ambos.grid(row=0, column=3, padx=5)
    # Nota: Checkbox 'Incluir diagnóstico' removido (opção não exibida no UI)

        # Botões Exportar e Sair (dentro do left_frame)
        self.frame_botoes = tk.Frame(self.left_frame, bg='#f4f6f9')
        self.frame_botoes.pack(pady=(6, 12), fill='x')
        self.export_btn = ttk.Button(self.frame_botoes, text="💾 Exportar", style='Export.TButton', command=self.exportar)
        self.export_btn.grid(row=0, column=0, padx=10)
        self.exit_btn = ttk.Button(self.frame_botoes, text="✖ Sair", style='Exit.TButton', command=self.quit)
        self.exit_btn.grid(row=0, column=1, padx=10)

        # Ajusta a altura do Listbox em linhas para alinhar exatamente com a altura do ScrolledText
        def adjust_listbox_height():
            try:
                # Garantir que a geometria esteja atualizada
                self.update_idletasks()
                # Usar a fonte do listbox para calcular altura de linha
                lb_font = tkfont.Font(font=self.machines_listbox.cget('font'))
                line_px = lb_font.metrics('linespace') or 12
                # Altura em pixels do widget de texto (conteúdo visível)
                text_h = self.info_text.winfo_height()
                if not text_h or text_h < line_px:
                    return
                # Calcular número de linhas para o listbox
                lines = max(3, int(text_h / line_px))
                # Aplicar sem quebrar se já estiver igual
                try:
                    current_h = int(self.machines_listbox.cget('height'))
                except Exception:
                    current_h = None
                if current_h != lines:
                    self.machines_listbox.config(height=lines)
            except Exception:
                pass

        # Chamar após pequenas mudanças de layout e quando o frame de resultado for redimensionado
        self.frame_result.bind('<Configure>', lambda e: self.after(50, adjust_listbox_height))
        # Agendamento inicial para ajustar logo após a criação da janela
        self.after(200, adjust_listbox_height)

        # Inicialmente, desabilita botões de exportação
        self.export_btn.config(state=tk.DISABLED)
        self.radio_txt.config(state=tk.DISABLED)
        self.radio_pdf.config(state=tk.DISABLED)
        self.radio_ambos.config(state=tk.DISABLED)

        # Tooltips para controles principais
        try:
            Tooltip(self.start_btn, 'Iniciar coleta de informações da máquina (local ou remota)')
            Tooltip(self.creds_btn, 'Definir credenciais remotas no formato DOMAIN\\user')
            Tooltip(self.alias_entry, 'Apelido usado para salvar/exportar o relatório (obrigatório para salvar no histórico)')
            Tooltip(self.machines_listbox, 'Lista de máquinas salvas. Duplo clique inicia a análise na máquina selecionada')
            Tooltip(self.btn_save, 'Salvar a máquina atual no histórico (é necessário informar um apelido)')
            Tooltip(self.btn_edit, 'Editar a entrada selecionada no histórico')
            Tooltip(self.btn_remove, 'Remover a entrada selecionada do histórico')
            Tooltip(self.btn_open_hist, 'Abrir a pasta que contém o arquivo de histórico (machines_history.json)')
            Tooltip(self.export_btn, 'Exportar relatório (TXT/PDF) para disco')
            Tooltip(self.exit_btn, 'Fechar o aplicativo')
            Tooltip(self.radio_txt, 'Exportar somente em formato TXT')
            Tooltip(self.radio_pdf, 'Exportar somente em formato PDF')
            Tooltip(self.radio_ambos, 'Exportar em ambos os formatos (TXT e PDF)')
            Tooltip(self.chk_debug, 'Incluir informações de diagnóstico adicionais no arquivo TXT')
        except Exception:
            pass

        # Sincronizar dinamicamente a altura da listbox com a área de resultados
        def sync_history_height(event=None):
            try:
                h = self.frame_result.winfo_height()
                if h and h > 10:
                    extra = 80  # espaço para label e botões inferior
                    self.right_frame.config(height=h + extra)
                    self.right_frame.pack_propagate(False)
            except Exception:
                pass

        self.frame_result.bind('<Configure>', sync_history_height)
        self.bind('<Configure>', sync_history_height)

        # Rodapé centralizado com versão
        # Rodapé sem versão do app (apenas o nome da empresa)
        rodape_text = "CEOsoftware Sistemas"
        # Rodapé centralizado
        rodape = tk.Label(self, text=rodape_text, font=("Segoe UI", 8), fg="#666", bg="#f4f6f9")
        rodape.pack(side=tk.BOTTOM, pady=(0, 6), fill='x')
        rodape.configure(anchor="center", justify="center")

    def show_about(self):
        try:
            ver = getattr(csinfo, '__version__', 'desconhecida')
            info = f"CSInfo — versão {ver}\n\nGerado por CEOsoftware Sistemas"
        except Exception:
            info = "CSInfo — versão desconhecida"
        messagebox.showinfo("Sobre CSInfo", info)

    def open_credentials_modal(self):
        """Abre um modal simples para aceitar DOMAIN\\user e senha. Define credencial padrão em csinfo."""
        try:
            modal = tk.Toplevel(self)
            modal.title("Credenciais remotas")
            modal.geometry("420x180")
            modal.transient(self)
            modal.grab_set()

            lbl = tk.Label(modal, text="Informe credenciais no formato DOMAIN\\user", font=("Segoe UI", 10))
            lbl.pack(pady=(10, 6))

            frame = tk.Frame(modal)
            frame.pack(pady=(6, 6), padx=10, fill='x')
            tk.Label(frame, text="Usuário:", width=12, anchor='w').grid(row=0, column=0)
            user_var = tk.StringVar()
            user_entry = tk.Entry(frame, textvariable=user_var, width=36)
            user_entry.grid(row=0, column=1)

            tk.Label(frame, text="Senha:", width=12, anchor='w').grid(row=1, column=0)
            pwd_var = tk.StringVar()
            pwd_entry = tk.Entry(frame, textvariable=pwd_var, width=36, show='*')
            pwd_entry.grid(row=1, column=1)

            def on_ok():
                u = user_var.get().strip()
                p = pwd_var.get()
                if not u or '\\' not in u:
                    messagebox.showwarning("Formato inválido", "Informe no formato DOMAIN\\user")
                    return
                try:
                    csinfo.set_default_credential(u, p)
                    messagebox.showinfo("Credenciais definidas", "Credenciais salvas para uso nas chamadas remotas.")
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível salvar as credenciais: {e}")
                modal.grab_release()
                modal.destroy()

            def on_clear():
                try:
                    csinfo.clear_default_credential()
                    messagebox.showinfo("Credenciais", "Credenciais padrão removidas.")
                except Exception:
                    messagebox.showwarning("Aviso", "Não foi possível remover credenciais (ou nenhuma estava definida).")
                modal.grab_release()
                modal.destroy()

            btn_frame = tk.Frame(modal)
            btn_frame.pack(pady=(6, 8))
            ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok)
            ok_btn.grid(row=0, column=0, padx=6)
            clear_btn = ttk.Button(btn_frame, text="Remover credenciais", command=on_clear)
            clear_btn.grid(row=0, column=1, padx=6)
            cancel_btn = ttk.Button(btn_frame, text="Cancelar", command=lambda: (modal.grab_release(), modal.destroy()))
            cancel_btn.grid(row=0, column=2, padx=6)

            # Focar no campo usuário
            user_entry.focus_set()
            modal.wait_window()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o modal de credenciais: {e}")

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
                machine_name = valor_input.strip()
            # pegar apelido se informado
            alias = self.alias_var.get().strip() or None
            print('DEBUG: machine_name:', machine_name, 'alias:', alias)
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
            # Se o usuário definiu credenciais via GUI, elas já foram aplicadas via csinfo.set_default_credential
            # Apenas coleta os dados na análise; não gerar arquivos automaticamente
            # A opção 'Incluir diagnóstico' foi removida da UI; não passamos esse flag
            resultado = csinfo_main(export_type=None, barra_callback=gui_callback, computer_name=machine_name, machine_alias=alias)
            # garantir que o alias seja preservado para exportação posterior
            try:
                if isinstance(resultado, dict):
                    resultado['alias'] = alias
            except Exception:
                pass
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
                    # Se o último resultado tiver um alias, respeitar o padrão Info_maquina_<alias>_<nomemaquina>.txt
                    last_alias = self.last_result.get('alias') if isinstance(self.last_result, dict) else None
                    if last_alias:
                        base_path = os.path.join(os.getcwd(), f"Info_maquina_{csinfo.safe_filename(last_alias)}_{safe_name}.txt")
                    else:
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
                alias_for_call = self.alias_var.get().strip() or None
                resultado = csinfo_main(export_type=tipo, barra_callback=None, computer_name=machine_name, machine_alias=alias_for_call)
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

    # --- Histórico de máquinas (JSON) ---
    def load_history(self):
        try:
            if os.path.exists(self.history_path):
                with open(self.history_path, 'r', encoding='utf-8') as f:
                    data = f.read().strip()
                    if not data:
                        self.machines_history = []
                    else:
                        import json
                        self.machines_history = json.loads(data)
            else:
                self.machines_history = []
        except Exception:
            self.machines_history = []
        # ordenar por alias (case-insensitive) e truncar ao limite
        try:
            self.machines_history = sorted(self.machines_history, key=lambda x: (x.get('alias') or '').lower())
            if hasattr(self, 'MAX_HISTORY') and isinstance(self.MAX_HISTORY, int):
                self.machines_history = self.machines_history[:self.MAX_HISTORY]
        except Exception:
            pass

        # popular listbox
        try:
            self.machines_listbox.delete(0, tk.END)
            for item in self.machines_history:
                alias = item.get('alias')
                name = item.get('name')
                display = f"{name} - {alias}" if alias else name
                self.machines_listbox.insert(tk.END, display)
        except Exception:
            pass
        # Após carregar histórico, garantir estado dos controles
        try:
            self.btn_remove.config(state=tk.DISABLED)
            self.btn_edit.config(state=tk.DISABLED)
            # deixar apelido desabilitado até o usuário clicar em Editar
            self.alias_entry.config(state='disabled')
        except Exception:
            pass

    def save_machine(self):
        # salvar o par {name, alias} requer alias não vazio
        name = self.get_computer_name_input() or csinfo.get_machine_name(None)
        raw_alias = self.alias_var.get().strip()
        if not raw_alias:
            messagebox.showwarning("Apelido obrigatório", "Para salvar uma máquina no histórico, informe um apelido.")
            return
        # sanitizar apelido para armazenagem e para uso em filename
        safe_alias = csinfo.safe_filename(raw_alias)
        if safe_alias != raw_alias:
            messagebox.showinfo("Apelido alterado", f"O apelido informado foi ajustado para um formato seguro: {safe_alias}")
        entry = {'name': name, 'alias': safe_alias}
        # Se estamos editando um item, substituir
        if self.editing_index is not None and 0 <= self.editing_index < len(self.machines_history):
            # evitar duplicar outro alias
            self.machines_history = [m for m in self.machines_history if m.get('alias') != safe_alias or m is self.machines_history[self.editing_index]]
            self.machines_history[self.editing_index] = entry
            self.editing_index = None
        else:
            # evitar duplicatas pelo alias
            self.machines_history = [m for m in self.machines_history if m.get('alias') != safe_alias]
            self.machines_history.insert(0, entry)

        try:
            import json
            with open(self.history_path, 'w', encoding='utf-8') as f:
                # garantir ordenação antes de salvar e truncamento
                try:
                    self.machines_history = sorted(self.machines_history, key=lambda x: (x.get('alias') or '').lower())
                    if hasattr(self, 'MAX_HISTORY') and isinstance(self.MAX_HISTORY, int):
                        self.machines_history = self.machines_history[:self.MAX_HISTORY]
                except Exception:
                    pass
                json.dump(self.machines_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o histórico: {e}")
            return
        self.load_history()
        # Após salvar, bloquear o campo apelido novamente e desabilitar editar/remover
        try:
            self.alias_entry.config(state='disabled')
            self.btn_edit.config(state=tk.DISABLED)
            self.btn_remove.config(state=tk.DISABLED)
        except Exception:
            pass

    def remove_selected_machine(self):
        try:
            sel = self.machines_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            item = self.machines_history[idx]
            alias = item.get('alias')
            if messagebox.askyesno("Confirmar remoção", f"Deseja remover '{alias}' do histórico? Esta ação não pode ser desfeita."):
                self.machines_history = [m for m in self.machines_history if m.get('alias') != alias]
                import json
                with open(self.history_path, 'w', encoding='utf-8') as f:
                    json.dump(self.machines_history, f, ensure_ascii=False, indent=2)
                self.load_history()
                # Depois de remover, bloquear apelido e desabilitar botões
                try:
                    self.alias_entry.config(state='disabled')
                    self.btn_edit.config(state=tk.DISABLED)
                    self.btn_remove.config(state=tk.DISABLED)
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível remover: {e}")

    def on_history_double_click(self, event):
        try:
            sel = self.machines_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            item = self.machines_history[idx]
            name = item.get('name')
            alias = item.get('alias')
            if name:
                self.machine_var.set(name)
                self.machine_entry.config(fg='#222', font=('Segoe UI', 10), justify='center')
            if alias:
                self.alias_var.set(alias)
            # manter apelido bloqueado (somente leitura) após selecionar
            try:
                self.alias_entry.config(state='disabled')
            except Exception:
                pass
            # iniciar processamento automaticamente (mesmo comportamento do botão Iniciar)
            self.start_process()
        except Exception:
            pass

    def edit_selected_machine(self):
        try:
            sel = self.machines_listbox.curselection()
            if not sel:
                messagebox.showwarning("Selecionar", "Selecione uma máquina na lista para editar.")
                return
            idx = sel[0]
            item = self.machines_history[idx]
            name = item.get('name')
            alias = item.get('alias')
            # preencher campos e colocar em modo edição
            if name:
                self.machine_var.set(name)
                self.machine_entry.config(fg='#222', font=('Segoe UI', 10), justify='center')
            if alias:
                self.alias_var.set(alias)
            self.editing_index = idx
            # permitir edição do apelido enquanto estiver em modo edição
            try:
                self.alias_entry.config(state='normal')
                self.alias_entry.focus_set()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível iniciar edição: {e}")

    def open_history_folder(self):
        try:
            pasta = os.path.dirname(self.history_path)
            if not os.path.exists(pasta):
                messagebox.showwarning("Aviso", "Pasta de histórico não encontrada.")
                return
            # abrir no explorer no Windows, no caso de outros sistemas tentar abrir com o método padrão
            try:
                if os.name == 'nt':
                    subprocess.Popen(['explorer', pasta])
                else:
                    # tentar abrir com xdg-open, open (mac) ou fallback para listar
                    if shutil.which('xdg-open'):
                        subprocess.Popen(['xdg-open', pasta])
                    elif shutil.which('open'):
                        subprocess.Popen(['open', pasta])
                    else:
                        messagebox.showinfo('Pasta', f"Pasta do histórico: {pasta}")
            except Exception:
                messagebox.showinfo('Pasta', f"Pasta do histórico: {pasta}")
        except Exception as e:
            messagebox.showerror('Erro', f"Não foi possível abrir a pasta: {e}")

    def on_history_select(self, event):
        """Handler chamado quando um item do histórico é selecionado.
        Habilita os botões Editar/Remover e mantém o campo Apelido desabilitado até Editar ser clicado.
        """
        try:
            sel = self.machines_listbox.curselection()
            if not sel:
                try:
                    self.btn_remove.config(state=tk.DISABLED)
                    self.btn_edit.config(state=tk.DISABLED)
                except Exception:
                    pass
                return
            # item selecionado
            try:
                self.btn_remove.config(state=tk.NORMAL)
                self.btn_edit.config(state=tk.NORMAL)
            except Exception:
                pass
            # garantir apelido em modo somente leitura até Editar
            try:
                self.alias_entry.config(state='disabled')
            except Exception:
                pass
        except Exception:
            pass

if __name__ == "__main__":
    app = CSInfoApp()
    app.mainloop()
