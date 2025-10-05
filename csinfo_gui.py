r"""CSInfo GUI

Front-end Tkinter limpo e consolidado para o backend `csinfo`.

Funcionalidades:
- Chama csinfo.main(...) com barra_callback para streaming de linhas/progresso
- Persiste máquinas em %APPDATA%\CSInfo\machines_history.json
- Ping para status ONLINE/OFFLINE
- Campos de credenciais (usuário/senha)
- Exporta TXT/PDF com padrão Info_maquina_<apelido>_<nomemaquina>
- UI bloqueada enquanto processamento está em andamento
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import re

try:
    import csinfo
except Exception:
    csinfo = None
try:
    from version import __version__
except Exception:
    __version__ = '0.0.0'

# assets
ASSETS_DIR = os.path.join(os.path.dirname(__file__), 'assets')
APP_ICON_ICO = os.path.join(ASSETS_DIR, 'app.ico')
APP_ICON_PNG = os.path.join(ASSETS_DIR, 'ico.png')
APP_LOGO = APP_ICON_PNG  # usado no relatório PDF


def get_appdata_path():
    if sys.platform.startswith('win'):
        return os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.expanduser('~')


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


class CSInfoGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        # título e icon
        self.title(f'CSInfo v{__version__}')
        try:
            # preferir ico para exe; para Tkinter usar PhotoImage para PNG
            if os.path.exists(APP_ICON_PNG):
                try:
                    img = tk.PhotoImage(file=APP_ICON_PNG)
                    self.iconphoto(False, img)
                except Exception:
                    pass
            elif os.path.exists(APP_ICON_ICO) and sys.platform.startswith('win'):
                try:
                    self.iconbitmap(APP_ICON_ICO)
                except Exception:
                    pass
        except Exception:
            pass
        self.geometry('980x640')

        # state
        self.queue = queue.Queue()
        self.worker_thread = None
        self._processing = False
        self._pinging = False
        self._keep_progress = False
        self.last_lines = []
        self._last_collection_computer = None
        self.machine_list = []
        self.machine_json_path = self._get_machine_json_path()

        # UI
        self._build_ui()
        self.load_machine_list()
        self.protocol('WM_DELETE_WINDOW', self._on_close_attempt)
        self.after(100, self._process_queue)

    def _get_machine_json_path(self):
        base = os.path.join(get_appdata_path(), 'CSInfo')
        try:
            os.makedirs(base, exist_ok=True)
        except Exception:
            base = os.getcwd()
        return os.path.join(base, 'machines_history.json')

    def _build_ui(self):
        frm = ttk.Frame(self)
        frm.pack(fill='both', expand=True, padx=8, pady=8)

        left = ttk.Frame(frm)
        left.pack(side='left', fill='y')

        ttk.Label(left, text='Máquina:').pack(anchor='w')
        self.ent_computer = ttk.Entry(left, width=32, justify='center')
        self.ent_computer.pack()
        self.ent_computer.bind('<KeyRelease>', self._on_name_keyrelease)

        ttk.Label(left, text='Apelido:').pack(anchor='w', pady=(8, 0))
        self.ent_alias = ttk.Entry(left, width=32, justify='center')
        self.ent_alias.pack()
        self.ent_alias.bind('<KeyRelease>', self._on_alias_keyrelease)

        btn_fr = ttk.Frame(left)
        btn_fr.pack(pady=8, fill='x')
        # Nova | Salvar | Excluir
        self.btn_new = ttk.Button(btn_fr, text='Nova', command=self.new_machine)
        self.btn_new.pack(side='left', fill='x', expand=True)
        self.btn_save = ttk.Button(btn_fr, text='Salvar', command=self.save_selected_or_new_machine)
        self.btn_save.pack(side='left', fill='x', expand=True, padx=(6, 0))
        self.btn_delete = ttk.Button(btn_fr, text='Excluir', command=self.delete_selected_machine)
        self.btn_delete.pack(side='left', fill='x', expand=True, padx=(6, 0))

        # separador fino antes das credenciais
        sep = ttk.Separator(left, orient='horizontal')
        sep.pack(fill='x', pady=(8, 8))

        # Campos de credenciais movidos para baixo dos botões (rótulos completos)
        # rótulo de seção
        ttk.Label(left, text='Administrador da rede', font=('Segoe UI', 9, 'bold'), foreground='#333').pack(anchor='w', pady=(0, 6))
        ttk.Label(left, text='Usuário:').pack(anchor='w')
        self.ent_user = ttk.Entry(left, width=32)
        self.ent_user.pack()
        ttk.Label(left, text='Senha:').pack(anchor='w', pady=(8, 0))
        self.ent_pass = ttk.Entry(left, width=32, show='*')
        self.ent_pass.pack()
        
        # separador após credenciais do administrador da rede
        ttk.Separator(left, orient='horizontal').pack(fill='x', pady=(8, 10))

        # Botão Coletar logo abaixo do separador das credenciais
        self.btn_start = ttk.Button(left, text='Coletar (F3)', command=self.start_collection)
        self.btn_start.pack(fill='x', pady=(0, 8))
        try:
            Tooltip(self.btn_start, 'Inicia coleta (atalho F3)')
        except Exception:
            pass
        # bind F3 para iniciar coleta
        try:
            self.bind('<F3>', lambda e: self.start_collection())
        except Exception:
            pass

        # separador antes da seção de exportação
        ttk.Separator(left, orient='horizontal').pack(fill='x', pady=(8, 8))

        ttk.Label(left, text='Exportar como:').pack(anchor='w', pady=(8, 0))
        # radiobuttons para formato de exportação: 'pdf' (padrão), 'txt', 'ambos'
        self.export_var = tk.StringVar(value='pdf')
        self.export_frame = ttk.Frame(left)
        self.export_frame.pack(fill='x')
        self.rb_pdf = ttk.Radiobutton(self.export_frame, text='PDF', value='pdf', variable=self.export_var, command=self._update_export_button_state)
        self.rb_pdf.pack(side='left', padx=(0, 6))
        self.rb_txt = ttk.Radiobutton(self.export_frame, text='TXT', value='txt', variable=self.export_var, command=self._update_export_button_state)
        self.rb_txt.pack(side='left', padx=(0, 6))
        self.rb_both = ttk.Radiobutton(self.export_frame, text='Ambos', value='ambos', variable=self.export_var, command=self._update_export_button_state)
        self.rb_both.pack(side='left')

        self.btn_export = ttk.Button(left, text='Exportar (F10)', command=self._do_export, state='disabled')
        self.btn_export.pack(fill='x', pady=(8, 0))
        # opção para abrir automaticamente a pasta do arquivo exportado (inicialmente desabilitada até haver coleta)
        self.open_export_dir_var = tk.BooleanVar(value=False)
        try:
            self.chk_open_export_dir = ttk.Checkbutton(left, text='Listar diretório automaticamente', variable=self.open_export_dir_var, state='disabled')
            self.chk_open_export_dir.pack(anchor='w', pady=(6, 0))
            # separador logo abaixo do checkbox
            ttk.Separator(left, orient='horizontal').pack(fill='x', pady=(6, 8))
        except Exception:
            pass
        try:
            self.bind('<F10>', lambda e: self._do_export())
            Tooltip(self.btn_export, 'Exportar relatório (F10)')
        except Exception:
            pass

        self.btn_open_folder = ttk.Button(left, text='Abrir pasta de máquinas', command=self.open_machine_json_folder)
        self.btn_open_folder.pack(fill='x', pady=(8, 0))

        # Botão para atualizar status das máquinas (F5)
        self.btn_refresh = ttk.Button(left, text='Atualizar (F5)', command=self.refresh_machine_status)
        self.btn_refresh.pack(fill='x', pady=(8, 0))
        try:
            Tooltip(self.btn_refresh, 'Atualiza o estado (ONLINE/OFFLINE) das máquinas')
        except Exception:
            pass
        try:
            self.bind('<F5>', lambda e: self.refresh_machine_status())
        except Exception:
            pass

        mid = ttk.Frame(frm)
        mid.pack(side='left', fill='both', expand=True, padx=(8, 8))
        self.tree = ttk.Treeview(mid, columns=('name', 'alias', 'status'), show='headings', selectmode='browse')
        # habilitar ordenação clicando no cabeçalho
        self._sort_column = None
        self._sort_reverse = False
        # labels originais (usados para compor o texto com a seta)
        self._col_labels = {'name': 'Máquina', 'alias': 'Apelido', 'status': 'Estado'}
        # aplicar headings via helper que também desenha a seta quando necessário
        def _make_heading(col):
            self.tree.heading(col, text=self._col_labels.get(col, col), command=lambda c=col: self._sort_by_column(c))
        _make_heading('name')
        _make_heading('alias')
        _make_heading('status')
        self.tree.column('name', width=200, anchor='center')
        self.tree.column('alias', width=160, anchor='center')
        self.tree.column('status', width=90, anchor='center')
        self.tree.pack(fill='both', expand=True)
        try:
            # online: fundo verde escuro com texto branco
            self.tree.tag_configure('online', background='#006400', foreground='#ffffff')
            # offline: fundo vermelho escuro com texto branco
            self.tree.tag_configure('offline', background='#8B0000', foreground='#ffffff')
        except Exception:
            pass
        # ao selecionar (single-click) limpar saída e desabilitar export até nova coleta
        self.tree.bind('<<TreeviewSelect>>', lambda e: self._on_tree_selection_change())
        self.tree.bind('<Double-1>', lambda e: self._load_selection_into_form())
        # menu de contexto (botão direito) na lista de máquinas
        try:
            self.tree_menu = tk.Menu(self, tearoff=0)
            self.tree_menu.add_command(label='Reiniciar', command=lambda: self._on_context_restart())
            self.tree_menu.add_command(label='Desligar', command=lambda: self._on_context_shutdown())
            # bind right click
            self.tree.bind('<Button-3>', self._on_tree_right_click)
        except Exception:
            self.tree_menu = None

        right = ttk.Frame(frm)
        right.pack(side='right', fill='both', expand=True)
        self.txt_output = ScrolledText(right, height=20, state='disabled', wrap='word', font=('Consolas', 10))
        self.txt_output.pack(fill='both', expand=True)

        bar_fr = ttk.Frame(right)
        bar_fr.pack(fill='x')
        self.progress = ttk.Progressbar(bar_fr, orient='horizontal', mode='determinate')
        self.progress.pack(fill='x', side='left', expand=True, padx=(0, 8))
        # indicador temporário usado durante refresh de ping
        self.lbl_ping_status = ttk.Label(bar_fr, text='', foreground='#0066cc')
        self.lbl_ping_status.pack(side='right', padx=(0, 8))
        self.lbl_progress = ttk.Label(bar_fr, text='Pronto')
        self.lbl_progress.pack(side='right')

        rodape = tk.Label(self, text='CSInfo GUI', font=('Segoe UI', 8), fg='#666')
        rodape.pack(side='bottom', pady=(0, 6), fill='x')

        # topo com logotipo, nome e subtítulo
        try:
            top_frame = ttk.Frame(self)
            top_frame.place(x=8, y=6)
            # logo interno removido para evitar duplicidade com o caption da janela
            hdr = ttk.Frame(top_frame)
            hdr.pack(side='left')
            # label de título removida: a caption da janela (self.title) será usada no topo
        except Exception:
            pass

    # persistence
    def load_machine_list(self):
        try:
            if os.path.exists(self.machine_json_path):
                with open(self.machine_json_path, 'r', encoding='utf-8') as fh:
                    data = json.load(fh)
                    # backward compat: file could be a list (old format) or dict (new format)
                    if isinstance(data, list):
                        self.machine_list = data
                        self._loaded_meta = {}
                    elif isinstance(data, dict):
                        self.machine_list = data.get('machines', []) if isinstance(data.get('machines', []), list) else []
                        self._loaded_meta = data.get('_meta', {}) or {}
                    else:
                        self.machine_list = []
                        self._loaded_meta = {}
        except Exception:
            self.machine_list = []
            self._loaded_meta = {}
        self.populate_machine_tree()
        # aplicar ordenação carregada (se houver)
        try:
            meta = getattr(self, '_loaded_meta', {}) or {}
            col = meta.get('sort_column')
            rev = bool(meta.get('sort_reverse'))
            if col:
                self._sort_column = col
                self._sort_reverse = rev
                try:
                    # ordenar com a mesma lógica que _sort_by_column
                    self._sort_by_column(col)
                except Exception:
                    pass
                try:
                    self._refresh_sort_indicators()
                except Exception:
                    pass
        except Exception:
            pass
        # iniciar verificação de ping assim que a lista for carregada
        try:
            threading.Thread(target=self._ping_worker, daemon=True).start()
        except Exception:
            pass

    def save_machine_list(self):
        try:
            base = os.path.dirname(self.machine_json_path)
            os.makedirs(base, exist_ok=True)
            # salvar formato novo: objeto com lista e metadados (backward compat handled on load)
            tosave = {
                'machines': self.machine_list,
                '_meta': {
                    'sort_column': getattr(self, '_sort_column', None),
                    'sort_reverse': bool(getattr(self, '_sort_reverse', False)),
                }
            }
            with open(self.machine_json_path, 'w', encoding='utf-8') as fh:
                json.dump(tosave, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def populate_machine_tree(self):
        try:
            self.tree.delete(*self.tree.get_children())
        except Exception:
            pass
        for m in self.machine_list:
            name = (m.get('name') or '').strip()
            alias = (m.get('alias') or '').strip()
            online = bool(m.get('online'))
            tag = 'online' if online else 'offline'
            try:
                self.tree.insert('', 'end', values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'), tags=(tag,))
            except Exception:
                try:
                    self.tree.insert('', 'end', values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'))
                except Exception:
                    pass

    def _sort_by_column(self, col):
        try:
            # determinar chave de ordenação
            if col == 'name':
                keyfunc = lambda m: (m.get('name') or '').strip().upper()
            elif col == 'alias':
                keyfunc = lambda m: (m.get('alias') or '').strip().upper()
            elif col == 'status':
                # ordenar por online: True primeiro quando ascendente
                keyfunc = lambda m: (0 if bool(m.get('online')) else 1, (m.get('name') or '').strip().upper())
            else:
                keyfunc = lambda m: (m.get(col) or '')

            # alternar direção se a mesma coluna for clicada
            if self._sort_column == col:
                self._sort_reverse = not self._sort_reverse
            else:
                self._sort_reverse = False
                self._sort_column = col

            self.machine_list.sort(key=keyfunc, reverse=self._sort_reverse)
            self.populate_machine_tree()
            # atualizar indicadores visuais nos headings e fazer um flash sutil
            try:
                self._refresh_sort_indicators()
            except Exception:
                pass
            try:
                self._flash_sort_indicator(self._sort_column)
            except Exception:
                pass
        except Exception:
            pass

    def _refresh_sort_indicators(self):
        """Atualiza o texto dos headings adicionando ▴ (asc) ou ▾ (desc) na coluna ordenada."""
        try:
            for col, label in self._col_labels.items():
                txt = label
                if self._sort_column == col:
                    # usar setas discretas: ▴ (asc) / ▾ (desc)
                    arrow = '▴' if not self._sort_reverse else '▾'
                    txt = f"{label} {arrow}"
                try:
                    # reatribui o heading com o mesmo command
                    self.tree.heading(col, text=txt, command=lambda c=col: self._sort_by_column(c))
                except Exception:
                    pass
        except Exception:
            pass

    def _flash_sort_indicator(self, col):
        """Mostra uma seta 'grande' temporária para destaque e volta ao indicador discreto."""
        try:
            if not col:
                return
            label = self._col_labels.get(col, col)
            big = '▲' if not self._sort_reverse else '▼'
            try:
                self.tree.heading(col, text=f"{label} {big}", command=lambda c=col: self._sort_by_column(c))
            except Exception:
                pass
            # agendar retorno ao indicador discreto
            try:
                self.after(200, self._refresh_sort_indicators)
            except Exception:
                pass
        except Exception:
            pass

    # CRUD
    def save_selected_or_new_machine(self):
        name = (self.ent_computer.get() or '').strip()
        alias = (self.ent_alias.get() or '').strip()
        if not name:
            messagebox.showwarning('Salvar', 'Informe o nome da máquina antes de salvar')
            return
        if not alias:
            messagebox.showwarning('Salvar', 'Informe um apelido antes de salvar')
            return
        name = name.upper()
        alias = alias.upper()
        existing = next((x for x in self.machine_list if str(x.get('name') or '').strip().upper() == name), None)
        if existing:
            existing['alias'] = alias
        else:
            self.machine_list.append({'name': name, 'alias': alias, 'online': False})
        self.save_machine_list()
        try:
            # limpar saída e manter export desabilitado até que uma nova coleta ocorra
            try:
                self.clear_output()
            except Exception:
                pass
            try:
                self._update_export_button_state()
            except Exception:
                pass
            try:
                if getattr(self, 'chk_open_export_dir', None) is not None:
                    try:
                        self.chk_open_export_dir.configure(state='disabled')
                    except Exception:
                        pass
            except Exception:
                pass
            threading.Thread(target=self._ping_single_and_queue, args=(name,), daemon=True).start()
        except Exception:
            pass
        # atualizar a listagem e selecionar a máquina salva
        try:
            self.populate_machine_tree()
            # selecionar o item recém-salvo
            for child in self.tree.get_children():
                try:
                    vals = self.tree.item(child, 'values')
                    if vals and str(vals[0]).strip().upper() == name:
                        self.tree.selection_set(child)
                        self.tree.see(child)
                        break
                except Exception:
                    continue
        except Exception:
            pass

    def delete_selected_machine(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        name = vals[0] if vals else None
        if not name:
            return
        self.machine_list = [m for m in self.machine_list if str(m.get('name') or '').strip().upper() != str(name).strip().upper()]
        self.save_machine_list()
        # atualizar a listagem e limpar formulário se necessário
        try:
            self.populate_machine_tree()
            # se a máquina deletada estava no formulário, limpar campos
            cur_name = (self.ent_computer.get() or '').strip().upper()
            if cur_name == str(name).strip().upper():
                try:
                    self.ent_computer.delete(0, tk.END)
                    self.ent_alias.delete(0, tk.END)
                except Exception:
                    pass
        except Exception:
            pass

    def new_machine(self):
        # limpa formulário para inserir nova máquina
        try:
            self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass
        try:
            self.ent_computer.delete(0, tk.END)
            self.ent_alias.delete(0, tk.END)
            self.ent_computer.focus_set()
            # desabilitar botão export enquanto não houver coleta
            try:
                self.btn_export.configure(state='disabled')
            except Exception:
                pass
            try:
                if getattr(self, 'chk_open_export_dir', None) is not None:
                    try:
                        self.chk_open_export_dir.configure(state='disabled')
                    except Exception:
                        pass
            except Exception:
                pass
            # limpar progresso mantido
            try:
                self._keep_progress = False
            except Exception:
                pass
            try:
                self.progress['value'] = 0
            except Exception:
                pass
        except Exception:
            pass

    def open_machine_json_folder(self):
        try:
            folder = os.path.dirname(self.machine_json_path)
            if os.path.exists(folder):
                if sys.platform.startswith('win'):
                    os.startfile(folder)
                else:
                    subprocess.run(['xdg-open', folder])
        except Exception:
            pass

    # UI helpers
    def _on_name_keyrelease(self, event=None):
        try:
            v = (self.ent_computer.get() or '').upper()
            pos = self.ent_computer.index(tk.INSERT)
            self.ent_computer.delete(0, tk.END)
            self.ent_computer.insert(0, v)
            try:
                self.ent_computer.icursor(pos)
            except Exception:
                pass
        except Exception:
            pass

    def _on_alias_keyrelease(self, event=None):
        try:
            v = (self.ent_alias.get() or '').upper()
            pos = self.ent_alias.index(tk.INSERT)
            self.ent_alias.delete(0, tk.END)
            self.ent_alias.insert(0, v)
            try:
                self.ent_alias.icursor(pos)
            except Exception:
                pass
        except Exception:
            pass

    # ping helpers
    def _ping_host(self, host):
        if not host:
            return False
        host = host.strip()
        try:
            if sys.platform.startswith('win'):
                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                proc = subprocess.run(['ping', '-n', '1', '-w', '1000', host], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
                return proc.returncode == 0
            else:
                proc = subprocess.run(['ping', '-c', '1', '-W', '1', host], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return proc.returncode == 0
        except Exception:
            return False

    def _ping_single_and_queue(self, name):
        try:
            on = self._ping_host(name)
            status_text = 'ONLINE' if on else 'OFFLINE'
            self.queue.put(('machine_status', name, status_text))
        except Exception:
            pass

    def refresh_machine_status(self):
        """Inicia uma atualização de status (ping) para todas as máquinas.
        Evita reentrância usando a flag self._pinging.
        """
        try:
            if self._pinging or self._processing:
                return
            def _runner():
                try:
                    self._pinging = True
                    # reutiliza a implementação existente do _ping_worker
                    try:
                        self._ping_worker()
                    except Exception:
                        # fallback: ping manualmente
                        for m in list(self.machine_list):
                            name = (m.get('name') or '').strip()
                            try:
                                on = self._ping_host(name)
                                status_text = 'ONLINE' if on else 'OFFLINE'
                                self.queue.put(('machine_status', name, status_text))
                            except Exception:
                                pass
                    # persist após atualização
                    try:
                        self.save_machine_list()
                    except Exception:
                        pass
                finally:
                    self._pinging = False
            threading.Thread(target=_runner, daemon=True).start()
        except Exception:
            pass

    def _ping_worker(self):
        try:
            for m in list(self.machine_list):
                name = (m.get('name') or '').strip()
                try:
                    on = self._ping_host(name)
                    m['online'] = on
                    status_text = 'ONLINE' if on else 'OFFLINE'
                    # enqueue an update for the UI
                    self.queue.put(('machine_status', name, status_text))
                except Exception as ie:
                    # errors are silently ignored for individual hosts
                    pass
            # persist results
            self.save_machine_list()
            # sinalizar conclusão ao loop principal para limpar indicador
            try:
                self.queue.put(('ping_done',))
            except Exception:
                pass
        except Exception as e:
            try:
                self.queue.put(('ping_done',))
            except Exception:
                pass

    def _verify_admin_credentials(self, host, user, passwd, timeout=3):
        r"""Verificação rápida se as credenciais administrativas funcionam contra o host.
        Implementação leve: no Windows tenta um `net use \\\\host\ipc$` com timeout curto.
        Retorna uma tupla: (ok: bool, returncode: int|None, stdout: str, stderr: str).
        Esta função tenta limpar (net use /delete) em caso de sucesso.
        """
        try:
            if not host or not user or not passwd:
                return (False, None, '', '')
            if not sys.platform.startswith('win'):
                # fallback simples: não conseguimos verificar no não-Windows; assume True apenas se houver credenciais
                return (True, 0, '', '')
            # executar net use \\host\ipc$ <passwd> /user:<user>
            target = f"\\\\{host}\\ipc$"
            # limpar conexão existente (se houver) para forçar nova autenticação
            try:
                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                subprocess.run(['net', 'use', target, '/delete'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=2, creationflags=creationflags)
            except Exception:
                pass
            # incluir /persistent:no para evitar mapeamentos persistentes
            cmd = ['net', 'use', target, passwd, f'/user:{user}', '/persistent:no']
            try:
                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=timeout, creationflags=creationflags)
                rc = proc.returncode
                out = proc.stdout or ''
                err = proc.stderr or ''
                ok = rc == 0
            except subprocess.TimeoutExpired:
                return (False, None, '', 'timeout')
            except Exception as e:
                return (False, None, '', str(e))
            # se ok, tentar validar que a sessão realmente autoriza acesso a shares administrativos
            validated = False
            if ok:
                try:
                    # testar admin$ e C$ via PowerShell; se qualquer um listar, consideramos validação
                    shares = [f"\\\\{host}\\admin$", f"\\\\{host}\\C$"]
                    for s in shares:
                            try:
                                # comando PowerShell que retorna 0 em sucesso, 1 em falha
                                ps_cmd = [
                                    'powershell', '-NoProfile', '-Command',
                                    f"Try {{ Get-ChildItem -Path '{s}' -ErrorAction Stop | Out-Null; exit 0 }} Catch {{ exit 1 }}"
                                ]
                                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                                p2 = subprocess.run(ps_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5, creationflags=creationflags)
                                if p2.returncode == 0:
                                    validated = True
                                    break
                            except Exception:
                                continue
                except Exception:
                    validated = False
            # desconectar para não deixar sessão mapeada
            try:
                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                subprocess.run(['net', 'use', target, '/delete'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=2, creationflags=creationflags)
            except Exception:
                pass
            # considerar como ok somente se a validação adicional passar
            if ok and not validated:
                ok = False
                # ajustar mensagem de erro se necessário
                if not err:
                    err = 'verificacao_de_share_falhou'
            return (ok, rc, out, err)
        except Exception as e:
            return (False, None, '', str(e))

    # collection
    def start_collection(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        computer = (self.ent_computer.get() or '').strip() or None
        alias = (self.ent_alias.get() or '').strip() or None

        self.clear_output()
        self._set_controls_state('disabled')
        try:
            # garantir radiobuttons/estado de export desabilitados enquanto processa
            self._update_export_button_state()
        except Exception:
            pass
        self._processing = True
        try:
            self.protocol('WM_DELETE_WINDOW', self._on_close_attempt)
        except Exception:
            pass
        self.btn_start.configure(text='⏳ Processando...')
        self.lbl_progress.configure(text='Iniciando...')
        self.progress['value'] = 0

        def barra_callback(percent_or_none, line_or_stage):
            try:
                if percent_or_none is None:
                    self.queue.put(('line', str(line_or_stage)))
                else:
                    try:
                        p = int(percent_or_none)
                    except Exception:
                        p = 0
                    self.queue.put(('progress', p, str(line_or_stage)))
            except Exception:
                pass

        def worker():
            try:
                user = (self.ent_user.get() or '').strip()
                passwd = (self.ent_pass.get() or '')
                try:
                    self._last_collection_computer = computer
                except Exception:
                    self._last_collection_computer = None
                if csinfo and user and passwd:
                    try:
                        csinfo.set_default_credential(user, passwd)
                    except Exception:
                        pass
                try:
                    if csinfo:
                        # Não pedir export automático ao backend aqui; o usuário deve clicar em Exportar
                        csinfo.main(barra_callback=barra_callback,
                                    computer_name=computer,
                                    machine_alias=alias)
                    else:
                        for i in range(0, 101, 10):
                            barra_callback(i, f"Simulando etapa {i}")
                            threading.Event().wait(0.05)
                        barra_callback(None, 'Linha simulada 1')
                        barra_callback(None, 'Linha simulada 2')
                    self.queue.put(('done', None))
                finally:
                    if csinfo:
                        try:
                            csinfo.clear_default_credential()
                        except Exception:
                            pass
            except Exception as e:
                self.queue.put(('error', str(e)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    # queue processing
    def _append_output(self, text):
        try:
            self.txt_output.configure(state='normal')
            self.txt_output.insert(tk.END, text + '\n')
            self.txt_output.see(tk.END)
            self.txt_output.configure(state='disabled')
        except Exception:
            pass

    def _process_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if not item:
                    continue
                kind = item[0]
                if kind == 'line':
                    line = item[1]
                    try:
                        self.last_lines.append(str(line))
                    except Exception:
                        pass
                    self.txt_output.configure(state='normal')
                    try:
                        normalized = str(line).replace('\r\n', '\n').replace('\r', '\n')
                        self.txt_output.insert(tk.END, normalized + '\n')
                        self.txt_output.see(tk.END)
                    except Exception:
                        try:
                            self.txt_output.insert(tk.END, str(line) + '\n')
                        except Exception:
                            pass
                    self.txt_output.configure(state='disabled')
                elif kind == 'progress':
                    pct = item[1]
                    stage = item[2]
                    try:
                        self.progress['value'] = int(pct)
                    except Exception:
                        pass
                    try:
                        self.lbl_progress.configure(text=str(stage))
                    except Exception:
                        pass
                elif kind == 'machine_status':
                    key = item[1]
                    status = item[2]
                    status_text = 'ONLINE' if str(status).strip().upper() == 'ONLINE' else 'OFFLINE'
                    # update tree if present
                    for child in self.tree.get_children():
                        try:
                            v = self.tree.item(child, 'values')
                            if v and (str(v[0]).strip().upper() == str(key).strip().upper() or str(v[1]).strip().upper() == str(key).strip().upper()):
                                try:
                                    self.tree.item(child, values=(v[0], v[1], status_text), tags=('online' if status_text == 'ONLINE' else 'offline',))
                                except Exception:
                                    pass
                                break
                        except Exception:
                            continue
                    try:
                        for m in self.machine_list:
                            if str(m.get('name') or '').strip().upper() == str(key).strip().upper():
                                m['online'] = (status_text == 'ONLINE')
                                break
                    except Exception:
                        pass
                    # não escrever logs de ping no painel para manter interface limpa
                    pass
                elif kind == 'ping_done':
                    # limpar indicador visual de atualização
                    try:
                        self.lbl_ping_status.configure(text='')
                    except Exception:
                        pass
                elif kind == 'done':
                    self._processing = False
                    self._set_controls_state('normal')
                    self.btn_start.configure(text='Coletar (F3)')
                    self.lbl_progress.configure(text='Pronto')
                    self.progress['value'] = 100
                    # manter barra preenchida até que o usuário selecione outra máquina
                    try:
                        self._keep_progress = True
                    except Exception:
                        pass
                    # atualizar estado dos controles de export (radiobuttons + botão)
                    try:
                        self._update_export_button_state()
                    except Exception:
                        # fallback: habilitar botão de export se houver linhas
                        try:
                            if self.last_lines:
                                self.btn_export.configure(state='normal')
                        except Exception:
                            pass
                    # habilitar checkbox de abrir diretório automático após coleta bem-sucedida
                    try:
                        if getattr(self, 'chk_open_export_dir', None) is not None:
                            try:
                                if self.last_lines:
                                    self.chk_open_export_dir.configure(state='normal')
                                else:
                                    self.chk_open_export_dir.configure(state='disabled')
                            except Exception:
                                pass
                    except Exception:
                        pass
                elif kind == 'error':
                    self._processing = False
                    self._set_controls_state('normal')
                    self.btn_start.configure(text='Coletar (F3)')
                    self.lbl_progress.configure(text='Erro')
                    messagebox.showerror('Erro', str(item[1]))
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            if not self._processing:
                # somente resetar se não estivermos mantendo o progresso pós-coleta
                if not getattr(self, '_keep_progress', False):
                    self.progress['value'] = 0
            self.after(100, self._process_queue)

    def _set_controls_state(self, state='normal'):
        # incluir botões que devem ser desabilitados durante processamento
        widgets = [
            self.ent_computer, self.ent_alias, self.ent_user, self.ent_pass,
            getattr(self, 'export_frame', None), self.btn_save, self.btn_delete, self.btn_export,
            getattr(self, 'chk_open_export_dir', None),
            self.btn_start, getattr(self, 'btn_new', None), getattr(self, 'btn_open_folder', None), getattr(self, 'btn_refresh', None), getattr(self, 'tree', None)
        ]
        for w in widgets:
            try:
                if state == 'disabled':
                    try:
                        w.configure(state='disabled')
                    except Exception:
                        pass
                else:
                    try:
                        w.configure(state='normal')
                    except Exception:
                        pass
            except Exception:
                pass
        # alterar cursor da janela para indicar espera durante processamento
        try:
            if state == 'disabled':
                try:
                    self.configure(cursor='watch')
                except Exception:
                    try:
                        self.configure(cursor='wait')
                    except Exception:
                        pass
            else:
                try:
                    self.configure(cursor='')
                except Exception:
                    pass
            try:
                # forçar atualização visual imediata
                self.update_idletasks()
            except Exception:
                pass
        except Exception:
            pass

        # bloquear cabeçalhos da tree (comandos de ordenação) para evitar interação
        try:
            if getattr(self, 'tree', None) is not None:
                if state == 'disabled':
                    try:
                        self._disable_tree_headings()
                    except Exception:
                        pass
                else:
                    try:
                        self._restore_tree_headings()
                    except Exception:
                        pass
        except Exception:
            pass

    def _disable_tree_headings(self):
        """Substitui temporariamente os comandos dos headings por no-ops para bloquear interação."""
        try:
            # armazenar comandos atuais para restauração
            self._saved_heading_cmds = getattr(self, '_saved_heading_cmds', {})
            for col in list(self._col_labels.keys()):
                try:
                    # salvar estado
                    try:
                        cmd = self.tree.heading(col, option='command')
                    except Exception:
                        cmd = None
                    self._saved_heading_cmds[col] = cmd
                    # atribuir no-op
                    self.tree.heading(col, command=lambda: None)
                except Exception:
                    pass
        except Exception:
            pass

    def _restore_tree_headings(self):
        """Restaura os comandos salvos dos headings e reaplica indicadores de ordenação."""
        try:
            saved = getattr(self, '_saved_heading_cmds', {}) or {}
            for col, label in self._col_labels.items():
                try:
                    cmd = saved.get(col)
                    if cmd:
                        self.tree.heading(col, text=self.tree.heading(col, 'text'), command=cmd)
                    else:
                        # restaurar via helper padrão
                        self.tree.heading(col, text=self.tree.heading(col, 'text'), command=lambda c=col: self._sort_by_column(c))
                except Exception:
                    pass
            try:
                # atualizar indicadores visuais
                self._refresh_sort_indicators()
            except Exception:
                pass
        except Exception:
            pass

    def clear_output(self):
        try:
            self.txt_output.configure(state='normal')
            self.txt_output.delete('1.0', tk.END)
            self.txt_output.configure(state='disabled')
            self.last_lines = []
            self.btn_export.configure(state='disabled')
        except Exception:
            pass

    def _update_export_button_state(self):
        fmt = self.export_var.get() if getattr(self, 'export_var', None) is not None else 'pdf'
        enabled = (bool(self.last_lines) and not self._processing)
        try:
            if enabled:
                self.btn_export.configure(state='normal')
            else:
                self.btn_export.configure(state='disabled')
        except Exception:
            pass
        # habilitar/desabilitar radiobuttons conforme disponibilidade
        try:
            state = 'normal' if enabled else 'disabled'
            try:
                self.rb_pdf.configure(state=state)
            except Exception:
                pass
            try:
                self.rb_txt.configure(state=state)
            except Exception:
                pass
            try:
                self.rb_both.configure(state=state)
            except Exception:
                pass
        except Exception:
            pass

    def _do_export(self):
        if not self.last_lines:
            messagebox.showinfo('Exportar', 'Nenhum dado para exportar')
            return
        fmt = self.export_var.get() if getattr(self, 'export_var', None) is not None else 'pdf'
        alias = (self.ent_alias.get() or '').strip() or None
        comp = (self._last_collection_computer or (self.ent_computer.get() or '').strip() or '')

        def _safe(s):
            try:
                if csinfo and hasattr(csinfo, 'safe_filename'):
                    return csinfo.safe_filename(s)
            except Exception:
                pass
            return re.sub(r'[^0-9A-Za-z_.-]+', '_', str(s or ''))

        base = f"Info_maquina_{_safe(alias) + '_' if alias else ''}{_safe(comp)}_v{__version__}"
        base_cwd = os.getcwd()
        # definir pastas de saída por tipo
        pdf_folder = os.path.join(base_cwd, 'Relatorio', 'PDF')
        txt_folder = os.path.join(base_cwd, 'Relatorio', 'TXT')
        try:
            os.makedirs(pdf_folder, exist_ok=True)
        except Exception:
            pass
        try:
            os.makedirs(txt_folder, exist_ok=True)
        except Exception:
            pass
        try:
            if fmt in ('txt', 'ambos'):
                p = os.path.join(txt_folder, base + '.txt')
                try:
                    if csinfo and hasattr(csinfo, 'write_report'):
                        try:
                            csinfo.write_report(p, self.last_lines)
                        except TypeError:
                            csinfo.write_report(p, self.last_lines)
                    else:
                        with open(p, 'w', encoding='utf-8') as fh:
                            fh.write('\n'.join(self.last_lines))
                    self._append_output(f'Exportado TXT: {p}')
                except Exception as e:
                    messagebox.showerror('Exportar', f'Erro ao escrever TXT: {e}')
                    return
            if fmt in ('pdf', 'ambos'):
                p = os.path.join(pdf_folder, base + '.pdf')
                try:
                    if csinfo and hasattr(csinfo, 'write_pdf_report'):
                        try:
                            # propagar versão do front-end para o backend para que o cabeçalho PDF mostre a versão correta
                            try:
                                csinfo.__version__ = __version__
                            except Exception:
                                pass
                            # também informar logo e nome do app
                            try:
                                csinfo.__logo_path__ = APP_LOGO
                            except Exception:
                                pass
                            try:
                                csinfo.__app_name__ = 'CSInfo'
                            except Exception:
                                pass
                            csinfo.write_pdf_report(p, self.last_lines, comp)
                        except TypeError:
                            try:
                                csinfo.__version__ = __version__
                            except Exception:
                                pass
                            try:
                                csinfo.__logo_path__ = APP_LOGO
                            except Exception:
                                pass
                            try:
                                csinfo.__app_name__ = 'CSInfo'
                            except Exception:
                                pass
                            csinfo.write_pdf_report(p, self.last_lines)
                    else:
                        with open(p, 'w', encoding='utf-8') as fh:
                            fh.write('\n'.join(self.last_lines))
                    self._append_output(f'Exportado PDF: {p}')
                except Exception as e:
                    messagebox.showerror('Exportar', f'Erro ao escrever PDF: {e}')
                    return
            # se opção de abrir pasta estiver marcada, abrir a pasta onde o(s) arquivo(s) foram salvos
            try:
                if getattr(self, 'open_export_dir_var', None) and self.open_export_dir_var.get():
                    try:
                        # abrir a pasta apropriada dependendo do formato selecionado
                        if fmt == 'pdf':
                            os.startfile(pdf_folder)
                        elif fmt == 'txt':
                            os.startfile(txt_folder)
                        else:
                            # ambos: abrir a pasta raiz Relatorio
                            os.startfile(os.path.join(base_cwd, 'Relatorio'))
                    except Exception:
                        try:
                            # fallback para explorer
                            if fmt == 'pdf':
                                subprocess.Popen(['explorer', pdf_folder])
                            elif fmt == 'txt':
                                subprocess.Popen(['explorer', txt_folder])
                            else:
                                subprocess.Popen(['explorer', os.path.join(base_cwd, 'Relatorio')])
                        except Exception:
                            pass
            except Exception:
                pass
            messagebox.showinfo('Exportar', 'Exportação concluída')
        except Exception as e:
            messagebox.showerror('Exportar', str(e))

    def _on_close_attempt(self):
        if self._processing:
            messagebox.showwarning('Aguarde', 'Processamento em andamento. Aguarde o fim antes de fechar.')
            return
        try:
            self.destroy()
        except Exception:
            try:
                self.quit()
            except Exception:
                pass

    def _load_selection_into_form(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        if vals:
            try:
                # limpar saída (resultados de coleta anteriores) e desabilitar export
                try:
                    self.clear_output()
                except Exception:
                    pass
                self.ent_computer.delete(0, tk.END)
                self.ent_computer.insert(0, vals[0])
                self.ent_alias.delete(0, tk.END)
                self.ent_alias.insert(0, vals[1] if len(vals) > 1 else '')
            except Exception:
                pass

    def _on_tree_selection_change(self):
        # chamado quando a seleção muda: limpar saída e desabilitar export
        try:
            self.clear_output()
            try:
                self._update_export_button_state()
            except Exception:
                pass
            try:
                # seleção nova: limpar progresso mantido
                try:
                    self._keep_progress = False
                except Exception:
                    pass
                try:
                    self.progress['value'] = 0
                except Exception:
                    pass
                try:
                    self.lbl_progress.configure(text='Pronto')
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass

    def _on_tree_right_click(self, event):
        try:
            # identificar item sob o cursor
            iid = self.tree.identify_row(event.y)
            if not iid:
                return
            # selecionar o item clicado
            try:
                self.tree.selection_set(iid)
            except Exception:
                pass
            vals = self.tree.item(iid, 'values')
            if not vals:
                return
            status = (vals[2] if len(vals) > 2 else '').strip().upper()
            # ativar menu somente se ONLINE e se as credenciais administrativas forem validadas
            try:
                user = (self.ent_user.get() or '').strip()
                passwd = (self.ent_pass.get() or '')
            except Exception:
                user = ''
                passwd = ''

            # estado inicial: desabilitar ações por padrão
            try:
                self.tree_menu.entryconfigure('Reiniciar', state='disabled')
            except Exception:
                pass
            try:
                self.tree_menu.entryconfigure('Desligar', state='disabled')
            except Exception:
                pass

            # se não estiver ONLINE, mantém ambos desabilitados
            if status != 'ONLINE':
                pass
            else:
                # se houver credenciais preenchidas, realizar verificação assíncrona para habilitar ações
                if user and passwd:
                    host_to_check = (vals[0] or '').strip().upper()
                    def _verify_and_enable_both():
                        try:
                            result = self._verify_admin_credentials(host_to_check, user, passwd)
                            try:
                                ok, rc, out, err = (result if isinstance(result, tuple) else (bool(result), None, '', ''))
                            except Exception:
                                ok, rc, out, err = (False, None, '', '')
                            try:
                                # somente aplicar o resultado se ainda estivermos com a mesma seleção
                                try:
                                    cur_sel = self.tree.selection()
                                    if not cur_sel:
                                        return
                                    cur_vals = self.tree.item(cur_sel[0], 'values')
                                    cur_name = (cur_vals[0] if cur_vals else '')
                                    if not cur_name or cur_name.strip().upper() != host_to_check:
                                        return
                                    # também garantir que os campos de credenciais não mudaram durante a verificação
                                    try:
                                        cur_user = (self.ent_user.get() or '').strip()
                                        cur_pass = (self.ent_pass.get() or '')
                                        if cur_user != user or cur_pass != passwd:
                                            # credenciais alteradas enquanto verificávamos: abortar
                                            return
                                    except Exception:
                                        return
                                except Exception:
                                    return
                                state = 'normal' if ok else 'disabled'
                                try:
                                    self.tree_menu.entryconfigure('Reiniciar', state=state)
                                except Exception:
                                    pass
                                try:
                                    self.tree_menu.entryconfigure('Desligar', state=state)
                                except Exception:
                                    pass
                                # log detalhado do resultado para depuração
                                try:
                                    if ok:
                                        self._append_output(f'Credenciais confirmadas para {host_to_check} (rc={rc})')
                                    else:
                                        self._append_output(f'Falha na verificação para {host_to_check} (rc={rc}) stderr={err.strip()}')
                                except Exception:
                                    pass
                            except Exception:
                                pass
                        except Exception:
                            pass
                    try:
                        threading.Thread(target=_verify_and_enable_both, daemon=True).start()
                    except Exception:
                        pass
                else:
                    # sem credenciais, mantém ambos desabilitados
                    pass
            try:
                # mostrar o menu no ponto do clique
                self.tree_menu.tk_popup(event.x_root, event.y_root)
            except Exception:
                pass
            finally:
                try:
                    self.tree_menu.grab_release()
                except Exception:
                    pass
        except Exception:
            pass

    def _on_context_restart(self):
        # obtém a seleção e inicia reinício em thread
        try:
            sel = self.tree.selection()
            if not sel:
                return
            vals = self.tree.item(sel[0], 'values')
            if not vals:
                return
            name = (vals[0] or '').strip()
            if not name:
                return
            # (aviso informativo removido por solicitação) — mantém apenas a confirmação final
            # confirmar ação
            if not messagebox.askyesno('Reiniciar', f'Deseja reiniciar a máquina "{name}" agora?'):
                return
            # executar comando em thread para não bloquear a GUI
            def _runner():
                try:
                    cmd = ['shutdown', '/r', '/m', f'\\{name}', '/t', '0', '/f']
                    try:
                        self._append_output(f'Executando: {" ".join(cmd)}')
                    except Exception:
                        pass
                    # chamar subprocess
                    try:
                        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=creationflags)
                        out = proc.stdout or ''
                        err = proc.stderr or ''
                        if out:
                            try:
                                self._append_output(out.strip())
                            except Exception:
                                pass
                        if err:
                            try:
                                self._append_output(f'Erro: {err.strip()}')
                            except Exception:
                                pass
                        try:
                            messagebox.showinfo('Reiniciar', f'Comando enviado para {name}.')
                        except Exception:
                            pass
                    except Exception as e:
                        try:
                            self._append_output(f'Erro ao executar reinício: {e}')
                        except Exception:
                            pass
                except Exception:
                    pass
            try:
                threading.Thread(target=_runner, daemon=True).start()
            except Exception:
                pass
        except Exception:
            pass

    def _on_context_shutdown(self):
        # obtém a seleção e inicia desligamento em thread
        try:
            sel = self.tree.selection()
            if not sel:
                return
            vals = self.tree.item(sel[0], 'values')
            if not vals:
                return
            name = (vals[0] or '').strip()
            if not name:
                return
            # confirmar ação
            if not messagebox.askyesno('Desligar', f'Deseja desligar a máquina "{name}" agora?'):
                return
            # executar comando em thread para não bloquear a GUI
            def _runner_shutdown():
                try:
                    cmd = ['shutdown', '/s', '/m', f'\\{name}', '/t', '0', '/f']
                    try:
                        self._append_output(f'Executando: {" ".join(cmd)}')
                    except Exception:
                        pass
                    try:
                        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=creationflags)
                        out = proc.stdout or ''
                        err = proc.stderr or ''
                        if out:
                            try:
                                self._append_output(out.strip())
                            except Exception:
                                pass
                        if err:
                            try:
                                self._append_output(f'Erro: {err.strip()}')
                            except Exception:
                                pass
                        try:
                            messagebox.showinfo('Desligar', f'Comando enviado para {name}.')
                        except Exception:
                            pass
                    except Exception as e:
                        try:
                            self._append_output(f'Erro ao executar desligamento: {e}')
                        except Exception:
                            pass
                except Exception:
                    pass
            try:
                threading.Thread(target=_runner_shutdown, daemon=True).start()
            except Exception:
                pass
        except Exception:
            pass


def main():
    app = CSInfoGUI()
    app.mainloop()


if __name__ == '__main__':
    main()


