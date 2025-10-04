
"""csinfo GUI

Interface Tkinter que chama o backend `csinfo.main(...)` em background e recebe
eventos via callback para atualizar progresso e saída textual.

Funcionalidades principais:
- Chama o backend real (se disponível) e streama linhas/progressos para a UI.
- Bloqueia controles e impede fechar a janela enquanto houver processamento.
- Persiste lista de máquinas em %APPDATA%\\CSInfo\\machines_history.json.
- Faz ping nas máquinas para marcar ONLINE/OFFLINE e colore as linhas.
- Normaliza entradas "Máquina" e "Apelido" para MAIÚSCULAS e centraliza os campos.
- Campos de credenciais (usuário/senha) são passados ao backend via helpers.
- Botão "Exportar" habilitado apenas após coleta concluída e formato selecionado.

"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import re

try:
    import csinfo
except Exception:
    csinfo = None


def get_appdata_path():
    # Retorna a pasta de appdata do usuário no Windows, senão usa home
    if sys.platform.startswith('win'):
        return os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.expanduser('~')


class CSInfoGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CSInfo GUI')
        self.geometry('980x640')

        # state
        self.queue = queue.Queue()
        self.worker_thread = None
        self._processing = False
        self.last_lines = []
        self._last_collection_computer = None
        self.machine_list = []  # list of dicts: {name, alias, online}
        self.machine_json_path = self._get_machine_json_path()
        self.debug = bool(os.environ.get('CSINFO_GUI_DEBUG'))

        # build UI
        self._build_ui()

        # load persisted machines
        self.load_machine_list()

        # start queue processor
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

        # credentials
        ttk.Label(left, text='Usuário (opcional):').pack(anchor='w', pady=(8, 0))
        self.ent_user = ttk.Entry(left, width=32)
        self.ent_user.pack()
        ttk.Label(left, text='Senha (opcional):').pack(anchor='w', pady=(8, 0))
        self.ent_pass = ttk.Entry(left, width=32, show='*')
        self.ent_pass.pack()

        # save / delete
        btn_fr = ttk.Frame(left)
        btn_fr.pack(pady=8, fill='x')
        self.btn_save = ttk.Button(btn_fr, text='Salvar', command=self.save_selected_or_new_machine)
        self.btn_save.pack(side='left', fill='x', expand=True)
        self.btn_delete = ttk.Button(btn_fr, text='Excluir', command=self.delete_selected_machine)
        self.btn_delete.pack(side='left', fill='x', expand=True, padx=(6, 0))

        ttk.Label(left, text='Exportar como:').pack(anchor='w', pady=(8, 0))
        self.cmb_export = ttk.Combobox(left, values=('nenhum', 'txt', 'pdf', 'ambos'), state='readonly')
        self.cmb_export.current(0)
        self.cmb_export.pack()
        self.cmb_export.bind('<<ComboboxSelected>>', lambda e: self._update_export_button_state())

        self.btn_export = ttk.Button(left, text='Exportar', command=self._do_export, state='disabled')
        self.btn_export.pack(fill='x', pady=(8, 0))

        self.btn_start = ttk.Button(left, text='Coletar', command=self.start_collection)
        self.btn_start.pack(fill='x', pady=(12, 0))

        ttk.Button(left, text='Abrir pasta de máquinas', command=self.open_machine_json_folder).pack(fill='x', pady=(8, 0))

        # middle: machine list
        mid = ttk.Frame(frm)
        mid.pack(side='left', fill='both', expand=True, padx=(8, 8))
        self.tree = ttk.Treeview(mid, columns=('name', 'alias', 'status'), show='headings', selectmode='browse')
        self.tree.heading('name', text='Máquina')
        self.tree.heading('alias', text='Apelido')
        self.tree.heading('status', text='Estado')
        self.tree.column('name', width=200, anchor='center')
        self.tree.column('alias', width=160, anchor='center')
        self.tree.column('status', width=90, anchor='center')
        self.tree.pack(fill='both', expand=True)

        # tag styling (backgrounds)
        try:
            self.tree.tag_configure('online', background='#e6ffed')
            self.tree.tag_configure('offline', background='#ffe6e6')
        except Exception:
            pass

        # right: output + progress
        right = ttk.Frame(frm)
        right.pack(side='right', fill='both', expand=True)
        # wrap='word' ensures lines wrap to the visible width (dynamic with resize)
        self.txt_output = ScrolledText(right, height=20, state='disabled', wrap='word')
        self.txt_output.pack(fill='both', expand=True)

        bar_fr = ttk.Frame(right)
        bar_fr.pack(fill='x')
        self.progress = ttk.Progressbar(bar_fr, orient='horizontal', mode='determinate')
        self.progress.pack(fill='x', side='left', expand=True, padx=(0, 8))
        self.lbl_progress = ttk.Label(bar_fr, text='Pronto')
        self.lbl_progress.pack(side='right')

        # attach double-click on tree to load selected
        self.tree.bind('<Double-1>', lambda e: self._load_selection_into_form())

    # ---------------- persistence ----------------
    def load_machine_list(self):
        try:
            if os.path.exists(self.machine_json_path):
                with open(self.machine_json_path, 'r', encoding='utf-8') as fh:
                    data = json.load(fh)
                    if isinstance(data, list):
                        self.machine_list = data
        except Exception:
            self.machine_list = []
        self.populate_machine_tree()

    def save_machine_list(self):
        try:
            base = os.path.dirname(self.machine_json_path)
            os.makedirs(base, exist_ok=True)
            with open(self.machine_json_path, 'w', encoding='utf-8') as fh:
                json.dump(self.machine_list, fh, ensure_ascii=False, indent=2)
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
            iid = name or alias
            try:
                self.tree.insert('', 'end', iid, values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'), tags=(tag,))
            except Exception:
                # fallback: insert without iid
                try:
                    self.tree.insert('', 'end', values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'), tags=(tag,))
                except Exception:
                    pass

    # ---------------- machine CRUD ----------------
    def _on_name_keyrelease(self, event=None):
        v = (self.ent_computer.get() or '').upper()
        pos = self.ent_computer.index(tk.INSERT)
        self.ent_computer.delete(0, tk.END)
        self.ent_computer.insert(0, v)
        try:
            self.ent_computer.icursor(pos)
        except Exception:
            pass

    def _on_alias_keyrelease(self, event=None):
        v = (self.ent_alias.get() or '').upper()
        pos = self.ent_alias.index(tk.INSERT)
        self.ent_alias.delete(0, tk.END)
        self.ent_alias.insert(0, v)
        try:
            self.ent_alias.icursor(pos)
        except Exception:
            pass

    def save_selected_or_new_machine(self):
        name = (self.ent_computer.get() or '').strip()
        alias = (self.ent_alias.get() or '').strip()
        if not name:
            self._append_output('Nome da máquina é obrigatório para salvar.')
            return
        if not alias:
            self._append_output('Apelido é obrigatório para salvar.')
            return
        existing = next((x for x in self.machine_list if str(x.get('name') or '').strip().upper() == name.upper()), None)
        if existing:
            existing['alias'] = alias
        else:
            self.machine_list.append({'name': name, 'alias': alias, 'online': False})
        self.save_machine_list()
        self.populate_machine_tree()
        # quick ping this host
        try:
            threading.Thread(target=self._ping_single_and_queue, args=(name,), daemon=True).start()
        except Exception:
            pass
        self._append_output(f"Máquina '{name}' salva.")

    def delete_selected_machine(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        # find by name (case-insensitive)
        name = iid
        self.machine_list = [m for m in self.machine_list if str(m.get('name') or '').strip().upper() != str(name).strip().upper()]
        self.save_machine_list()
        self.populate_machine_tree()
        self._append_output(f"Máquina '{name}' excluída.")

    def open_machine_json_folder(self):
        p = os.path.abspath(self.machine_json_path)
        folder = os.path.dirname(p)
        try:
            if sys.platform.startswith('win'):
                os.startfile(folder)
            else:
                subprocess.run(['xdg-open', folder])
        except Exception as e:
            self._append_output(f"Falha ao abrir pasta: {e}")

    def refresh_machine_statuses(self):
        thr = threading.Thread(target=self._ping_worker, daemon=True)
        thr.start()

    def _ping_host(self, host):
        if not host:
            return False
        if sys.platform.startswith('win'):
            cmd = ['ping', '-n', '1', '-w', '1000', host]
        else:
            cmd = ['ping', '-c', '1', '-W', '1', host]
        try:
            res = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return res.returncode == 0
        except Exception:
            return False

    def _ping_worker(self):
        for m in list(self.machine_list):
            name = (m.get('name') or '').strip()
            on = self._ping_host(name)
            m['online'] = on
            status_text = 'ONLINE' if on else 'OFFLINE'
            # push by name so tree matching tries both iid and first-column match
            self.queue.put(('machine_status', name, status_text))
        self.save_machine_list()

    def _ping_single_and_queue(self, name):
        try:
            on = self._ping_host(name)
            status_text = 'ONLINE' if on else 'OFFLINE'
            self.queue.put(('machine_status', name, status_text))
        except Exception:
            pass

    # ---------------- collection / backend ----------------
    def start_collection(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        computer = self.ent_computer.get().strip() or None
        alias = self.ent_alias.get().strip() or None
        export = self.cmb_export.get()

        # clear output
        self.clear_output()

        # lock UI
        self._set_controls_state('disabled')
        self._processing = True
        # prevent window close
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
                        perc = int(percent_or_none)
                    except Exception:
                        perc = 0
                    self.queue.put(('progress', perc, str(line_or_stage)))
            except Exception:
                pass

        def worker():
            try:
                user = (self.ent_user.get() or '').strip()
                passwd = (self.ent_pass.get() or '')
                if user and passwd and csinfo:
                    try:
                        csinfo.set_default_credential(user, passwd)
                    except Exception:
                        pass
                try:
                    # record the computer used for this collection so exports (PDF) can include it
                    try:
                        self._last_collection_computer = computer
                    except Exception:
                        self._last_collection_computer = None

                    if csinfo:
                        csinfo.main(export_type=(export if export != 'nenhum' else None), barra_callback=barra_callback, computer_name=computer, machine_alias=alias)
                    else:
                        # simulate
                        for i in range(0, 101, 10):
                            barra_callback(i, f"Simulando etapa {i}")
                            threading.Event().wait(0.08)
                        barra_callback(None, 'Simulated line 1')
                        barra_callback(None, 'Simulated line 2')
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

    # ---------------- queue processing ----------------
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

                elif kind == 'machine_status':
                    iid = item[1]
                    status = item[2]
                    if status in ('🟢', '🔴'):
                        status_text = 'ONLINE' if status == '🟢' else 'OFFLINE'
                    else:
                        status_text = str(status).strip().upper()

                    vals = None
                    try:
                        vals = self.tree.item(iid, 'values')
                    except Exception:
                        vals = None

                    if not vals:
                        found = None
                        for child in self.tree.get_children():
                            v = self.tree.item(child, 'values')
                            if self.debug:
                                self._append_output(f"[debug] checking child={child} values={v}")
                            if v and str(v[0]).strip().upper() == str(iid).strip().upper():
                                found = child
                                vals = v
                                break
                        if found:
                            iid = found

                    if vals:
                        try:
                            self.tree.set(iid, 'status', status_text)
                            tag = 'online' if status_text == 'ONLINE' else 'offline'
                            try:
                                self.tree.item(iid, tags=(tag,))
                            except Exception:
                                pass
                            try:
                                self.update_idletasks()
                            except Exception:
                                pass
                        except Exception:
                            try:
                                self.tree.item(iid, values=(vals[0], vals[1], status_text))
                            except Exception:
                                pass

                    try:
                        name_key = str((vals[0] if vals else iid)).strip()
                        for m in self.machine_list:
                            if str(m.get('name') or '').strip().upper() == name_key.upper():
                                m['online'] = (status_text == 'ONLINE')
                                break
                    except Exception:
                        pass

                    try:
                        self._append_output(f"[ping] {name_key} -> {status_text.lower()}")
                    except Exception:
                        pass

                elif kind == 'progress':
                    perc = item[1]
                    etapa = item[2]
                    try:
                        self.progress['value'] = perc
                        self.lbl_progress.configure(text=f"{etapa} — {perc}%")
                    except Exception:
                        pass

                elif kind == 'done':
                    self.lbl_progress.configure(text='Concluído')
                    self.progress['value'] = 100
                    self.btn_start.configure(text='Coletar')
                    self._set_controls_state('normal')
                    self._processing = False
                    try:
                        self.protocol('WM_DELETE_WINDOW', self.destroy)
                    except Exception:
                        pass
                    self._update_export_button_state()

                elif kind == 'error':
                    msg = item[1]
                    self._append_output(f"ERRO: {msg}")
                    self.lbl_progress.configure(text='Erro')
                    self.btn_start.configure(text='Coletar')
                    self._set_controls_state('normal')
                    self._processing = False
                    try:
                        self.protocol('WM_DELETE_WINDOW', self.destroy)
                    except Exception:
                        pass

        except queue.Empty:
            pass
        # re-schedule
        self.after(100, self._process_queue)

    # ---------------- UI helpers ----------------
    def _set_controls_state(self, state='normal'):
        widgets = [self.ent_computer, self.ent_alias, self.ent_user, self.ent_pass, self.btn_save, self.btn_delete, self.cmb_export, self.btn_start]
        for w in widgets:
            try:
                if state == 'disabled':
                    w.state(['disabled']) if isinstance(w, ttk.Widget) else w.configure(state='disabled')
                else:
                    # enable
                    try:
                        w.state(['!disabled'])
                    except Exception:
                        try:
                            w.configure(state='normal')
                        except Exception:
                            pass
            except Exception:
                try:
                    w.configure(state=state)
                except Exception:
                    pass

    def clear_output(self):
        try:
            self.txt_output.configure(state='normal')
            self.txt_output.delete('1.0', tk.END)
            self.txt_output.configure(state='disabled')
        except Exception:
            pass

    def _on_close_attempt(self):
        if getattr(self, '_processing', False):
            self._append_output('A coleta está em execução. Aguarde a conclusão antes de fechar a janela.')
            return
        try:
            self.destroy()
        except Exception:
            pass

    # ---------------- export ----------------
    def _update_export_button_state(self):
        fmt = self.cmb_export.get()
        enabled = (fmt != 'nenhum' and bool(self.last_lines) and not self._processing)
        try:
            if enabled:
                self.btn_export.configure(state='normal')
            else:
                self.btn_export.configure(state='disabled')
        except Exception:
            pass

    def _do_export(self):
        fmt = self.cmb_export.get()
        if fmt == 'nenhum':
            return
        if not self.last_lines:
            self._append_output('Nada para exportar. Execute a coleta primeiro.')
            return
        # construir nome de arquivo seguindo padrão do backend: Info_maquina_<apelido>_<nomemaquina>.txt
        folder = os.getcwd()

        def _safe_filename(s):
            try:
                if csinfo and hasattr(csinfo, 'safe_filename'):
                    return csinfo.safe_filename(s)
            except Exception:
                pass
            if not s:
                return ''
            return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(s))

        # usar apelido se fornecido, senão apenas o nome da máquina da última coleta
        alias = (self.ent_alias.get() or '').strip()
        comp = getattr(self, '_last_collection_computer', None) or self.ent_computer.get().strip() or ''
        safe_name = _safe_filename(comp)
        if alias:
            safe_alias = _safe_filename(alias)
            base = f"Info_maquina_{safe_alias}_{safe_name}"
        else:
            base = f"Info_maquina_{safe_name}"
        try:
            if fmt in ('txt', 'ambos'):
                path_txt = os.path.join(folder, base + '.txt')
                if csinfo and hasattr(csinfo, 'write_report'):
                    try:
                        csinfo.write_report(path_txt, self.last_lines)
                    except TypeError:
                        # older/newer signatures: try fallback with 2 args
                        csinfo.write_report(path_txt, self.last_lines)
                else:
                    with open(path_txt, 'w', encoding='utf-8') as fh:
                        fh.write('\n'.join(self.last_lines))
                self._append_output(f"Exportado TXT: {path_txt}")
            if fmt in ('pdf', 'ambos'):
                path_pdf = os.path.join(folder, base + '.pdf')
                if csinfo and hasattr(csinfo, 'write_pdf_report'):
                    comp_arg = getattr(self, '_last_collection_computer', None) or comp
                    csinfo.write_pdf_report(path_pdf, self.last_lines, comp_arg)
                else:
                    with open(path_pdf, 'w', encoding='utf-8') as fh:
                        fh.write('\n'.join(self.last_lines))
                self._append_output(f"Exportado PDF: {path_pdf}")
        except Exception as e:
            self._append_output(f"Falha ao exportar: {e}")
            return

        try:
            if sys.platform.startswith('win'):
                os.startfile(folder)
            else:
                subprocess.run(['xdg-open', folder])
        except Exception:
            pass

    # ---------------- small helpers ----------------
    def _load_selection_into_form(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = None
        try:
            vals = self.tree.item(iid, 'values')
        except Exception:
            vals = None
        if vals:
            try:
                self.ent_computer.delete(0, tk.END)
                self.ent_computer.insert(0, vals[0])
                self.ent_alias.delete(0, tk.END)
                self.ent_alias.insert(0, vals[1])
            except Exception:
                pass


def main():
    app = CSInfoGUI()
    app.mainloop()


if __name__ == '__main__':
    main()


