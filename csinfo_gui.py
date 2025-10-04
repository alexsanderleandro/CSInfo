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
        self.title('CSInfo GUI')
        self.geometry('980x640')

        # state
        self.queue = queue.Queue()
        self.worker_thread = None
        self._processing = False
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
        self.btn_save = ttk.Button(btn_fr, text='Salvar', command=self.save_selected_or_new_machine)
        self.btn_save.pack(side='left', fill='x', expand=True)
        self.btn_delete = ttk.Button(btn_fr, text='Excluir', command=self.delete_selected_machine)
        self.btn_delete.pack(side='left', fill='x', expand=True, padx=(6, 0))

        # separador fino antes das credenciais
        sep = ttk.Separator(left, orient='horizontal')
        sep.pack(fill='x', pady=(8, 8))

        # Campos de credenciais movidos para baixo dos botões (rótulos completos)
        # rótulo de seção
        ttk.Label(left, text='Administrador da rede', font=('Segoe UI', 9, 'bold'), foreground='#333').pack(anchor='w', pady=(0, 6))
        ttk.Label(left, text='Usuário administrador da rede:').pack(anchor='w')
        self.ent_user = ttk.Entry(left, width=32)
        self.ent_user.pack()
        ttk.Label(left, text='Senha administrador da rede:').pack(anchor='w', pady=(8, 0))
        self.ent_pass = ttk.Entry(left, width=32, show='*')
        self.ent_pass.pack()

        ttk.Label(left, text='Exportar como:').pack(anchor='w', pady=(8, 0))
        # valores internos usados pelo código: 'nenhum', 'txt', 'pdf', 'ambos'
        self.cmb_export = ttk.Combobox(left, values=('nenhum', 'txt', 'pdf', 'ambos'), state='readonly')
        self.cmb_export.current(0)
        self.cmb_export.pack()
        self.cmb_export.bind('<<ComboboxSelected>>', lambda e: self._update_export_button_state())

        self.btn_export = ttk.Button(left, text='Exportar', command=self._do_export, state='disabled')
        self.btn_export.pack(fill='x', pady=(8, 0))

        self.btn_start = ttk.Button(left, text='Coletar', command=self.start_collection)
        self.btn_start.pack(fill='x', pady=(12, 0))

        ttk.Button(left, text='Abrir pasta de máquinas', command=self.open_machine_json_folder).pack(fill='x', pady=(8, 0))

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
        try:
            self.tree.tag_configure('online', background='#e6ffed')
            self.tree.tag_configure('offline', background='#ffe6e6')
        except Exception:
            pass
        self.tree.bind('<Double-1>', lambda e: self._load_selection_into_form())

        right = ttk.Frame(frm)
        right.pack(side='right', fill='both', expand=True)
        self.txt_output = ScrolledText(right, height=20, state='disabled', wrap='word', font=('Consolas', 10))
        self.txt_output.pack(fill='both', expand=True)

        bar_fr = ttk.Frame(right)
        bar_fr.pack(fill='x')
        self.progress = ttk.Progressbar(bar_fr, orient='horizontal', mode='determinate')
        self.progress.pack(fill='x', side='left', expand=True, padx=(0, 8))
        self.lbl_progress = ttk.Label(bar_fr, text='Pronto')
        self.lbl_progress.pack(side='right')

        rodape = tk.Label(self, text='CSInfo GUI', font=('Segoe UI', 8), fg='#666')
        rodape.pack(side='bottom', pady=(0, 6), fill='x')

    # persistence
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
            try:
                self.tree.insert('', 'end', values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'), tags=(tag,))
            except Exception:
                try:
                    self.tree.insert('', 'end', values=(name.upper(), alias.upper(), 'ONLINE' if online else 'OFFLINE'))
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
            threading.Thread(target=self._ping_single_and_queue, args=(name,), daemon=True).start()
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
                proc = subprocess.run(['ping', '-n', '1', '-w', '1000', host], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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

    def _ping_worker(self):
        for m in list(self.machine_list):
            name = (m.get('name') or '').strip()
            on = self._ping_host(name)
            m['online'] = on
            status_text = 'ONLINE' if on else 'OFFLINE'
            self.queue.put(('machine_status', name, status_text))
        self.save_machine_list()

    # collection
    def start_collection(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        computer = (self.ent_computer.get() or '').strip() or None
        alias = (self.ent_alias.get() or '').strip() or None

        self.clear_output()
        self._set_controls_state('disabled')
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
                        csinfo.main(export_type=(self.cmb_export.get() if self.cmb_export.get() != 'nenhum' else None),
                                    barra_callback=barra_callback,
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
                    try:
                        self._append_output(f"[ping] {key} -> {status_text}")
                    except Exception:
                        pass
                elif kind == 'done':
                    self._processing = False
                    self._set_controls_state('normal')
                    self.btn_start.configure(text='Coletar')
                    self.lbl_progress.configure(text='Pronto')
                    self.progress['value'] = 100
                    if self.last_lines:
                        self.btn_export.configure(state='normal')
                elif kind == 'error':
                    self._processing = False
                    self._set_controls_state('normal')
                    self.btn_start.configure(text='Coletar')
                    self.lbl_progress.configure(text='Erro')
                    messagebox.showerror('Erro', str(item[1]))
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            if not self._processing:
                self.progress['value'] = 0
            self.after(100, self._process_queue)

    def _set_controls_state(self, state='normal'):
        widgets = [self.ent_computer, self.ent_alias, self.ent_user, self.ent_pass, self.cmb_export, self.btn_save, self.btn_delete, self.btn_export, self.btn_start]
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
        if not self.last_lines:
            messagebox.showinfo('Exportar', 'Nenhum dado para exportar')
            return
        fmt = self.cmb_export.get()
        alias = (self.ent_alias.get() or '').strip() or None
        comp = (self._last_collection_computer or (self.ent_computer.get() or '').strip() or '')

        def _safe(s):
            try:
                if csinfo and hasattr(csinfo, 'safe_filename'):
                    return csinfo.safe_filename(s)
            except Exception:
                pass
            return re.sub(r'[^0-9A-Za-z_.-]+', '_', str(s or ''))

        base = f"Info_maquina_{_safe(alias) + '_' if alias else ''}{_safe(comp)}"
        folder = os.getcwd()
        try:
            if fmt in ('txt', 'ambos'):
                p = os.path.join(folder, base + '.txt')
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
                p = os.path.join(folder, base + '.pdf')
                try:
                    if csinfo and hasattr(csinfo, 'write_pdf_report'):
                        try:
                            csinfo.write_pdf_report(p, self.last_lines, comp)
                        except TypeError:
                            csinfo.write_pdf_report(p, self.last_lines)
                    else:
                        with open(p, 'w', encoding='utf-8') as fh:
                            fh.write('\n'.join(self.last_lines))
                    self._append_output(f'Exportado PDF: {p}')
                except Exception as e:
                    messagebox.showerror('Exportar', f'Erro ao escrever PDF: {e}')
                    return
            try:
                if sys.platform.startswith('win'):
                    os.startfile(folder)
                else:
                    subprocess.run(['xdg-open', folder])
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
                self.ent_computer.delete(0, tk.END)
                self.ent_computer.insert(0, vals[0])
                self.ent_alias.delete(0, tk.END)
                self.ent_alias.insert(0, vals[1] if len(vals) > 1 else '')
            except Exception:
                pass


def main():
    app = CSInfoGUI()
    app.mainloop()


if __name__ == '__main__':
    main()


