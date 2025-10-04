"""GUI simples para csinfo.

Comportamento principal:
- Ao iniciar coleta, todos os campos de entrada são desabilitados.
- A coleta roda em uma thread de background chamando `csinfo.main(...)`.
- Saída incremental (linhas) enviada por `barra_callback` é exibida progressivamente.
- Ao terminar (ou erro), os campos são reabilitados.
"""

import threading
import queue
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import csinfo
import time
import json
import os
import subprocess
import pathlib
import sys


class CSInfoGUI(tk.Tk):
	def __init__(self):
		super().__init__()
		self.title('CSInfo - Coletor')
		self.geometry('800x600')

		# debug mode controllable via env CSINFO_GUI_DEBUG=1
		self.debug = os.environ.get('CSINFO_GUI_DEBUG', '').lower() in ('1', 'true', 'yes')

		self._create_widgets()
		self._layout_widgets()

		# fila para comunicar da thread worker -> mainloop
		self.queue = queue.Queue()
		self.worker_thread = None
		self.controls = []

		# iniciar loop de consumo de fila
		self.after(100, self._process_queue)

	def _create_widgets(self):
		self.frm_top = ttk.Frame(self)

		# Frame para lista de máquinas e ações
		self.frm_machines = ttk.Frame(self)
		# agora exibimos também o nome da máquina na lista
		self.tree = ttk.Treeview(self.frm_machines, columns=('name','alias','status'), show='headings', selectmode='browse', height=8)
		self.tree.heading('name', text='Máquina')
		self.tree.heading('alias', text='Apelido')
		self.tree.heading('status', text='Estado')
		self.tree.column('name', width=220)
		self.tree.column('alias', width=150)
		self.tree.column('status', width=100, anchor='center')
		# configurar tags para colorir o estado de forma consistente
		try:
			self.tree.tag_configure('online', foreground='green')
			self.tree.tag_configure('offline', foreground='gray')
		except Exception:
			pass
		self.tree_scroll = ttk.Scrollbar(self.frm_machines, orient='vertical', command=self.tree.yview)
		self.tree.configure(yscrollcommand=self.tree_scroll.set)

		# botões da rotina de máquinas
		self.frm_machine_buttons = ttk.Frame(self.frm_machines)
		self.btn_load = ttk.Button(self.frm_machine_buttons, text='Carregar', command=self.load_selected_machine_into_fields)
		self.btn_save_machine = ttk.Button(self.frm_machine_buttons, text='Salvar', command=self.save_selected_or_new_machine)
		self.btn_delete = ttk.Button(self.frm_machine_buttons, text='Excluir', command=self.delete_selected_machine)
		self.btn_open = ttk.Button(self.frm_machine_buttons, text='Abrir pasta', command=self.open_machine_json_folder)
		self.btn_refresh = ttk.Button(self.frm_machine_buttons, text='Refresh', command=self.refresh_machine_statuses)


		self.lbl_computer = ttk.Label(self.frm_top, text='Nome da máquina (opcional):')
		# Nome da máquina deve ser centralizado e em caixa alta
		self.ent_computer = ttk.Entry(self.frm_top, justify='center')
		self.ent_computer.bind('<KeyRelease>', lambda e: self._on_name_keyrelease(e))

		self.lbl_alias = ttk.Label(self.frm_top, text='Apelido para arquivo (opcional):')
		# apelido deve ser caixa alta e centralizado
		self.ent_alias = ttk.Entry(self.frm_top, justify='center')
		self.ent_alias.bind('<KeyRelease>', lambda e: self._on_alias_keyrelease(e))

		# campos de credenciais (adicionados conforme solicitado)
		self.lbl_user = ttk.Label(self.frm_top, text='Usuário:')
		self.ent_user = ttk.Entry(self.frm_top)
		self.lbl_pass = ttk.Label(self.frm_top, text='Senha:')
		self.ent_pass = ttk.Entry(self.frm_top, show='*')

		self.lbl_export = ttk.Label(self.frm_top, text='Formato de exportação:')
		self.cmb_export = ttk.Combobox(self.frm_top, values=('nenhum','txt','pdf','ambos'), state='readonly')
		self.cmb_export.set('nenhum')

		# checkbox de debug removido conforme solicitado

		self.btn_start = ttk.Button(self.frm_top, text='Coletar', command=self.start_collection)
		self.btn_clear = ttk.Button(self.frm_top, text='Limpar', command=self.clear_output)

		self.progress = ttk.Progressbar(self, orient='horizontal', length=400, mode='determinate')
		self.lbl_progress = ttk.Label(self, text='Pronto')

		# área de saída: mantida desabilitada por padrão (escrita via código)
		self.txt_output = ScrolledText(self, wrap='word', state='disabled')

		# lista completa de controles que devem ser bloqueados
		# incluir campos de credenciais para que sejam desabilitados durante processamento
		self.controls = [self.ent_computer, self.ent_alias, self.ent_user, self.ent_pass, self.cmb_export, self.btn_start, self.btn_clear]

		# dados em memória
		self.machine_list = []
		# caminho do arquivo json
		self.machine_json_path = self.get_machine_json_path()
		# carregar lista ao iniciar
		self.load_machine_list()
		# iniciar ping inicial (background)
		self.refresh_machine_statuses()

	def _layout_widgets(self):
		self.frm_top.pack(fill='x', padx=10, pady=8)

		# posicionar frame de máquinas logo abaixo do topo
		self.frm_machines.pack(fill='x', padx=10, pady=(0,8))
		self.tree.pack(side='left', fill='x', expand=True)
		self.tree_scroll.pack(side='left', fill='y')
		self.frm_machine_buttons.pack(side='left', padx=8)
		self.btn_load.pack(fill='x', pady=2)
		self.btn_save_machine.pack(fill='x', pady=2)
		self.btn_delete.pack(fill='x', pady=2)
		self.btn_refresh.pack(fill='x', pady=2)
		self.btn_open.pack(fill='x', pady=2)

		self.lbl_computer.grid(row=0, column=0, sticky='w')
		self.ent_computer.grid(row=0, column=1, sticky='ew', padx=6)
		self.lbl_alias.grid(row=0, column=2, sticky='w', padx=(12,0))
		self.ent_alias.grid(row=0, column=3, sticky='ew', padx=6)

		# campos de credenciais em nova linha
		self.lbl_user.grid(row=1, column=2, sticky='w', padx=(12,0))
		self.ent_user.grid(row=1, column=3, sticky='ew', padx=6)
		self.lbl_pass.grid(row=1, column=0, sticky='w')
		self.ent_pass.grid(row=1, column=1, sticky='ew', padx=6)

		self.lbl_export.grid(row=2, column=0, sticky='w', pady=(8,0))
		self.cmb_export.grid(row=2, column=1, sticky='w', pady=(8,0))
		# chk_debug removido

		self.btn_start.grid(row=2, column=0, pady=10)
		self.btn_clear.grid(row=2, column=1, pady=10)

		# permitir que as colunas 1 e 3 expandam
		self.frm_top.columnconfigure(1, weight=1)
		self.frm_top.columnconfigure(3, weight=1)

		self.progress.pack(fill='x', padx=10, pady=(0,4))
		self.lbl_progress.pack(anchor='w', padx=12)
		self.txt_output.pack(fill='both', expand=True, padx=10, pady=8)

		# bind seleção na tree
		# ao selecionar uma linha, carregar automaticamente os campos
		self.tree.bind('<<TreeviewSelect>>', lambda e: self.load_selected_machine_into_fields())

	def _set_controls_state(self, state='disabled'):
		for w in self.controls:
			try:
				w.configure(state=state)
			except Exception:
				try:
					# alguns widgets usam 'variable' e não aceitam state change; ignorar
					pass
				except Exception:
					pass

	def clear_output(self):
		self.txt_output.configure(state='normal')
		self.txt_output.delete('1.0', tk.END)
		self.txt_output.configure(state='disabled')

	def _on_name_keyrelease(self, event):
		# converte texto para caixa alta mantendo a posição do cursor
		w = event.widget
		try:
			pos = w.index(tk.INSERT)
			text = w.get()
			up = text.upper()
			if text != up:
				w.delete(0, tk.END)
				w.insert(0, up)
				# restaurar cursor: limitar à lenght
				try:
					newpos = min(pos, len(up))
					w.icursor(newpos)
				except Exception:
					pass
		except Exception:
			pass

	def _on_alias_keyrelease(self, event):
		# converte alias para caixa alta mantendo cursor
		w = event.widget
		try:
			pos = w.index(tk.INSERT)
			text = w.get()
			up = text.upper()
			if text != up:
				w.delete(0, tk.END)
				w.insert(0, up)
				try:
					newpos = min(pos, len(up))
					w.icursor(newpos)
				except Exception:
					pass
		except Exception:
			pass

	def get_machine_json_path(self):
		# usar AppData\Roaming\CSInfo\machines_history.json no Windows (histórico conforme solicitado)
		appdata = os.environ.get('APPDATA') or os.path.expanduser('~')
		dirpath = os.path.join(appdata, 'CSInfo')
		os.makedirs(dirpath, exist_ok=True)
		return os.path.join(dirpath, 'machines_history.json')

	def load_machine_list(self):
		try:
			if os.path.exists(self.machine_json_path):
				with open(self.machine_json_path, 'r', encoding='utf-8') as fh:
					self.machine_list = json.load(fh) or []
			else:
				self.machine_list = []
		except Exception:
			self.machine_list = []
		# garantir sort por nome da maquina
		self.machine_list = sorted(self.machine_list, key=lambda x: (x.get('name') or '').lower())
		self.populate_machine_tree()

	def save_machine_list(self):
		try:
			with open(self.machine_json_path + '.tmp', 'w', encoding='utf-8') as fh:
				json.dump(self.machine_list, fh, ensure_ascii=False, indent=2)
			# atomic replace
			os.replace(self.machine_json_path + '.tmp', self.machine_json_path)
		except Exception as e:
			# mostrar no output
			self.txt_output.configure(state='normal')
			self.txt_output.insert(tk.END, f"Falha ao salvar lista de máquinas: {e}\n")
			self.txt_output.configure(state='disabled')

	def populate_machine_tree(self):
		# limpar tree
		for it in self.tree.get_children():
			self.tree.delete(it)
		# adicionar ordenada
		for m in sorted(self.machine_list, key=lambda x: (x.get('name') or '').lower()):
			name = (m.get('name') or '').strip().upper()
			alias = (m.get('alias') or '').strip().upper()
			online = bool(m.get('online'))
			status = 'ONLINE' if online else 'OFFLINE'
			tag = 'online' if online else 'offline'
			# usar iid como nome para seleção/identificação
			self.tree.insert('', 'end', iid=name, values=(name, alias, status), tags=(tag,))

	def load_selected_machine_into_fields(self):
		sel = self.tree.selection()
		if not sel:
			return
		name = sel[0]
		m = next((x for x in self.machine_list if x.get('name') == name), None)
		if not m:
			return
		self.ent_computer.delete(0, tk.END)
		self.ent_computer.insert(0, (m.get('name') or '').upper())
		self.ent_alias.delete(0, tk.END)
		self.ent_alias.insert(0, (m.get('alias') or '').upper())

	def save_selected_or_new_machine(self):
		# garantir maiúsculas
		name = (self.ent_computer.get() or '').strip().upper()
		alias = (self.ent_alias.get() or '').strip().upper()
		if not name:
			self.txt_output.configure(state='normal')
			self.txt_output.insert(tk.END, "Nome da máquina é obrigatório para salvar.\n")
			self.txt_output.configure(state='disabled')
			return
		if not alias:
			self.txt_output.configure(state='normal')
			self.txt_output.insert(tk.END, "Apelido é obrigatório para salvar.\n")
			self.txt_output.configure(state='disabled')
			return
		# atualizar ou inserir
		existing = next((x for x in self.machine_list if x.get('name') == name), None)
		if existing:
			existing['alias'] = alias
		else:
			self.machine_list.append({'name': name, 'alias': alias, 'online': False})
		# persistir e repintar
		self.save_machine_list()
		self.populate_machine_tree()
		# após salvar, executar um ping rápido para atualizar o estado dessa máquina na listagem
		try:
			threading.Thread(target=self._ping_single_and_queue, args=(name,), daemon=True).start()
		except Exception:
			pass
		self.txt_output.configure(state='normal')
		self.txt_output.insert(tk.END, f"Máquina '{name}' salva.\n")
		self.txt_output.configure(state='disabled')

	def delete_selected_machine(self):
		sel = self.tree.selection()
		if not sel:
			return
		name = sel[0]
		self.machine_list = [m for m in self.machine_list if m.get('name') != name]
		self.save_machine_list()
		self.populate_machine_tree()
		self.txt_output.configure(state='normal')
		self.txt_output.insert(tk.END, f"Máquina '{name}' excluída.\n")
		self.txt_output.configure(state='disabled')

	def open_machine_json_folder(self):
		p = os.path.abspath(self.machine_json_path)
		folder = os.path.dirname(p)
		try:
			if sys.platform.startswith('win'):
				os.startfile(folder)
			else:
				subprocess.run(['xdg-open', folder])
		except Exception as e:
			self.txt_output.configure(state='normal')
			self.txt_output.insert(tk.END, f"Falha ao abrir pasta: {e}\n")
			self.txt_output.configure(state='disabled')

	def refresh_machine_statuses(self):
		# executar ping em background para todas as máquinas
		thr = threading.Thread(target=self._ping_worker, daemon=True)
		thr.start()

	def _ping_host(self, host):
		# platform-specific ping
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
		for m in self.machine_list:
			name = (m.get('name') or '').strip()
			# usar o nome original para o ping (sem forçar maiúsculas) — nomes FQDN podem ser sensíveis
			on = self._ping_host(name)
			m['online'] = on
			# enviar status textual consistente
			status_text = 'ONLINE' if on else 'OFFLINE'
			self.queue.put(('machine_status', name, status_text))
		# após todas atualizações, salvar estado
		self.save_machine_list()

	def _ping_single_and_queue(self, name):
		"""Ping uma máquina específica em background e envie resultado para a queue."""
		try:
			on = self._ping_host(name)
			status_text = 'ONLINE' if on else 'OFFLINE'
			self.queue.put(('machine_status', name, status_text))
		except Exception:
			pass

	def start_collection(self):
		# prevenir múltiplos cliques
		if self.worker_thread and self.worker_thread.is_alive():
			return

		computer = self.ent_computer.get().strip() or None
		alias = self.ent_alias.get().strip() or None
		export = self.cmb_export.get()
		include_debug = False

		# limpar área de saída
		self.clear_output()

		# bloquear campos
		self._set_controls_state('disabled')
		self.btn_start.configure(text='⏳ Processando...')
		self.lbl_progress.configure(text='Iniciando...')
		self.progress['value'] = 0

		# callback que será chamado pelo csinfo.main
		def barra_callback(percent_or_none, line_or_stage):
			# se percent_or_none é None -> é uma linha de saída
			try:
				if percent_or_none is None:
					self.queue.put(('line', str(line_or_stage)))
				else:
					# valor percentual
					try:
						perc = int(percent_or_none)
					except Exception:
						perc = 0
					self.queue.put(('progress', perc, str(line_or_stage)))
			except Exception:
				pass

		# worker thread
		def worker():
			try:
				# se credenciais foram preenchidas no formulário, definir como default para o run_powershell
				user = (self.ent_user.get() or '').strip()
				passwd = (self.ent_pass.get() or '')
				if user and passwd:
					try:
						csinfo.set_default_credential(user, passwd)
					except Exception:
						pass
				try:
					csinfo.main(export_type=(export if export != 'nenhum' else None), barra_callback=barra_callback, computer_name=computer, machine_alias=alias)
					self.queue.put(('done', None))
				finally:
					# limpar credencial padrão para não vazar dados
					try:
						csinfo.clear_default_credential()
					except Exception:
						pass
			except Exception as e:
				self.queue.put(('error', str(e)))

		self.worker_thread = threading.Thread(target=worker, daemon=True)
		self.worker_thread.start()

	def _process_queue(self):
		try:
			while True:
				item = self.queue.get_nowait()
				if not item:
					continue
				kind = item[0]
				if kind == 'line':
					line = item[1]
					# adicionar linha de saída (manter histórico completo)
					self.txt_output.configure(state='normal')
					# normalizar CR/LF e garantir que todo o conteúdo é inserido
					try:
						normalized = str(line).replace('\r\n', '\n').replace('\r', '\n')
						self.txt_output.insert(tk.END, normalized + "\n")
						self.txt_output.see(tk.END)
					except Exception:
						try:
							self.txt_output.insert(tk.END, str(line) + "\n")
						except Exception:
							pass
					self.txt_output.configure(state='disabled')
				elif kind == 'machine_status':
					# atualizar a coluna de status da linha com iid igual ao nome
					iid = item[1]
					status = item[2]
					# normalizar para formato textual esperado
					if status in ('🟢', '🔴'):
						status_text = 'ONLINE' if status == '🟢' else 'OFFLINE'
					else:
						status_text = str(status).strip().upper()
					# tentar atualizar diretamente pelo iid (geralmente o nome em MAIÚSCULAS)
					# se modo debug, registrar estado da árvore
					if getattr(self, 'debug', False):
						try:
							self.txt_output.configure(state='normal')
							self.txt_output.insert(tk.END, f"[debug] children iids: {self.tree.get_children()}\n")
							self.txt_output.see(tk.END)
							self.txt_output.configure(state='disabled')
						except Exception:
							pass
					try:
						vals = self.tree.item(iid, 'values')
					except Exception:
						vals = None
					# se não encontrou pelo iid, procurar por linha cujo primeiro valor (name) case-insensitive bata com iid
					if not vals:
						found = None
						for child in self.tree.get_children():
							v = self.tree.item(child, 'values')
							# debug: mostrar o valor avaliado
							if getattr(self, 'debug', False):
								try:
									self.txt_output.configure(state='normal')
									self.txt_output.insert(tk.END, f"[debug] checking child={child} values={v}\n")
									self.txt_output.see(tk.END)
									self.txt_output.configure(state='disabled')
								except Exception:
									pass
							if v and str(v[0]).strip().upper() == str(iid).strip().upper():
								found = child
								vals = v
								break
						if found:
							iid = found
						# se agora temos vals, atualizamos preferencialmente usando tree.set
						if vals:
							try:
								# atualizar apenas a coluna 'status' com o texto normalizado
								self.tree.set(iid, 'status', status_text)
								# aplicar tag correspondente
								tag = 'online' if status_text == 'ONLINE' else 'offline'
								try:
									self.tree.item(iid, tags=(tag,))
								except Exception:
									pass
								# forçar refresh UI
								try:
									self.update_idletasks()
								except Exception:
									pass
							except Exception:
								# fallback: tentar alterar via item
								try:
									self.tree.item(iid, values=(vals[0], vals[1], status_text))
								except Exception:
									pass
					# também atualizar o registro em machine_list para refletir o estado
					try:
						name_key = str((vals[0] if vals else iid)).strip()
						for m in self.machine_list:
							if str(m.get('name') or '').strip().upper() == name_key.upper():
								m['online'] = (status_text == 'ONLINE')
								break
					except Exception:
						pass
					# escrever log reduzido sobre o resultado do ping para ajudar diagnóstico
					try:
						self.txt_output.configure(state='normal')
						self.txt_output.insert(tk.END, f"[ping] {name_key} -> {status_text.lower()}\n")
						self.txt_output.see(tk.END)
						self.txt_output.configure(state='disabled')
					except Exception:
						pass
				elif kind == 'progress':
					perc = item[1]
					etapa = item[2]
					self.progress['value'] = perc
					self.lbl_progress.configure(text=f"{etapa} — {perc}%")
				elif kind == 'done':
					self.lbl_progress.configure(text='Concluído')
					self.progress['value'] = 100
					self.btn_start.configure(text='Coletar')
					self._set_controls_state('normal')
				elif kind == 'error':
					msg = item[1]
					self.txt_output.configure(state='normal')
					self.txt_output.insert(tk.END, f"ERRO: {msg}\n")
					self.txt_output.see(tk.END)
					self.txt_output.configure(state='disabled')
					self.lbl_progress.configure(text='Erro')
					self.btn_start.configure(text='Coletar')
					self._set_controls_state('normal')
		except queue.Empty:
			pass
		# re-schedule
		self.after(100, self._process_queue)


def main():
	app = CSInfoGUI()
	app.mainloop()


if __name__ == '__main__':
	main()

