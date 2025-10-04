import os
import json
import time
import platform
import sys
# Garantir que o diretório do projeto (pai de tests/) esteja em sys.path
proj_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if proj_root not in sys.path:
    sys.path.insert(0, proj_root)

# Import a GUI app class and core
from csinfo_gui import CSInfoApp
import csinfo

# Evitar caixas de diálogo bloqueantes durante o teste
from tkinter import messagebox
messagebox.askyesno = lambda *a, **k: True
messagebox.showwarning = lambda *a, **k: None
messagebox.showinfo = lambda *a, **k: None
messagebox.showerror = lambda *a, **k: None

print('Iniciando teste de histórico...')
app = CSInfoApp()
# Ocultar janela
try:
    app.withdraw()
except Exception:
    pass

# Limpar histórico de teste (usar caminho real definido pelo app)
hist_path = app.history_path
print('History path:', hist_path)
if os.path.exists(hist_path):
    try:
        os.remove(hist_path)
        print('Arquivo de histórico removido para iniciar testes.')
    except Exception as e:
        print('Não foi possível remover arquivo histórico:', e)

# 1) Salvar um par
app.machine_var.set(platform.node())
app.alias_var.set('test_alias_1')
app.save_machine()
print('Salvou alias test_alias_1')
# Ler arquivo
with open(hist_path, 'r', encoding='utf-8') as f:
    data = json.load(f)
print('Conteúdo após salvar:', data)

# 2) Editar o item recém criado (selecionar e editar)
# selecionar primeiro item
app.machines_listbox.select_set(0)
app.edit_selected_machine()
# modificar alias
app.alias_var.set('test_alias_renamed')
app.save_machine()
print('Editou para test_alias_renamed')
with open(hist_path, 'r', encoding='utf-8') as f:
    data = json.load(f)
print('Conteúdo após editar:', data)

# 3) Remover o item
app.machines_listbox.select_set(0)
app.remove_selected_machine()
print('Removido item selecionado')
with open(hist_path, 'r', encoding='utf-8') as f:
    data = json.load(f)
print('Conteúdo após remover:', data)

# 4) Chamar csinfo.main com alias de teste (gera TXT e PDF no diretório atual)
alias_for_report = 'AUTOMATED_TEST'
print('Chamando csinfo.main para gerar TXT + PDF com alias:', alias_for_report)
res = csinfo.main(export_type='ambos', computer_name=None, machine_alias=alias_for_report)
print('Resultado da chamada csinfo.main:', res)

# Verificar se arquivos TXT e PDF foram gerados
if res.get('txt') and os.path.exists(res['txt']):
    print('Arquivo TXT gerado:', res['txt'])
else:
    print('Arquivo TXT não encontrado ou não gerado')

if res.get('pdf') and os.path.exists(res['pdf']):
    print('Arquivo PDF gerado:', res['pdf'])
else:
    print('Arquivo PDF não encontrado ou não gerado')

# Limpeza rápida
try:
    app.destroy()
except Exception:
    pass
print('Teste finalizado.')
