import sys
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from csinfo._impl import test_remote_connection

def run_diagnostic(computer_name):
    """Executa o diagnóstico e exibe os resultados em uma janela simples"""
    try:
        # Executa o teste de conexão remota
        results = test_remote_connection(computer_name)
        
        # Cria uma nova janela para exibir os resultados
        window = tk.Tk()
        window.title(f"Resultados do Diagnóstico - {computer_name}")
        window.geometry("800x600")
        
        # Área de texto para exibir os resultados
        text = tk.Text(window, wrap=tk.WORD, padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True)
        
        # Adiciona os resultados ao texto
        text.insert(tk.END, f"=== DIAGNÓSTICO PARA {computer_name.upper()} ===\n\n")
        
        # Exibe o resumo
        summary = results.get('summary', {})
        text.insert(tk.END, f"RESUMO: {summary.get('passed_tests', 0)}/{summary.get('total_tests', 0)} testes aprovados ")
        text.insert(tk.END, f"({summary.get('success_rate', '0%')})\n\n")
        
        # Exibe detalhes de cada teste
        for test_name, test_result in results.get('tests', {}).items():
            status = "✅" if test_result.get('success', False) else "❌"
            text.insert(tk.END, f"{status} {test_name.upper().replace('_', ' ')}: ")
            
            if test_result.get('success', False):
                details = test_result.get('details', {})
                if isinstance(details, dict):
                    if 'disk_count' in details:
                        text.insert(tk.END, f"Encontrados {details['disk_count']} discos\n")
                    else:
                        text.insert(tk.END, "Sucesso\n")
                        for k, v in details.items():
                            # Verifica se o valor é um dicionário ou lista para formatar melhor
                            if isinstance(v, (dict, list)):
                                text.insert(tk.END, f"  - {k}: {json.dumps(v, indent=2, ensure_ascii=False)}\n")
                            else:
                                text.insert(tk.END, f"  - {k}: {v}\n")
                else:
                    text.insert(tk.END, f"{details}\n")
            else:
                text.insert(tk.END, f"Falha: {test_result.get('error', 'Erro desconhecido')}\n")
            
            text.insert(tk.END, "\n")
        
        # Botão para fechar
        btn_close = tk.Button(window, text="Fechar", command=window.destroy)
        btn_close.pack(pady=10)
        
        window.mainloop()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao executar o diagnóstico:\n{str(e)}")

if __name__ == "__main__":
    import json
    import sys
    
    # Verifica se foi fornecido um nome de computador como argumento
    if len(sys.argv) > 1:
        computer_name = sys.argv[1]
    else:
        # Se não foi fornecido, pede ao usuário
        root = tk.Tk()
        root.withdraw()  # Esconde a janela principal
        computer_name = simpledialog.askstring("Diagnóstico", "Digite o nome do computador:")
        root.destroy()
        
        if not computer_name:
            sys.exit("Nenhum computador especificado.")
    
    run_diagnostic(computer_name)
