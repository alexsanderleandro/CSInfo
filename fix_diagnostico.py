import sys
import os

# Adiciona o diretório atual ao path para importar o módulo csinfo
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def fix_gui_file():
    file_path = os.path.join('csinfo_gui.py')
    
    # Lê o conteúdo atual do arquivo
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Encontra a linha onde está o código que precisa ser modificado
    start_line = -1
    for i, line in enumerate(lines):
        if '# Verifica problemas de permissão' in line:
            start_line = i
            break
    
    if start_line != -1:
        # Remove as linhas antigas
        end_line = start_line + 5  # Remove as próximas 5 linhas
        
        # Novo conteúdo
        new_lines = [
            '                # Verifica problemas de permissão\n',
            '                admin_test = results.get(\'tests\', {}).get(\'admin_check\', {})\n',
            '                if admin_test and not admin_test.get(\'success\', True):\n',
            '                    self.info_text.insert(tk.END, "- Não foi possível verificar os privilégios administrativos.\\n")\n',
            '                    if \'error\' in admin_test:\n',
            '                        self.info_text.insert(tk.END, rf"  Erro: {admin_test.get(\'error\', \'Erro desconhecido\')}\\n")\n',
            '                elif admin_test and not admin_test.get(\'is_admin\', False):\n',
            '                    self.info_text.insert(tk.END, "- O usuário atual não tem privilégios administrativos no computador remoto.\\n")\n',
            '                    self.info_text.insert(tk.END, rf"  Usuário atual: {admin_test.get(\'current_user\', \'Desconhecido\')}\\n")\n',
            '                else:\n',
            '                    self.info_text.insert(tk.END, "- Verificação de privilégios administrativos não disponível.\\n")\n'
        ]
        
        # Substitui as linhas antigas pelas novas
        
        # Salva o arquivo modificado
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(lines)
        
        print("Arquivo csinfo_gui.py atualizado com sucesso!")
    else:
        print("Não foi possível encontrar o local para a correção no arquivo.")

if __name__ == "__main__":
    fix_gui_file()
