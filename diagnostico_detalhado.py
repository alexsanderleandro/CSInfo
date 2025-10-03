import sys
import os
import platform
import socket
import subprocess
import json
import importlib.util
from datetime import datetime

# Verifica e instala dependências necessárias
def instalar_dependencias():
    dependencias = ['wmi', 'psutil', 'pywin32']
    
    for pacote in dependencias:
        if importlib.util.find_spec(pacote) is None:
            print(f"Instalando {pacote}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
                print(f"✅ {pacote} instalado com sucesso!")
            except subprocess.CalledProcessError as e:
                print(f"❌ Falha ao instalar {pacote}: {str(e)}")
                return False
    return True

# Tenta importar as dependências
try:
    import wmi
    import psutil
except ImportError:
    print("Dependências não encontradas. Instalando...")
    if not instalar_dependencias():
        print("\n⚠️  Não foi possível instalar todas as dependências automaticamente.")
        print("Por favor, instale manualmente com o comando:")
        print("pip install wmi psutil pywin32")
        sys.exit(1)
    
    # Tenta importar novamente após a instalação
    try:
        import wmi
        import psutil
    except ImportError:
        print("\n❌ Não foi possível importar as dependências necessárias.")
        print("Por favor, reinicie o script após instalar as dependências manualmente.")
        sys.exit(1)

def verificar_conexao(hostname):
    """Verifica a conectividade básica com o host"""
    try:
        # Tenta resolver o nome do host para IP
        ip = socket.gethostbyname(hostname)
        
        # Testa o ping
        param = '-n' if platform.system().lower() == 'windows' else '-c'
        command = ['ping', param, '2', hostname]
        ping_result = subprocess.run(command, capture_output=True, text=True)
        
        return {
            'sucesso': ping_result.returncode == 0,
            'ip': ip,
            'saida': ping_result.stdout,
            'erro': ping_result.stderr if ping_result.returncode != 0 else None
        }
    except Exception as e:
        return {
            'sucesso': False,
            'erro': str(e)
        }

def verificar_servicos_remotos(hostname):
    """Verifica se os serviços necessários para coleta remota estão em execução"""
    servicos_necessarios = [
        'RemoteRegistry',
        'WinRM',
        'WMI',
        'DCOM'
    ]
    
    resultados = {}
    
    try:
        # Conecta ao computador remoto
        conexao = wmi.WMI(hostname)
        
        for servico in servicos_necessarios:
            try:
                status = conexao.Win32_Service(Name=servico)
                if status:
                    resultados[servico] = {
                        'status': status[0].State,
                        'iniciado': status[0].Started,
                        'tipo_inicializacao': status[0].StartMode
                    }
                else:
                    resultados[servico] = {
                        'erro': 'Serviço não encontrado'
                    }
            except Exception as e:
                resultados[servico] = {
                    'erro': str(e)
                }
    except Exception as e:
        return {
            'erro': f'Erro ao conectar ao WMI remoto: {str(e)}'
        }
    
    return resultados

def verificar_permissoes_wmi(hostname):
    """Verifica as permissões de WMI no computador remoto"""
    try:
        conexao = wmi.WMI(hostname)
        # Tenta acessar algumas classes comuns
        sistema = conexao.Win32_ComputerSystem()[0]
        return {
            'sucesso': True,
            'nome_computador': sistema.Name,
            'fabricante': sistema.Manufacturer,
            'modelo': sistema.Model
        }
    except Exception as e:
        return {
            'sucesso': False,
            'erro': str(e)
        }

def verificar_arquivos_necessarios():
    """Verifica se os arquivos necessários para a coleta existem"""
    arquivos_necessarios = [
        'csinfo/__init__.py',
        'csinfo/_impl.py',
        'csinfo_gui.py'
    ]
    
    resultados = {}
    
    for arquivo in arquivos_necessarios:
        caminho = os.path.join(os.path.dirname(__file__), arquivo)
        resultados[arquivo] = {
            'existe': os.path.exists(caminho),
            'caminho': caminho
        }
    
    return resultados

def executar_diagnostico(hostname):
    """Executa todos os testes de diagnóstico"""
    resultados = {
        'computador': hostname,
        'data_hora': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'sistema_operacional': platform.platform(),
        'python_version': platform.python_version(),
        'testes': {}
    }
    
    print(f"\n=== INICIANDO DIAGNÓSTICO PARA {hostname.upper()} ===\n")
    
    # Teste 1: Verificar arquivos necessários
    print("1. Verificando arquivos necessários...")
    resultados['testes']['arquivos_necessarios'] = verificar_arquivos_necessarios()
    
    # Teste 2: Verificar conectividade básica
    print("2. Testando conectividade básica...")
    resultados['testes']['conectividade'] = verificar_conexao(hostname)
    
    # Teste 3: Verificar permissões WMI
    print("3. Verificando permissões WMI...")
    resultados['testes']['permissoes_wmi'] = verificar_permissoes_wmi(hostname)
    
    # Teste 4: Verificar serviços remotos
    print("4. Verificando serviços remotos...")
    resultados['testes']['servicos_remotos'] = verificar_servicos_remotos(hostname)
    
    # Salva os resultados em um arquivo JSON
    nome_arquivo = f"diagnostico_{hostname}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        json.dump(resultados, f, indent=2, ensure_ascii=False)
    
    print(f"\n=== DIAGNÓSTICO CONCLUÍDO ===")
    print(f"Resultados salvos em: {os.path.abspath(nome_arquivo)}")
    
    return resultados

if __name__ == "__main__":
    if len(sys.argv) > 1:
        hostname = sys.argv[1]
    else:
        hostname = input("Digite o nome do computador para diagnóstico: ")
    
    try:
        resultados = executar_diagnostico(hostname)
        
        # Exibe um resumo dos resultados
        print("\n=== RESUMO DOS RESULTADOS ===")
        print(f"Conectividade: {'✅' if resultados['testes']['conectividade']['sucesso'] else '❌'}")
        
        if 'permissoes_wmi' in resultados['testes']:
            print(f"Permissões WMI: {'✅' if resultados['testes']['permissoes_wmi'].get('sucesso', False) else '❌'}")
        
        print(f"\nArquivos necessários:")
        for arquivo, status in resultados['testes']['arquivos_necessarios'].items():
            print(f"- {arquivo}: {'✅' if status['existe'] else '❌'} {status['caminho']}")
        
    except Exception as e:
        print(f"\n❌ ERRO durante o diagnóstico: {str(e)}")
