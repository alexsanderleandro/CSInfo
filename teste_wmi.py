import wmi
import sys

def testar_wmi(computador):
    print(f"\n=== TESTE DE CONEXÃO WMI PARA {computador.upper()} ===\n")
    
    try:
        print("1. Tentando conectar ao WMI...")
        conexao = wmi.WMI(computer=computador)
        print("✅ Conexão WMI bem-sucedida!")
        
        print("\n2. Obtendo informações do sistema...")
        sistema = conexao.Win32_ComputerSystem()[0]
        print(f"   Computador: {sistema.Name}")
        print(f"   Fabricante: {sistema.Manufacturer}")
        print(f"   Modelo: {sistema.Model}")
        
        print("\n3. Verificando sistema operacional...")
        os_info = conexao.Win32_OperatingSystem()[0]
        print(f"   SO: {os_info.Caption}")
        print(f"   Versão: {os_info.Version}")
        
        print("\n✅ Teste WMI concluído com sucesso!")
        return True
        
    except Exception as e:
        print(f"\n❌ Erro ao acessar WMI: {str(e)}")
        
        if "Access is denied" in str(e):
            print("\n⚠️  Acesso negado. Verifique se você tem privilégios administrativos no computador remoto.")
        elif "The RPC server is unavailable" in str(e):
            print("\n⚠️  O servidor RPC não está disponível. Verifique se o serviço RemoteRegistry está em execução.")
        elif "The service cannot be started" in str(e):
            print("\n⚠️  O serviço WMI não pode ser iniciado. Pode haver um problema com a instalação do Windows.")
        
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador: ")
    
    testar_wmi(computador)
