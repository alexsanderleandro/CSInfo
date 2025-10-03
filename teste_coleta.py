import wmi
import sys
from datetime import datetime

def testar_coleta(computador):
    print(f"\n=== TESTE DE COLETA DE DADOS PARA {computador.upper()} ===\n")
    
    try:
        # Conecta ao WMI
        print("1. Conectando ao WMI...")
        conexao = wmi.WMI(computer=computador)
        
        # 1. Informações do Sistema
        print("\n2. Coletando informações do sistema...")
        sistema = conexao.Win32_ComputerSystem()[0]
        print(f"   - Computador: {sistema.Name}")
        print(f"   - Fabricante: {sistema.Manufacturer}")
        print(f"   - Modelo: {sistema.Model}")
        
        # 2. Sistema Operacional
        print("\n3. Coletando informações do SO...")
        os_info = conexao.Win32_OperatingSystem()[0]
        print(f"   - SO: {os_info.Caption}")
        print(f"   - Versão: {os_info.Version}")
        print(f"   - Arquitetura: {os_info.OSArchitecture}")
        
        # 3. Processador
        print("\n4. Coletando informações do processador...")
        processador = conexao.Win32_Processor()[0]
        print(f"   - Processador: {processador.Name.strip()}")
        print(f"   - Núcleos: {processador.NumberOfCores}")
        print(f"   - Threads: {processador.NumberOfLogicalProcessors}")
        
        # 4. Memória
        print("\n5. Coletando informações de memória...")
        memoria = conexao.Win32_ComputerSystem()[0]
        total_mem = int(int(memoria.TotalPhysicalMemory) / (1024**3))  # Convertendo para GB
        print(f"   - Memória Total: {total_mem} GB")
        
        # 5. Discos
        print("\n6. Coletando informações de discos...")
        discos = conexao.Win32_DiskDrive()
        for i, disco in enumerate(discos, 1):
            tamanho_gb = int(int(disco.Size or 0) / (1024**3))
            print(f"   - Disco {i}: {disco.Caption}")
            print(f"     Tamanho: {tamanho_gb} GB")
            print(f"     Interface: {disco.InterfaceType or 'Desconhecida'}")
        
        # 6. Rede
        print("\n7. Coletando informações de rede...")
        adaptadores = conexao.Win32_NetworkAdapterConfiguration(IPEnabled=True)
        for i, adaptador in enumerate(adaptadores, 1):
            print(f"   - Adaptador {i}: {adaptador.Description}")
            if adaptador.IPAddress:
                print(f"     IP: {', '.join(adaptador.IPAddress)}")
            if adaptador.MACAddress:
                print(f"     MAC: {adaptador.MACAddress}")
        
        print("\n✅ Coleta de dados concluída com sucesso!")
        
    except Exception as e:
        print(f"\n❌ Erro durante a coleta: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    return True

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador: ")
    
    testar_coleta(computador)
