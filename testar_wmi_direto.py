import wmi
import sys
import json

def testar_wmi_direto(computador):
    print(f"\n{'='*80}")
    print(f"TESTANDO WMI DIRETO EM {computador}")
    print(f"{'='*80}")
    
    try:
        # Conecta ao WMI do computador remoto
        print(f"Conectando a {computador} via WMI...")
        conexao = wmi.WMI(computer=computador)
        print("✅ Conexão WMI bem-sucedida!")
        
        # Testa consulta ao processador
        print("\nTestando consulta ao processador (Win32_Processor)...")
        try:
            cpus = conexao.Win32_Processor()
            print(f"✅ Encontrado(s) {len(cpus)} processador(es):")
            for i, cpu in enumerate(cpus, 1):
                print(f"  CPU {i}: {cpu.Name.strip() if cpu.Name else 'Nome não disponível'}")
                print(f"     Arquitetura: {cpu.Architecture}")
                print(f"     Núcleos: {cpu.NumberOfCores}")
                print(f"     Threads: {cpu.NumberOfLogicalProcessors}")
                print(f"     Clock: {cpu.MaxClockSpeed} MHz")
        except Exception as e:
            print(f"❌ Erro ao consultar Win32_Processor: {str(e)}")
        
        # Testa consulta aos discos
        print("\nTestando consulta aos discos (Win32_DiskDrive)...")
        try:
            discos = conexao.Win32_DiskDrive()
            print(f"✅ Encontrado(s) {len(discos)} disco(s):")
            for i, disco in enumerate(discos, 1):
                tamanho_gb = round(int(disco.Size or 0) / (1024**3), 2)
                print(f"  Disco {i}: {disco.Caption or 'Sem descrição'}")
                print(f"     Modelo: {disco.Model or 'N/A'}")
                print(f"     Tamanho: {tamanho_gb} GB")
                print(f"     Interface: {disco.InterfaceType or 'N/A'}")
                print(f"     Partições: {disco.Partitions or 'N/A'}")
        except Exception as e:
            print(f"❌ Erro ao consultar Win32_DiskDrive: {str(e)}")
        
        # Testa consulta a partições lógicas
        print("\nTestando consulta a partições lógicas (Win32_LogicalDisk)...")
        try:
            discos_logicos = conexao.Win32_LogicalDisk()
            print(f"✅ Encontrada(s) {len(discos_logicos)} partição(ões) lógica(s):")
            for i, disco in enumerate(discos_logicos, 1):
                tamanho_gb = round(int(disco.Size or 0) / (1024**3), 2)
                livre_gb = round(int(disco.FreeSpace or 0) / (1024**3), 2)
                print(f"  {disco.DeviceID}: {disco.VolumeName or 'Sem rótulo'}")
                print(f"     Sistema de arquivos: {disco.FileSystem or 'N/A'}")
                print(f"     Tamanho total: {tamanho_gb} GB")
                print(f"     Espaço livre: {livre_gb} GB")
        except Exception as e:
            print(f"❌ Erro ao consultar Win32_LogicalDisk: {str(e)}")
        
        return True
        
    except Exception as e:
        print(f"❌ Falha na conexão WMI: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main(computador):
    print(f"\n{'='*80}")
    print(f"TESTE DIRETO DE WMI EM {computador}")
    print(f"{'='*80}")
    
    # Testa WMI direto
    wmi_ok = testar_wmi_direto(computador)
    
    if not wmi_ok:
        print("\n⚠️  A conexão WMI direta falhou. Verifique:")
        print("1. Se o serviço 'Windows Management Instrumentation' está em execução no computador remoto")
        print("2. Se a conta tem permissões administrativas no computador remoto")
        print("3. Se as regras de firewall permitem o tráfego WMI (porta 135 e portas dinâmicas RPC)")
    
    print("\n" + "="*80)
    print("TESTE CONCLUÍDO")
    print("="*80)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador: ").strip()
    
    main(computador)
