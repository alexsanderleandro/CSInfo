import sys
import os
import json
import traceback
from csinfo._impl import get_processor_info, get_disk_info, run_powershell

def depurar_funcao(nome_funcao, funcao, *args):
    print(f"\n{'='*80}")
    print(f"DEPURANDO: {nome_funcao}")
    print(f"{'='*80}")
    
    try:
        print(f"Chamando {nome_funcao}...")
        resultado = funcao(*args)
        print(f"✅ Função executada com sucesso!")
        print(f"Tipo de retorno: {type(resultado)}")
        
        if resultado is not None:
            if isinstance(resultado, (list, tuple)):
                print(f"Número de itens retornados: {len(resultado)}")
                if len(resultado) > 0:
                    print("Primeiro item:")
                    print(json.dumps(resultado[0], indent=2, default=str))
            else:
                print("Resultado:")
                print(json.dumps(resultado, indent=2, default=str))
        else:
            print("A função retornou None")
            
        return True, resultado
        
    except Exception as e:
        print(f"❌ Erro ao executar {nome_funcao}:")
        traceback.print_exc()
        return False, None

def testar_powershell(computador):
    print("\n" + "="*80)
    print("TESTANDO POWERSHELL REMOTO")
    print("="*80)
    
    # Comando simples para testar o PowerShell remoto
    cmd = "Get-Process | Select-Object -First 3 | ConvertTo-Json"
    print(f"Executando comando remoto: {cmd}")
    
    try:
        resultado = run_powershell(cmd, computer_name=computador)
        print("✅ Comando PowerShell executado com sucesso!")
        print("Saída:")
        print(resultado)
        return True
    except Exception as e:
        print("❌ Falha ao executar comando PowerShell:")
        print(str(e))
        return False

def main(computador):
    print(f"\n{'='*80}")
    print(f"DEPURAÇÃO DAS FUNÇÕES PARA {computador}")
    print(f"{'='*80}")
    
    # Testa o PowerShell remoto primeiro
    ps_ok = testar_powershell(computador)
    
    if not ps_ok:
        print("\n⚠️  O PowerShell remoto não está funcionando corretamente.")
        print("Isso pode afetar o funcionamento das funções do CSInfo.")
    
    # Depura get_processor_info
    sucesso_cpu, resultado_cpu = depurar_funcao(
        "get_processor_info", 
        get_processor_info, 
        computador if computador != "localhost" else None
    )
    
    # Depura get_disk_info
    sucesso_disco, resultado_disco = depurar_funcao(
        "get_disk_info", 
        get_disk_info, 
        computador if computador != "localhost" else None
    )
    
    # Se ambas falharem, tenta com um comando PowerShell direto
    if not sucesso_cpu or not sucesso_disco:
        print("\n" + "="*80)
        print("TENTANDO COMANDO POWERSHELL DIRETO")
        print("="*80)
        
        # Comando para obter informações do processador
        cmd_cpu = """
        $cpu = Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, L2CacheSize, L3CacheSize
        $cpu | ConvertTo-Json
        """
        
        print("\nExecutando consulta direta ao processador...")
        try:
            resultado = run_powershell(cmd_cpu, computer_name=computador)
            print("✅ Consulta ao processador bem-sucedida!")
            print("Resultado:")
            print(resultado)
        except Exception as e:
            print("❌ Falha na consulta direta ao processador:")
            print(str(e))
        
        # Comando para obter informações do disco
        cmd_disco = """
        $discos = Get-WmiObject -Class Win32_DiskDrive | Select-Object Model, Size, Partitions, InterfaceType, MediaType, SerialNumber
        $discos | ConvertTo-Json
        """
        
        print("\nExecutando consulta direta aos discos...")
        try:
            resultado = run_powershell(cmd_disco, computer_name=computador)
            print("✅ Consulta aos discos bem-sucedida!")
            print("Resultado:")
            print(resultado)
        except Exception as e:
            print("❌ Falha na consulta direta aos discos:")
            print(str(e))
    
    print("\n" + "="*80)
    print("DEPURAÇÃO CONCLUÍDA")
    print("="*80)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador (ou 'localhost' para o computador local): ").strip()
    
    main(computador)
