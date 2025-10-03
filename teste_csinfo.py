import os
import sys
import importlib.util

# Adiciona o diretório atual ao path para importar o módulo csinfo
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Importa funções específicas do CSInfo para teste
from csinfo._impl import (
    get_network_details,
    get_firewall_status,
    get_windows_update_status,
    get_running_processes,
    get_critical_services,
    get_firewall_controller,
    get_machine_name,
    get_os_version,
    get_memory_info,
    get_disk_info,
    get_processor_info
)

def testar_funcao(nome, funcao, *args):
    print(f"\n=== TESTANDO: {nome} ===")
    try:
        resultado = funcao(*args)
        print(f"✅ Sucesso! Retornou: {type(resultado)}")
        if resultado is not None:
            if isinstance(resultado, dict):
                print(f"   Primeiros itens: {list(resultado.items())[:2]}")
            elif isinstance(resultado, (list, tuple)):
                print(f"   Primeiros itens: {resultado[:2]}")
        return True
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main(computador):
    print(f"\n{'='*50}")
    print(f"TESTANDO CONEXÃO COM {computador}")
    print(f"{'='*50}\n")
    
    # Testa funções básicas
    funcoes = [
        ("get_machine_name", get_machine_name, computador),
        ("get_os_version", get_os_version, computador),
        ("get_network_details", get_network_details, computador),
        ("get_firewall_status", get_firewall_status, computador),
        ("get_windows_update_status", get_windows_update_status, computador),
        ("get_running_processes", get_running_processes, computador),
        ("get_critical_services", get_critical_services, computador),
        ("get_firewall_controller", get_firewall_controller, computador),
        ("get_memory_info", get_memory_info, computador),
        ("get_processor_info", get_processor_info, computador),
        ("get_disk_info", get_disk_info, computador)
    ]
    
    sucessos = 0
    for nome, funcao, *args in funcoes:
        if testar_funcao(nome, funcao, *args):
            sucessos += 1
    
    print(f"\n{'='*50}")
    print(f"RESUMO: {sucessos}/{len(funcoes)} testes passaram")
    print(f"{'='*50}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador (ou deixe em branco para local): ").strip() or None
    
    main(computador)
