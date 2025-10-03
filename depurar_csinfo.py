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
        print(f"Chamando {nome_funcao} com argumentos: {args if args else 'Nenhum'}")
        
        # Habilita o log detalhado do que está acontecendo
        print("\n=== LOG DE EXECUÇÃO ===")
        resultado = funcao(*args)
        
        print("\n=== RESULTADO ===")
        print(f"✅ Função executada com sucesso!")
        print(f"Tipo de retorno: {type(resultado)}")
        
        if resultado is not None:
            if isinstance(resultado, (list, tuple)):
                print(f"Número de itens retornados: {len(resultado)}")
                if len(resultado) > 0:
                    print("Primeiro item:")
                    print(json.dumps(resultado[0], indent=2, default=str))
            elif isinstance(resultado, dict):
                print("Chaves retornadas:")
                print(json.dumps(list(resultado.keys()), indent=2))
                print("Primeiros itens:")
                items = list(resultado.items())[:3]  # Mostra apenas os 3 primeiros itens
                for k, v in items:
                    print(f"  {k}: {v}")
            else:
                print(f"Valor retornado: {resultado}")
        else:
            print("A função retornou None")
            
        return True, resultado
        
    except Exception as e:
        print(f"\n❌ Erro ao executar {nome_funcao}:")
        print(f"Tipo de erro: {type(e).__name__}")
        print(f"Mensagem: {str(e)}")
        print("\nStack trace:")
        traceback.print_exc()
        return False, None

def testar_comandos_diretos(computador):
    print("\n" + "="*80)
    print("TESTANDO COMANDOS POWERSHELL DIRETOS")
    print("="*80)
    
    # Comando para obter informações do processador
    cmd_cpu = """
    try {
        $cpu = Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, L2CacheSize, L3CacheSize
        $cpu | ConvertTo-Json -Depth 5
    } catch {
        Write-Output "Erro ao obter informações do processador: $_"
    }
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
    try {
        $discos = Get-WmiObject -Class Win32_DiskDrive | Select-Object Model, Size, Partitions, InterfaceType, MediaType, SerialNumber
        $discos | ConvertTo-Json -Depth 5
    } catch {
        Write-Output "Erro ao obter informações dos discos: $_"
    }
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

def main(computador):
    print(f"\n{'='*80}")
    print(f"DEPURAÇÃO DETALHADA DAS FUNÇÕES DO CSINFO EM {computador}")
    print(f"{'='*80}")
    
    # Verifica se o parâmetro é 'localhost' ou vazio
    if not computador or computador.lower() == 'localhost':
        computador = None
    
    # Depura get_processor_info
    print("\n" + "="*80)
    print(f"INICIANDO DEPURAÇÃO DE get_processor_info")
    sucesso_cpu, resultado_cpu = depurar_funcao(
        "get_processor_info", 
        get_processor_info, 
        computador
    )
    
    # Depura get_disk_info
    print("\n" + "="*80)
    print(f"INICIANDO DEPURAÇÃO DE get_disk_info")
    sucesso_disco, resultado_disco = depurar_funcao(
        "get_disk_info", 
        get_disk_info, 
        computador
    )
    
    # Se alguma das funções falhou, tenta comandos diretos
    if not sucesso_cpu or not sucesso_disco:
        testar_comandos_diretos(computador if computador else 'localhost')
    
    print("\n" + "-"*80)
    print("ANÁLISE DE RESULTADOS")
    print("-"*80)
    
    if sucesso_cpu and resultado_cpu:
        print("✅ get_processor_info está funcionando corretamente.")
    else:
        print("❌ get_processor_info NÃO está retornando os dados esperados.")
    
    if sucesso_disco and resultado_disco:
        print("✅ get_disk_info está funcionando corretamente.")
    else:
        print("❌ get_disk_info NÃO está retornando os dados esperados.")
    
    print("\n" + "-"*80)
    print("PRÓXIMOS PASSOS")
    print("-"*80)
    print("1. Verifique as mensagens de erro acima para identificar o problema.")
    print("2. Se houver erros de permissão, execute como administrador.")
    print("3. Se os comandos diretos funcionarem, o problema está no processamento dos dados.")
    print("4. Verifique se há atualizações disponíveis para o CSInfo.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador (ou deixe em branco para local): ").strip() or None
    
    main(computador)
