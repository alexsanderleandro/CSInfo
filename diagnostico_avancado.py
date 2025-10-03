import sys
import os
import json
from csinfo._impl import get_processor_info, get_disk_info

def testar_wmi(computador):
    """Testa a conexão WMI básica com o computador remoto."""
    print("\n=== TESTE DE CONEXÃO WMI ===")
    
    try:
        import wmi
        print("✅ Módulo wmi importado com sucesso")
        
        print(f"\nTentando conectar a {computador} via WMI...")
        conexao = wmi.WMI(computer=computador)
        print("✅ Conexão WMI bem-sucedida!")
        
        # Testa consultas básicas
        print("\nTestando consultas WMI...")
        
        # Testa Win32_Processor
        try:
            cpus = conexao.Win32_Processor()
            print(f"✅ Win32_Processor: Encontrado {len(cpus)} processador(es)")
            for i, cpu in enumerate(cpus, 1):
                print(f"  CPU {i}: {cpu.Name.strip() if cpu.Name else 'Nome não disponível'}")
        except Exception as e:
            print(f"❌ Erro ao consultar Win32_Processor: {str(e)}")
        
        # Testa Win32_DiskDrive
        try:
            discos = conexao.Win32_DiskDrive()
            print(f"✅ Win32_DiskDrive: Encontrado {len(discos)} disco(s)")
            for i, disco in enumerate(discos, 1):
                print(f"  Disco {i}: {disco.Caption or 'Sem descrição'} - {round(int(disco.Size or 0)/1e9, 2)} GB")
        except Exception as e:
            print(f"❌ Erro ao consultar Win32_DiskDrive: {str(e)}")
            
        return True
        
    except Exception as e:
        print(f"❌ Falha na conexão WMI: {str(e)}")
        return False

def testar_funcao(nome, funcao, *args):
    print(f"\n=== TESTANDO: {nome} ===")
    try:
        resultado = funcao(*args)
        print(f"✅ Sucesso! Retornou: {type(resultado)}")
        
        if resultado is not None:
            if isinstance(resultado, (list, tuple)):
                print(f"   Itens retornados: {len(resultado)}")
                for i, item in enumerate(resultado[:3], 1):  # Mostra até 3 itens
                    print(f"   Item {i}: {type(item)} - {str(item)[:100]}")
                    if hasattr(item, 'items'):
                        for k, v in list(item.items())[:3]:  # Mostra até 3 chaves/valores
                            print(f"     {k}: {v}")
                        if len(item) > 3:
                            print(f"     ... e mais {len(item)-3} itens")
            elif isinstance(resultado, dict):
                print(f"   Chaves retornadas: {list(resultado.keys())}")
                for k, v in list(resultado.items())[:3]:  # Mostra até 3 itens
                    print(f"   {k}: {v}")
                if len(resultado) > 3:
                    print(f"   ... e mais {len(resultado)-3} itens")
            else:
                print(f"   Valor: {resultado}")
        
        return True, resultado
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, None

def main(computador):
    print(f"\n{'='*50}")
    print(f"DIAGNÓSTICO AVANÇADO PARA {computador}")
    print(f"{'='*50}")
    
    # Testa conexão WMI básica
    wmi_ok = testar_wmi(computador)
    
    if not wmi_ok:
        print("\n⚠️  A conexão WMI básica falhou. Verifique:")
        print("1. Se o serviço 'Windows Management Instrumentation' está em execução no computador remoto")
        print("2. Se a conta tem permissões administrativas no computador remoto")
        print("3. Se as regras de firewall permitem o tráfego WMI (porta 135 e portas dinâmicas RPC)")
    
    # Testa funções específicas com mais detalhes
    print("\n" + "="*50)
    print("TESTANDO FUNÇÕES ESPECÍFICAS")
    print("="*50)
    
    # Testa get_processor_info
    sucesso_cpu, resultado_cpu = testar_funcao("get_processor_info", get_processor_info, computador)
    
    # Testa get_disk_info
    sucesso_disco, resultado_disco = testar_funcao("get_disk_info", get_disk_info, computador)
    
    # Se ambas falharam, tenta com credenciais explícitas
    if not sucesso_cpu or not sucesso_disco:
        print("\n" + "="*50)
        print("TENTATIVA COM CREDENCIAIS EXPLÍCITAS")
        print("="*50)
        
        import getpass
        print("\nPor favor, insira as credenciais para tentar novamente:")
        usuario = input("Usuário (DOMÍNIO\\usuário ou .\\usuário para local): ").strip()
        senha = getpass.getpass("Senha: ")
        
        # Cria uma nova conexão WMI com credenciais
        try:
            import wmi
            conexao = wmi.WMI(computer=computador, user=usuario, password=senha)
            print("✅ Conexão WMI com credenciais bem-sucedida!")
            
            # Testa novamente com a nova conexão
            print("\nTestando consultas WMI com credenciais...")
            
            # Testa Win32_Processor
            try:
                cpus = conexao.Win32_Processor()
                print(f"✅ Win32_Processor: Encontrado {len(cpus)} processador(es)")
                for i, cpu in enumerate(cpus, 1):
                    print(f"  CPU {i}: {cpu.Name.strip() if cpu.Name else 'Nome não disponível'}")
            except Exception as e:
                print(f"❌ Erro ao consultar Win32_Processor: {str(e)}")
            
            # Testa Win32_DiskDrive
            try:
                discos = conexao.Win32_DiskDrive()
                print(f"✅ Win32_DiskDrive: Encontrado {len(discos)} disco(s)")
                for i, disco in enumerate(discos, 1):
                    print(f"  Disco {i}: {disco.Caption or 'Sem descrição'} - {round(int(disco.Size or 0)/1e9, 2)} GB")
            except Exception as e:
                print(f"❌ Erro ao consultar Win32_DiskDrive: {str(e)}")
            
        except Exception as e:
            print(f"❌ Falha na conexão WMI com credenciais: {str(e)}")
    
    print("\n" + "="*50)
    print("DIAGNÓSTICO CONCLUÍDO")
    print("="*50)
    
    # Salva o relatório
    relatorio = {
        "computador": computador,
        "data_hora": str(datetime.datetime.now()),
        "wmi_funcional": wmi_ok,
        "get_processor_info": {
            "sucesso": sucesso_cpu,
            "resultado": str(resultado_cpu)[:1000] if resultado_cpu else None
        },
        "get_disk_info": {
            "sucesso": sucesso_disco,
            "resultado": str(resultado_disco)[:1000] if resultado_disco else None
        }
    }
    
    nome_arquivo = f"diagnostico_avancado_{computador}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        json.dump(relatorio, f, indent=2, ensure_ascii=False)
    
    print(f"\nRelatório salvo em: {os.path.abspath(nome_arquivo)}")

if __name__ == "__main__":
    import datetime
    
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador: ").strip()
    
    main(computador)
