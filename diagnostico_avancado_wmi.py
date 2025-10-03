import wmi
import sys
import json
import platform
import subprocess
from ctypes import windll, byref, create_unicode_buffer, create_string_buffer

def testar_wmi_direto(computador):
    print(f"\n{'='*80}")
    print(f"TESTANDO WMI DIRETO EM {computador}")
    print(f"{'='*80}")
    
    resultados = {
        'conexao': False,
        'processador': None,
        'discos': None,
        'erro': None
    }
    
    try:
        # Conecta ao WMI do computador remoto
        print(f"Conectando a {computador} via WMI...")
        conexao = wmi.WMI(computer=computador)
        print("✅ Conexão WMI bem-sucedida!")
        resultados['conexao'] = True
        
        # Testa consulta ao processador
        print("\nTestando consulta ao processador (Win32_Processor)...")
        try:
            cpus = conexao.Win32_Processor()
            print(f"✅ Encontrado(s) {len(cpus)} processador(es):")
            resultados['processador'] = []
            
            for i, cpu in enumerate(cpus, 1):
                cpu_info = {
                    'nome': cpu.Name.strip() if cpu.Name else 'Nome não disponível',
                    'nucleos': cpu.NumberOfCores,
                    'threads': cpu.NumberOfLogicalProcessors,
                    'clock_mhz': cpu.MaxClockSpeed,
                    'arquitetura': cpu.Architecture
                }
                resultados['processador'].append(cpu_info)
                
                print(f"  CPU {i}: {cpu_info['nome']}")
                print(f"     Arquitetura: {cpu_info['arquitetura']}")
                print(f"     Núcleos: {cpu_info['nucleos']}")
                print(f"     Threads: {cpu_info['threads']}")
                print(f"     Clock: {cpu_info['clock_mhz']} MHz")
                
        except Exception as e:
            msg = f"Erro ao consultar Win32_Processor: {str(e)}"
            print(f"❌ {msg}")
            resultados['erro'] = msg
        
        # Testa consulta aos discos
        print("\nTestando consulta aos discos (Win32_DiskDrive)...")
        try:
            discos = conexao.Win32_DiskDrive()
            print(f"✅ Encontrado(s) {len(discos)} disco(s):")
            resultados['discos'] = []
            
            for i, disco in enumerate(discos, 1):
                tamanho_gb = round(int(disco.Size or 0) / (1024**3), 2)
                
                disco_info = {
                    'modelo': disco.Model or 'Desconhecido',
                    'tamanho_gb': tamanho_gb,
                    'interface': disco.InterfaceType or 'Desconhecida',
                    'particoes': disco.Partitions or 0,
                    'serial': getattr(disco, 'SerialNumber', 'N/A')
                }
                resultados['discos'].append(disco_info)
                
                print(f"  Disco {i}: {disco_info['modelo']}")
                print(f"     Tamanho: {disco_info['tamanho_gb']} GB")
                print(f"     Interface: {disco_info['interface']}")
                print(f"     Partições: {disco_info['particoes']}")
                print(f"     Número de Série: {disco_info['serial']}")
                
        except Exception as e:
            msg = f"Erro ao consultar Win32_DiskDrive: {str(e)}"
            print(f"❌ {msg}")
            resultados['erro'] = msg
        
    except Exception as e:
        msg = f"Falha na conexão WMI: {str(e)}"
        print(f"❌ {msg}")
        resultados['erro'] = msg
        import traceback
        traceback.print_exc()
    
    return resultados

def testar_powershell(computador):
    print(f"\n{'='*80}")
    print(f"TESTANDO POWERSHELL EM {computador}")
    print(f"{'='*80}")
    
    resultados = {
        'processador': None,
        'discos': None,
        'erro': None
    }
    
    # Comando para obter informações do processador
    cmd_cpu = """
    try {
        $cpu = Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed
        $cpu | ConvertTo-Json -Depth 5 -Compress
    } catch {
        Write-Output "ERRO: $($_.Exception.Message)"
    }
    """
    
    # Comando para obter informações dos discos
    cmd_disco = """
    try {
        $discos = Get-WmiObject -Class Win32_DiskDrive | Select-Object Model, Size, Partitions, InterfaceType, MediaType, SerialNumber
        $discos | ConvertTo-Json -Depth 5 -Compress
    } catch {
        Write-Output "ERRO: $($_.Exception.Message)"
    }
    """
    
    try:
        # Testa consulta ao processador
        print("\nExecutando consulta ao processador via PowerShell...")
        resultado = subprocess.run(
            ["powershell", "-Command", cmd_cpu],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if resultado.returncode == 0 and not resultado.stdout.startswith("ERRO:"):
            print("✅ Consulta ao processador bem-sucedida!")
            resultados['processador'] = json.loads(resultado.stdout)
            print(json.dumps(resultados['processador'], indent=2))
        else:
            erro = resultado.stderr or resultado.stdout
            print(f"❌ Falha na consulta ao processador: {erro}")
            resultados['erro'] = f"PowerShell (CPU): {erro}"
        
        # Testa consulta aos discos
        print("\nExecutando consulta aos discos via PowerShell...")
        resultado = subprocess.run(
            ["powershell", "-Command", cmd_disco],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if resultado.returncode == 0 and not resultado.stdout.startswith("ERRO:"):
            print("✅ Consulta aos discos bem-sucedida!")
            resultados['discos'] = json.loads(resultado.stdout)
            print(json.dumps(resultados['discos'], indent=2))
        else:
            erro = resultado.stderr or resultado.stdout
            print(f"❌ Falha na consulta aos discos: {erro}")
            resultados['erro'] = f"PowerShell (Discos): {erro}"
            
    except Exception as e:
        msg = f"Erro ao executar comandos PowerShell: {str(e)}"
        print(f"❌ {msg}")
        resultados['erro'] = msg
    
    return resultados

def verificar_permissoes_wmi():
    print(f"\n{'='*80}")
    print("VERIFICANDO PERMISSÕES WMI")
    print(f"{'='*80}")
    
    try:
        # Verifica se o usuário atual tem permissões administrativas
        is_admin = windll.shell32.IsUserAnAdmin() != 0
        print(f"Usuário é administrador: {'✅ Sim' if is_admin else '❌ Não'}")
        
        # Tenta acessar o namespace WMI raiz
        try:
            wmi.WMI()
            print("✅ Acesso ao namespace WMI raiz bem-sucedido")
            return True
        except Exception as e:
            print(f"❌ Falha ao acessar o namespace WMI raiz: {str(e)}")
            return False
            
    except Exception as e:
        print(f"❌ Erro ao verificar permissões: {str(e)}")
        return False

def main(computador):
    print(f"\n{'='*80}")
    print(f"DIAGNÓSTICO AVANÇADO WMI - {computador.upper() if computador else 'LOCAL'}")
    print(f"{'='*80}")
    
    # Verifica permissões locais primeiro
    if not computador or computador.lower() in ('localhost', '127.0.0.1', '.'):
        print("\nVerificando permissões locais...")
        tem_permissoes = verificar_permissoes_wmi()
        
        if not tem_permissoes:
            print("\n⚠️  Execute este script como administrador para obter melhores resultados.")
    
    # Testa WMI direto
    print("\n" + "="*80)
    print("TESTE DIRETO VIA WMI")
    print("="*80)
    
    resultados_wmi = testar_wmi_direto(computador if computador and computador.lower() not in ('localhost', '127.0.0.1', '.') else None)
    
    # Testa PowerShell
    print("\n" + "="*80)
    print("TESTE VIA POWERSHELL")
    print("="*80)
    
    resultados_ps = testar_powershell(computador if computador and computador.lower() not in ('localhost', '127.0.0.1', '.') else 'localhost')
    
    # Análise dos resultados
    print("\n" + "="*80)
    print("ANÁLISE DOS RESULTADOS")
    print("="*80)
    
    if resultados_wmi.get('erro') or resultados_ps.get('erro'):
        print("\n⚠️  Foram encontrados erros durante os testes:")
        if resultados_wmi.get('erro'):
            print(f"- WMI: {resultados_wmi['erro']}")
        if resultados_ps.get('erro'):
            print(f"- PowerShell: {resultados_ps['erro']}")
    
    if resultados_wmi.get('conexao', False):
        print("\n✅ Conexão WMI estabelecida com sucesso!")
    else:
        print("\n❌ Não foi possível estabelecer conexão WMI.")
    
    if resultados_wmi.get('processador') or resultados_ps.get('processador'):
        print("✅ Informações do processador obtidas com sucesso!")
    else:
        print("❌ Não foi possível obter informações do processador.")
    
    if resultados_wmi.get('discos') or resultados_ps.get('discos'):
        print("✅ Informações dos discos obtidas com sucesso!")
    else:
        print("❌ Não foi possível obter informações dos discos.")
    
    # Recomendações
    print("\n" + "="*80)
    print("RECOMENDAÇÕES")
    print("="*80)
    
    if not resultados_wmi.get('conexao', False):
        print("""
1. Verifique se o serviço 'Windows Management Instrumentation' está em execução no computador remoto.
2. Verifique se as regras de firewall permitem o tráfego WMI (porta 135 e portas dinâmicas RPC).
3. Verifique se a conta tem permissões administrativas no computador remoto.
""")
    
    if resultados_wmi.get('conexao') and (not resultados_wmi.get('processador') or not resultados_wmi.get('discos')):
        print("""
1. O WMI está acessível, mas algumas consultas estão falhando. Isso pode indicar:
   - Corrupção no repositório WMI (tente 'winmgmt /resetrepository' no computador remoto)
   - Falta de permissões para consultas específicas
   - Problemas com provedores WMI específicos
""")
    
    if resultados_ps.get('processador') or resultados_ps.get('discos'):
        print("""
✅ O PowerShell consegue obter as informações corretamente, mas o CSInfo não.
   Isso indica que o problema está no processamento dos dados pelo CSInfo.
   Considere atualizar o CSInfo ou reportar o problema aos desenvolvedores.
""")
    
    print("\n" + "="*80)
    print("DIAGNÓSTICO CONCLUÍDO")
    print("="*80)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        computador = sys.argv[1]
    else:
        computador = input("Digite o nome do computador (ou deixe em branco para local): ").strip() or None
    
    main(computador)
