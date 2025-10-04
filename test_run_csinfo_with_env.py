import os
import sys
import traceback

def print_result(r):
    if isinstance(r, dict) and 'lines' in r:
        lines = r.get('lines') or []
        print('LINES:', len(lines))
        for i, l in enumerate(lines[:200]):
            print(f'{i+1:03}: {l}')
    else:
        print('RETURN:', repr(r))


def barra_cb(progress, line=None):
    if line:
        print('PROG:', progress, line)


if __name__ == '__main__':
    try:
        # Ler credenciais do ambiente se fornecidas
        user = os.environ.get('CSINFO_USER')
        pwd = os.environ.get('CSINFO_PASS')

        if user and pwd:
            # Import set_default_credential diretamente do módulo interno
            try:
                from csinfo._impl import set_default_credential
                ok = set_default_credential(user, pwd)
                print(f'set_default_credential -> {ok}')
            except Exception as e:
                print('Aviso: não foi possível definir credencial via set_default_credential:', e)

        # aceitar nome da máquina como argumento opcional
        machine = None
        if len(sys.argv) > 1:
            machine = sys.argv[1]

        import csinfo

        print(f'Calling csinfo.main(export_type="gui") for machine: {machine}')
        r = csinfo.main(export_type='gui', barra_callback=barra_cb, computer_name=machine)
        print('OK: returned')
        print_result(r)
    except Exception:
        traceback.print_exc()
