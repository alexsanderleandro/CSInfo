import csinfo, traceback, json, sys


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
        machine = None
        if len(sys.argv) > 1:
            machine = sys.argv[1]
        print(f'Calling csinfo.main(export_type="gui") for machine: {machine}')
        r = csinfo.main(export_type='gui', barra_callback=barra_cb, computer_name=machine)
        print('OK: returned')
        print_result(r)
    except Exception:
        traceback.print_exc()
