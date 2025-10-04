import csinfo
import sys


def barra_callback(percent_or_none, line_or_stage):
    if percent_or_none is None:
        print(f"LINE: {line_or_stage}")
    else:
        try:
            p = int(percent_or_none)
        except Exception:
            p = percent_or_none
        print(f"PROGRESS: {p}% - {line_or_stage}")


if __name__ == '__main__':
    target = 'ceosoft-059'
    print(f"Iniciando coleta para: {target}")
    try:
        csinfo.main(export_type=None, barra_callback=barra_callback, computer_name=target)
        print("Coleta concluída com sucesso")
    except Exception as e:
        print("Erro durante coleta:", e)
        import traceback
        traceback.print_exc()
        sys.exit(1)
