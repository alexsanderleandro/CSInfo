# Script de inspeção para CSInfo GUI widgets
# Salve e execute com: python scripts\inspect_gui.py

import traceback
import os
import sys

# Garantir que o diretório raiz do projeto esteja no sys.path
ROOT = os.path.dirname(os.path.dirname(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import csinfo_gui

try:
    app = csinfo_gui.CSInfoApp()
    # Forçar atualização para avaliar manager/mapping
    app.update()

    def widget_info(w):
        try:
            exists = bool(w and w.winfo_exists())
        except Exception:
            exists = False
        try:
            manager = w.winfo_manager() if exists else None
        except Exception:
            manager = f"ERR:{traceback.format_exc().splitlines()[-1]}"
        # estado se for widget tipo Button
        state = None
        try:
            if exists and hasattr(w, 'cget'):
                try:
                    state = w.cget('state')
                except Exception:
                    state = None
        except Exception:
            state = None
        try:
            mapped = w.winfo_ismapped() if exists else False
        except Exception:
            mapped = False
        return {'exists': exists, 'manager': manager, 'state': state, 'mapped': mapped}

    names = [
        'export_btn', 'exit_btn', 'frame_botoes', 'frame_export',
        'left_frame', 'right_frame', 'frame_result', 'info_text', 'machines_listbox'
    ]

    print('--- CSInfoApp widget inspection ---')
    for n in names:
        w = getattr(app, n, None)
        info = widget_info(w) if w is not None else None
        print(f"{n}: {info}")

except Exception as e:
    print('Erro ao instanciar/inspecionar a GUI:')
    traceback.print_exc()
finally:
    try:
        # destruir a janela caso exista
        app.destroy()
    except Exception:
        pass

print('\nScript finalizado')
