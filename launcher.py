#!/usr/bin/env python3
r"""Launcher leve para CSInfo

Mostra uma splash imediata 'Carregando informações' enquanto inicia o executável principal
e espera pela criação de %TEMP%\csinfo_gui_debug.log para fechar a splash.

Uso: python launcher.py [--target PATH_TO_EXE] [--timeout SECONDS]
Se nenhum target for passado, tenta 'dist\csinfo.exe' e 'dist\csinfo_dir\csinfo_dir.exe'.
"""
import os
import sys
import threading
import time
import subprocess
import tkinter as tk
from tkinter import ttk

DEFAULT_CANDIDATES = [
    os.path.join('dist', 'csinfo.exe'),
    os.path.join('dist', 'csinfo_dir', 'csinfo_dir.exe'),
    os.path.join('dist', 'csinfo_with_appico.exe'),
]

def find_target(arg_path=None):
    """Procurar o executável alvo em caminhos razoáveis.

    Procura tanto relativo ao diretório atual quanto relativo ao diretório do
    próprio launcher (útil quando o usuário dá duplo-clique dentro de dist\).
    """
    if arg_path:
        if os.path.exists(arg_path):
            return arg_path
        cand = os.path.abspath(arg_path)
        if os.path.exists(cand):
            return cand
        return None

    exe_dir = None
    try:
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
    except Exception:
        exe_dir = None
    if not exe_dir:
        exe_dir = os.getcwd()

    # candidates to try: as-is (relative to cwd), relative to exe_dir, and parent dist
    tried = []
    parent = os.path.dirname(exe_dir)
    for c in DEFAULT_CANDIDATES:
        # as-is (relative to current working dir)
        tried.append(c)
        # relative to exe_dir (when launcher and app are colocados juntos)
        tried.append(os.path.join(exe_dir, os.path.basename(c)))
        tried.append(os.path.join(exe_dir, c))
        # sibling inside parent (ex: dist\csinfo.exe when launcher is in dist\csinfo_launcher)
        tried.append(os.path.join(parent, os.path.basename(c)))
        # sibling nested (ex: parent\csinfo_dir\csinfo_dir.exe)
        tried.append(os.path.join(parent, os.path.dirname(c), os.path.basename(c)))

    # also try common names directly in exe_dir
    tried.append(os.path.join(exe_dir, 'csinfo.exe'))
    tried.append(os.path.join(exe_dir, 'csinfo_dir', 'csinfo_dir.exe'))

    for t in tried:
        try:
            if t and os.path.exists(t):
                return t
        except Exception:
            continue
    return None

def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--target', help='Caminho do executável alvo (padrão: dist\\csinfo.exe)')
    p.add_argument('--timeout', type=float, default=60.0, help='Timeout em segundos para esperar o log')
    args = p.parse_args()

    target = find_target(args.target)
    if not target:
        # Sem console: mostrar uma mensagem amigável ao usuário e gravar em TEMP
        try:
            from tkinter import messagebox
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror('CSInfo Launcher', 'Executável alvo não encontrado. Gere o build em dist\\ e tente novamente.')
            try:
                root_err.destroy()
            except Exception:
                pass
        except Exception:
            pass
        try:
            fn = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')), 'csinfo_launcher_debug.log')
            with open(fn, 'a', encoding='utf-8') as fh:
                fh.write('Target not found\n')
        except Exception:
            pass
        sys.exit(2)

    log_path = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')), 'csinfo_gui_debug.log')
    # remover log antigo para medição limpa
    try:
        if os.path.exists(log_path):
            os.remove(log_path)
    except Exception:
        pass

    root = tk.Tk()
    root.title('CSInfo - Carregando informações')
    root.geometry('420x100')
    root.resizable(False, False)

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill='both', expand=True)
    lbl = ttk.Label(frm, text='Carregando informações', font=('Segoe UI', 11))
    lbl.pack(pady=(6,8))
    pb = ttk.Progressbar(frm, orient='horizontal', mode='indeterminate', length=360)
    pb.pack()
    try:
        pb.start(60)
    except Exception:
        pass

    start_time = time.time()

    # iniciar o executável alvo
    try:
        # No Windows, evitar janela de console para o processo filho
        creationflags = 0
        if sys.platform.startswith('win'):
            try:
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
            except Exception:
                creationflags = 0
        proc = subprocess.Popen([os.path.abspath(target)], close_fds=True, creationflags=creationflags)
    except Exception as e:
        # Mostrar erro via caixa de diálogo para o usuário quando não há console
        try:
            from tkinter import messagebox
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror('CSInfo Launcher', f'Falha ao iniciar o executável alvo:\n{e}')
            try:
                root_err.destroy()
            except Exception:
                pass
        except Exception:
            pass
        try:
            fn = os.path.join(os.environ.get('TEMP', os.path.expanduser('~')), 'csinfo_launcher_debug.log')
            with open(fn, 'a', encoding='utf-8') as fh:
                fh.write(f'Failed to start target: {e}\n')
        except Exception:
            pass
        root.destroy()
        sys.exit(3)

    closed = threading.Event()

    def waiter():
        # Esperar pela criação do log ou pelo término do processo
        timeout = float(args.timeout or 60.0)
        poll = 0.2
        elapsed = 0.0
        while elapsed < timeout and not closed.is_set():
            if os.path.exists(log_path):
                break
            if proc.poll() is not None:
                # processo terminou
                break
            time.sleep(poll)
            elapsed += poll
        # fechar splash
        try:
            root.after(100, root.destroy)
        except Exception:
            try:
                root.destroy()
            except Exception:
                pass

    t = threading.Thread(target=waiter, daemon=True)
    t.start()

    try:
        root.mainloop()
    except Exception:
        pass

    # opcional: não encerrar o processo filho (deixar rodando)

if __name__ == '__main__':
    main()
