"""csinfo package — API pública.

Este arquivo expõe as funções públicas do pacote (main, write_report,
write_pdf_report, safe_filename, get_machine_name, check_remote_machine, etc.)
redirecionando para a implementação interna em ``_impl.py``. Isso garante que
``import csinfo`` funcione de forma previsível tanto em desenvolvimento quanto
quando empacotado por ferramentas como o PyInstaller.
"""

from ._impl import (
    main as main,
    write_report as write_report,
    write_pdf_report as write_pdf_report,
    safe_filename as safe_filename,
    get_machine_name as get_machine_name,
    check_remote_machine as check_remote_machine,
    remove_duplicate_lines as remove_duplicate_lines,
    set_default_credential as set_default_credential,
    clear_default_credential as clear_default_credential,
    get_debug_session_log as get_debug_session_log,
)

__all__ = [
    'main',
    'write_report',
    'write_pdf_report',
    'safe_filename',
    'get_machine_name',
    'check_remote_machine',
    'remove_duplicate_lines',
    'set_default_credential',
    'clear_default_credential',
    'get_debug_session_log',
]

__version__ = '1.0.2'
