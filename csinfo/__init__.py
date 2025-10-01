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
)

__all__ = [
    'main',
    'write_report',
    'write_pdf_report',
    'safe_filename',
    'get_machine_name',
    'check_remote_machine',
    'remove_duplicate_lines',
]

__version__ = '0.1.0'
