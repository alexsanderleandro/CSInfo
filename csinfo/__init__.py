"""Public API for the csinfo package.

This module re-exports selected helpers implemented in :mod:`csinfo._impl` so that
code doing `import csinfo` keeps working and finds the expected symbols (main,
write_report, write_pdf_report, set_default_credential, clear_default_credential,
safe_filename, etc.).
"""

from ._impl import (
	main,
	write_report,
	write_pdf_report,
	set_default_credential,
	clear_default_credential,
	safe_filename,
	get_machine_name,
)

# Configurações públicas que podem ser sobrescritas pelo usuário do pacote
# Controla se a barra lateral do PDF é desenhada no canvas (True por padrão)
__pdf_sidebar_enabled__ = False
# Mapeamento opcional de cores por campo usado no PDF: {'Nome do computador': '#FF0000', ...}
__pdf_field_colors__ = {}

__all__ = [
	'main',
	'write_report',
	'write_pdf_report',
	'set_default_credential',
	'clear_default_credential',
	'safe_filename',
	'get_machine_name',
]
