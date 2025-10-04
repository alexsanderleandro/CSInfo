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

__all__ = [
	'main',
	'write_report',
	'write_pdf_report',
	'set_default_credential',
	'clear_default_credential',
	'safe_filename',
	'get_machine_name',
]
