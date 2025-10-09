"""Limpeza de artefatos temporários do workspace CSInfo.

Uso:
  python scripts/cleanup_temp.py [--yes] [--patterns PATTERN ...]

Opções:
  --yes       Executa a remoção (por padrão o script faz dry-run e apenas lista)
  --patterns  Lista adicional de padrões (glob) para remover

O script procura recursivamente arquivos que correspondam a padrões comuns de
arquivos temporários/artefatos gerados pelo processo de build/test do projeto
(e.g. *.tmp, *.fixed, *.annots.jsonl) e os remove. Por segurança ele preserva
arquivos finais óbvios como `test_report.pdf` e quaisquer arquivos em `scripts/`.
"""
import os
import sys
import fnmatch
from pathlib import Path

DEFAULT_PATTERNS = [
    "*.tmp",
    "*.tmp.*",
    "*.fixed",
    "*.annots.jsonl",
    "*.tmp.annots.jsonl",
    "*.tmp.pdf",
]

ROOT = Path(__file__).resolve().parents[1]

def find_candidates(root, patterns):
    candidates = []
    for dirpath, dirnames, filenames in os.walk(root):
        # skip .git and __pycache__
        if '.git' in dirnames:
            dirnames.remove('.git')
        if '__pycache__' in dirnames:
            dirnames.remove('__pycache__')
        # do not remove files inside scripts/ (we keep test helpers)
        rel = os.path.relpath(dirpath, root)
        for p in patterns:
            for name in fnmatch.filter(filenames, p):
                full = Path(dirpath) / name
                # safety: preserve main output
                if full.name == 'test_report.pdf':
                    continue
                # preserve anything inside scripts/ explicitly
                if rel.startswith('scripts'):
                    continue
                candidates.append(full)
    # deduplicate and sort
    uniq = sorted(set(candidates), key=lambda p: str(p))
    return uniq


def main(argv):
    args = argv[1:]
    do_delete = False
    patterns = list(DEFAULT_PATTERNS)
    i = 0
    while i < len(args):
        a = args[i]
        if a == '--yes':
            do_delete = True
            i += 1
            continue
        if a == '--patterns':
            i += 1
            if i < len(args):
                patterns.append(args[i])
            i += 1
            continue
        # allow passing extra patterns directly
        if a.startswith('--'):
            print('Unknown flag', a)
            return 1
        patterns.append(a)
        i += 1

    print('Workspace root:', ROOT)
    print('Patterns:', patterns)
    candidates = find_candidates(ROOT, patterns)
    if not candidates:
        print('Nenhum artefato temporário encontrado para remoção.')
        return 0

    print('\nArquivos candidatos:')
    for p in candidates:
        print(' -', p)

    if not do_delete:
        print('\nDry-run (nenhum arquivo foi removido). Rerun com --yes para apagar.')
        return 0

    removed = []
    failed = []
    for p in candidates:
        try:
            p.unlink()
            removed.append(p)
        except Exception as e:
            failed.append((p, str(e)))

    print('\nRemovidos:')
    for p in removed:
        print(' -', p)
    if failed:
        print('\nFalhas:')
        for p, e in failed:
            print(' -', p, e)
        return 2
    return 0

if __name__ == '__main__':
    sys.exit(main(sys.argv))
