import sys
p = sys.argv[1] if len(sys.argv) > 1 else 'test_report.pdf'
print('inspecting', p)
try:
    import PyPDF2
    r = PyPDF2.PdfReader(p)
    root = r.trailer.get('/Root')
    if root is None:
        print('no /Root in trailer')
    else:
        try:
            print('/Root keys:', list(root.keys()))
        except Exception as e:
            print('could not list root keys:', e)
    names = None
    try:
        names = root.get('/Names') if root else None
        print('/Names present?', bool(names))
        if names:
            try:
                print('Names keys:', list(names.keys()))
            except Exception:
                print('Names raw:', names)
    except Exception as e:
        print('error reading /Names:', e)
    try:
        nd = getattr(r, 'named_destinations', None) or {}
        print('named_destinations keys (pypdf style):', list(nd.keys()) if nd else None)
    except Exception as e:
        print('named_destinations error:', e)
    try:
        nd2 = getattr(r, 'namedDestinations', None)
        print('namedDestinations attr (deprecated):', bool(nd2))
    except Exception as e:
        print('namedDestinations error (deprecated):', e)
except Exception as e:
    print('failed to open PDF with PyPDF2/pypdf:', e)
    sys.exit(2)
