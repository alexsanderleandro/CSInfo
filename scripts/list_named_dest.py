import sys
p = sys.argv[1] if len(sys.argv)>1 else 'test_report.pdf'
try:
    from pypdf import PdfReader
    rdr = PdfReader(p)
    nd = getattr(rdr, 'named_destinations', None) or {}
    print('using pypdf, named_destinations keys:', list(nd.keys()))
except Exception:
    try:
        import PyPDF2
        rdr = PyPDF2.PdfReader(p)
        nd = getattr(rdr, 'namedDestinations', None) or {}
        print('using PyPDF2, namedDestinations keys:', list(nd.keys()))
    except Exception as e:
        print('no pdf lib', e)
        sys.exit(2)

# try to print mapping
for name in nd:
    try:
        dest = nd[name]
        print(name, '->', dest)
    except Exception:
        print('err getting', name)
