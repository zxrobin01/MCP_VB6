import re, struct, os

def exports(path):
    d = open(path, 'rb').read()
    e = struct.unpack_from('<I', d, 0x3C)[0]
    coff = e + 4
    nsec = struct.unpack_from('<H', d, coff + 2)[0]
    optsz = struct.unpack_from('<H', d, coff + 16)[0]
    opt = coff + 20
    magic = struct.unpack_from('<H', d, opt)[0]
    dd = opt + (96 if magic == 0x10b else 112)
    exp = struct.unpack_from('<I', d, dd)[0]
    secs = []
    so = opt + optsz
    for i in range(nsec):
        o = so + i * 40
        vsize, vaddr, rawsize, rawptr = struct.unpack_from('<IIII', d, o + 8)
        secs.append((vaddr, vsize, rawptr, rawsize))
    def r2o(rva):
        for vaddr, vsize, rawptr, rawsize in secs:
            if vaddr <= rva < vaddr + max(vsize, rawsize):
                return rawptr + (rva - vaddr)
        return None
    eo = r2o(exp)
    nnames = struct.unpack_from('<I', d, eo + 24)[0]
    an = struct.unpack_from('<I', d, eo + 32)[0]
    no = r2o(an)
    out = set()
    for i in range(nnames):
        nr = struct.unpack_from('<I', d, no + 4 * i)[0]
        o = r2o(nr)
        out.add(d[o:d.index(b'\0', o)].decode('latin1'))
    return out


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    vb98 = os.environ.get('VB6MCP_VB98') or r'C:\Program Files (x86)\Microsoft Visual Studio\VB98'
    dll = os.path.join(vb98, 'VBA6.DLL')
    src = open(os.path.join(root, 'editsec', 'Vba6Probe.bas'), encoding='utf-8', errors='replace').read()
    names = re.findall(r'Declare\s+Function\s+(\w+)\s+Lib\s+"VBA6\.DLL"', src, re.I)
    ex = exports(dll)
    missing = [n for n in names if n not in ex]
    print('DECLARES=%d EXPORTS=%d' % (len(names), len(ex)))
    for n in names:
        print('%s|%s' % ('FOUND' if n in ex else 'MISSING', n))
    print('SUMMARY declares=%d missing=%d' % (len(names), len(missing)))
    return 1 if missing else 0


raise SystemExit(main())
