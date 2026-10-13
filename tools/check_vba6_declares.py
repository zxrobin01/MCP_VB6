"""Gate: every function declared Lib "VBA6.DLL" in the VB6 source must exist in the real
export table of the VBA6.DLL of the VB6 installation being used.

The source file is CONFIGURABLE, because the path a maintainer uses is not necessarily the
path a reader of this tree has:

    python tools/check_vba6_declares.py                  # default: the shipped module
    python tools/check_vba6_declares.py --src <path>     # explicit source file
    VB6MCP_DECLARES_SRC=<path> python tools/check_vba6_declares.py

The VB6 installation is taken from VB6MCP_VB98, else the usual 32-bit VB98 directory.

Exit codes: 0 = every declare is exported | 1 = at least one is missing |
            2 = an input was not found (a missing input is NOT a passing verdict).
"""
import os, re, struct, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_SRC = os.path.join(ROOT, "Vba6Probe.bas")
DEFAULT_VB98 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98"
DECLARE_RE = r'Declare\s+Function\s+(\w+)\s+Lib\s+"VBA6\.DLL"'


def pick_src(argv):
    """--src wins, then $VB6MCP_DECLARES_SRC, then the shipped default; None if unusable."""
    if "--src" in argv:
        i = argv.index("--src")
        return argv[i + 1] if i + 1 < len(argv) else None
    env = os.environ.get("VB6MCP_DECLARES_SRC")
    return env if env else DEFAULT_SRC


def exports(path):
    d = open(path, "rb").read()
    e = struct.unpack_from("<I", d, 0x3C)[0]
    coff = e + 4
    nsec = struct.unpack_from("<H", d, coff + 2)[0]
    optsz = struct.unpack_from("<H", d, coff + 16)[0]
    opt = coff + 20
    magic = struct.unpack_from("<H", d, opt)[0]
    dd = opt + (96 if magic == 0x10b else 112)
    exp = struct.unpack_from("<I", d, dd)[0]
    secs = []
    so = opt + optsz
    for i in range(nsec):
        o = so + i * 40
        vsize, vaddr, rawsize, rawptr = struct.unpack_from("<IIII", d, o + 8)
        secs.append((vaddr, vsize, rawptr, rawsize))
    def r2o(rva):
        for vaddr, vsize, rawptr, rawsize in secs:
            if vaddr <= rva < vaddr + max(vsize, rawsize):
                return rawptr + (rva - vaddr)
        return None
    eo = r2o(exp)
    nnames = struct.unpack_from("<I", d, eo + 24)[0]
    an = struct.unpack_from("<I", d, eo + 32)[0]
    no = r2o(an)
    out = set()
    for i in range(nnames):
        nr = struct.unpack_from("<I", d, no + 4 * i)[0]
        o = r2o(nr)
        out.add(d[o:d.index(b"\0", o)].decode("latin1"))
    return out


def main():
    src_path = pick_src(sys.argv[1:])
    if not src_path or not os.path.isfile(src_path):
        sys.stderr.write("SOURCE_NOT_FOUND: %s\n" % (src_path,))
        sys.stderr.write("  pass --src <path> or set VB6MCP_DECLARES_SRC to the VB6 module\n")
        sys.stderr.write("  that declares the VBA6.DLL functions (default: %s)\n" % DEFAULT_SRC)
        return 2
    vb98 = os.environ.get("VB6MCP_VB98") or DEFAULT_VB98
    dll = os.path.join(vb98, "VBA6.DLL")
    if not os.path.isfile(dll):
        sys.stderr.write("VBA6_DLL_NOT_FOUND: %s\n" % dll)
        sys.stderr.write("  set VB6MCP_VB98 to the directory that contains VBA6.DLL\n")
        return 2
    src = open(src_path, encoding="utf-8", errors="replace").read()
    names = re.findall(DECLARE_RE, src, re.I)
    ex = exports(dll)
    missing = [n for n in names if n not in ex]
    print("SOURCE=%s" % src_path)
    print("DLL=%s" % dll)
    print("DECLARES=%d EXPORTS=%d" % (len(names), len(ex)))
    for n in names:
        print("%s|%s" % ("FOUND" if n in ex else "MISSING", n))
    print("SUMMARY declares=%d missing=%d" % (len(names), len(missing)))
    return 1 if missing else 0


raise SystemExit(main())
