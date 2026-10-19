#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Source stamp (algorithm identifier: vb6src/2) for a VB6 project.

WHY THIS EXISTS
    "Is this VB6MCP.dll built from this tree?" must be decidable, and it must NOT be
    answered by scanning strings out of the binary.  The answer is a single indicator I
    that is written INTO the artifact and READ OUT of it through a structured read
    point.

THE INVARIANT (one direction only -- corrected 2026-10-15 by the author)
        same source input set  =>  same I
        (I is a pure function of the input set: across repeated builds, across
         machines, and across VB6 service packs.  The BINARY may differ -- PE
         timestamps, COM GUIDs and the build machine are not part of I.)

    It is explicitly NOT required that a source edit changes I, and I is NOT a change
    detector for the build product: some source edits are invisible to the compiler, so
    the artifact is legitimately identical.  Read I as proof of SAME ORIGIN:
        I equal    -> the artifact is consistent with this input set (modulo a sha256
                      collision); the stamp is fresh
        I unequal  -> not this tree, or the stamp is stale
    I is nevertheless the sha256 of the input bytes, so any byte change in a *visible*
    input file does change it (that is collision-avoidance, not a promise about the
    compiled output).

INPUT SET (spec: docs/SOURCE_STAMP.md section 2) -- exhaustive, otherwise the "same
code => same I" direction breaks:
    1. MCP_VB6.vbp itself;
    2. the FILE of every component it lists (Form= / Class=X; X.cls / Module=X; X.bas /
       UserControl= / PropertyPage= / UserDocument= / Designer= / RelatedDoc= /
       ResFile32=; the part after ';' is the file name);
    3. each component's companion binary, when present: .frx / .res / .ctx / .dsr;
    4. EXCLUDED (they are the stamp's own outputs, or non-source noise):
       - the .vbp VersionComments= line (where the stamp is written back -- hashing it
         would be self-referential) and, defensively, a legacy invalid "Comments=" line;
       - *.bak*, *.obj, *.dll, *.exe, *.log, *.pdb, docs/**.

ALGORITHM (spec section 4)
    manifest line : <relative path> TAB <byte count> TAB <sha256(raw)>   (tab separated)
    ordering      : by the UTF-8 bytes of the relative path (cross-platform stable)
    header        : "vb6src/2" LF  +  mode name LF
    payload       : header + lines joined with LF + a final LF
    I             : sha256(payload) as lowercase hex

    The manifest file on disk IS the payload, so the useful identity
    I == sha256(bytes of the manifest file) holds and can be checked by hand.
    The algorithm identifier changes whenever the definition changes; moving to
    vb6src/2 requires recomputing every stamp.

MODES (spec section 3)
    strict   (default) : sha256 of the RAW bytes.  Any byte change (CRLF/LF, encoding,
                         comment, blank line) is visible in I.
    semantic (secondary): line endings normalised to LF, a leading UTF-8 BOM removed,
                         trailing blank lines collapsed.  Used only to separate a real
                         semantic change from formatting noise.  I is NEVER taken from
                         semantic mode; it is not written into any artifact.

WRITE POINTS (spec section 5) -- two redundant places, both excluded from the input set:
    1. runtime  : BuildStamp.bas with
                      Public Const VB6MCP_SOURCE_ALGO As String = "vb6src/2"
                      Public Const VB6MCP_SOURCE_HASH As String = "<I>"
                  which the add-in surfaces as mcp_info -> build.sourceHash/sourceAlgo
                  (structured read, no string parsing);
    2. offline  : the version resource, via the .vbp VersionComments= line
                  VersionComments=algo=<algo>;src=<I>
                  read with GetFileVersionInfoW + VerQueryValueW (the documented API),
                  cross-checked against an independent pure-python parse of RT_VERSION.
                  NOTE the deliberate name split: the .vbp KEY must be VersionComments
                  (VB6 rejects a bare "Comments" key as an invalid keyword and /make then
                  fails), while the version-resource FIELD read back is "Comments".
                  Write key and read field are different names on purpose -- do not
                  "align" them.

USAGE
    python tools/srcstamp.py compute <tree> [--mode strict|semantic] [--manifest PATH] [--json]
    python tools/srcstamp.py emit    <tree> [--vbp MCP_VB6.vbp] [--out BuildStamp.bas]
                                       (also sets Comments= in the --vbp file, unless
                                        --no-vbp-patch; both write points are idempotent)
    python tools/srcstamp.py check   <artifact.dll> <tree> [--mcp-url URL] [--baseline MANIFEST]
    python tools/srcstamp.py selftest <tree>

EXIT CODES
    0 PASS / success
    1 FAIL (check: stamp missing, stamp differs from the tree, or the two read points
            disagree)
    2 usage error (including "the .vbp input set is not exhaustive")
    3 refused target (port 8766 is reserved and is NEVER contacted) or I/O failure

PYTHON RUNTIME REQUIREMENT 
    This tool runs on BOTH Python 2.7 and Python 3.x - deliberately, because the build host
    (t16) only has a real 2.7 interpreter while this repository is developed on 3.x, and the
    same I must come out of either one (verified: identical SOURCE_HASH on 2.7.14 and 3.9.18).
    Portability rules for anyone editing it:
      * text-mode opens use io.open(path, "r", encoding=...) - `open(..., encoding=)` is py3-only;
      * `import io` (and every other import) must come AFTER `from __future__ import ...`,
        otherwise py2/py3 both raise "from __future__ imports must occur at the beginning";
      * HTTP uses the _UrlRequest/_urlopen/_urlsplit shims (urllib.request/urllib.parse are py3
        module layouts; py2.7 has urllib2/urlparse);
      * sys.stdout.reconfigure is py3.7+ only and is guarded by hasattr;
      * binary I/O (rb/wb) is unchanged and is what keeps the hash byte-identical across versions.

Known boundaries are listed in docs/evidence/vbmap/REPORT.md; the important one:
the stamp proves "the artifact is consistent with THIS source input set", not "the
build environment was clean".
"""
from __future__ import print_function

# NB: `import io` MUST come after the __future__ line - putting it before raises
# "SyntaxError: from __future__ imports must occur at the beginning of the file" (py2 and py3).
import argparse
import hashlib
import io
import json
import os
import re
import shutil
import struct
import sys
import tempfile

# py2/py3 module-layout shim (py2.7 has no urllib.request / urllib.parse).
try:                                        # python 3
    from urllib.request import Request as _UrlRequest, urlopen as _urlopen
    from urllib.parse import urlsplit as _urlsplit
except ImportError:                         # python 2.7
    from urllib2 import Request as _UrlRequest, urlopen as _urlopen
    from urlparse import urlsplit as _urlsplit

if hasattr(sys.stdout, "reconfigure"):
    # py3.7+ only; py2.7 has no reconfigure (and does not need one here).
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass

ALGO = "vb6src/2"   # spec sec 11.3.3: version VALUES no longer enter the hash
MODES = ("strict", "semantic")
DEFAULT_VBP = "MCP_VB6.vbp"
STAMP_BASENAME = "BuildStamp.bas"
# The .vbp key the stamp is written to.  It is VersionComments, NOT Comments: VB6 treats
# a bare "Comments" key as an invalid keyword and refuses to compile ("文件 ... 包含无效
# 的关键字 'Comments'", /make exit 1).  The version-resource FIELD that this becomes is
# still named "Comments", which is what the offline read point queries.
VBP_STAMP_KEY = "VersionComments"
# An earlier revision of this tool wrote this invalid key; it must never be hashed, and
# patch_vbp_comments migrates it away so the project stays compilable.
VBP_LEGACY_STAMP_KEY = "Comments"
EXCLUDED_BASENAMES = ("buildstamp.bas",)
COMPANION_EXTS = (".frx", ".res", ".ctx", ".dsr")
VBP_COMPONENT_KEYS = (
    "Form", "Class", "Module", "UserControl", "PropertyPage",
    "UserDocument", "Designer", "RelatedDoc", "ResFile32",
)
# semantic-mode normalisation applies to text project files only, never to the
# binary companions (.frx/.res/.ctx), whose bytes are always hashed raw.
SEMANTIC_TEXT_EXTS = (".vbp", ".bas", ".cls", ".frm", ".ctl", ".dsr", ".pag", ".dca")
NON_SOURCE_SUFFIXES = (".bak", ".obj", ".dll", ".exe", ".log", ".pdb", ".lib", ".exp")
NON_SOURCE_PREFIXES = ("docs/",)
# The input set is driven by the .vbp component list plus its sidecars, so an
# arbitrary work or output directory cannot contribute to the stamp on its own;
# extra prefixes can be added with VB6MCP_STAMP_EXCLUDE (comma separated).
NON_SOURCE_PREFIXES = tuple(NON_SOURCE_PREFIXES) + tuple(
    p.strip().replace("\\", "/") + "/"
    for p in (os.environ.get("VB6MCP_STAMP_EXCLUDE") or "").split(",")
    if p.strip()
)

RESERVED_PORT = 8766
MIN_SELF_OWNED_PORT = 8767


# --------------------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------------------
def eprint(*a):
    sys.stderr.write(" ".join(str(x) for x in a) + "\n")


def sha256_hex(b):
    return hashlib.sha256(b).hexdigest()


def read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def stat_bytes(path):
    """(mtime, size) or (None, None) when the path does not exist."""
    try:
        st = os.stat(path)
        return st.st_mtime, st.st_size
    except OSError:
        return None, None


def classify_write(path, data):
    """What a write of these bytes to this path WOULD do - touching nothing.

    Returns (state, old_sha256_or_None, new_sha256, mtime_before).  state is
    "created" (path absent) | "rewritten" (different bytes) | "byte-identical"
    (the file already carries exactly these bytes; a write would touch neither the
    bytes nor the mtime)."""
    new_sha = sha256_hex(data)
    mt_before, _sz = stat_bytes(path)
    if os.path.exists(path):
        try:
            old = read_bytes(path)
        except OSError:
            return "rewritten", None, new_sha, mt_before
        if old == data:
            return "byte-identical", sha256_hex(old), new_sha, mt_before
        return "rewritten", sha256_hex(old), new_sha, mt_before
    return "created", None, new_sha, mt_before


def plan_write(path, data):
    """The receipt a write WOULD produce, without performing it (--no-write)."""
    state, old_sha, new_sha, mt = classify_write(path, data)
    _mt, sz = stat_bytes(path)
    return dict(path=path, state=state, old_sha256=old_sha, new_sha256=new_sha,
                bytes=len(data), touched=False, planned=True, mtime_before=mt,
                mtime_after=mt, size_before=sz, size_after=sz)


def write_with_evidence(path, data):
    """Write these bytes and return the EVIDENCE of what happened to the file.

    {"path", "state", "old_sha256", "new_sha256", "bytes", "touched", "planned",
     "mtime_before", "mtime_after", "size_before", "size_after"}

    state: "created" | "rewritten" | "byte-identical" (see classify_write).

    The old/new sha256 travel WITH the state on purpose: a reader must be able to
CHECK the claim instead of trusting the word. Before the tool printed
    "(unchanged)" for a manifest whose content had in fact been rewritten, and the
    receipt carried no field that could have shown it."""
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    state, old_sha, new_sha, mt_before = classify_write(path, data)
    _mt, sz_before = stat_bytes(path)
    if state == "byte-identical":
        return dict(path=path, state=state, old_sha256=old_sha, new_sha256=new_sha,
                    bytes=len(data), touched=False, planned=False, mtime_before=mt_before,
                    mtime_after=mt_before, size_before=sz_before, size_after=sz_before)
    with open(path, "wb") as f:
        f.write(data)
    mt_after, sz_after = stat_bytes(path)
    return dict(path=path, state=state, old_sha256=old_sha, new_sha256=new_sha,
                bytes=len(data), touched=True, planned=False, mtime_before=mt_before,
                mtime_after=mt_after, size_before=sz_before, size_after=sz_after)


def write_if_changed(path, data):
    """Back-compat shim: the STATE of write_with_evidence."""
    return write_with_evidence(path, data)["state"]


def format_write_evidence(label, ev):
    """One truthful line per artifact: WHERE, WHAT HAPPENED, and the shas that prove it.

    The state is never inferred from "the call did not raise": it is the outcome of a
    byte comparison.  "byte-identical" is only reported after the on-disk bytes were
    compared with the payload, and it says the file was NOT touched."""
    old = ev["old_sha256"] if ev.get("old_sha256") is not None else "none"
    if ev.get("planned"):
        state_text = "NOT WRITTEN (--no-write); would be: %s" % ev["state"]
    elif ev["state"] == "byte-identical":
        state_text = "byte-identical: file NOT touched, 0 bytes written"
    else:
        state_text = ev["state"]
    return ("%s=%s (%s) old_sha256=%s new_sha256=%s"
            % (label, ev["path"], state_text, old, ev["new_sha256"]))


def is_excluded_rel(rel):
    low = rel.replace("\\", "/").lower()
    if os.path.basename(low) in EXCLUDED_BASENAMES:
        return True
    if low.startswith(NON_SOURCE_PREFIXES):
        return True
    for suf in NON_SOURCE_SUFFIXES:
        if low.endswith(suf):
            return True
    if ".bak" in low:
        return True
    return False


# --------------------------------------------------------------------------------------
# .vbp parsing and normalisation
# --------------------------------------------------------------------------------------
def _stamp_key_of(line):
    """"VersionComments" / legacy "Comments" when the line sets such a key, else None."""
    probe = line.lstrip(b" \t").lower()
    for key in (VBP_STAMP_KEY, VBP_LEGACY_STAMP_KEY):
        pfx = (key.lower() + "=").encode("ascii")
        if probe[:len(pfx)] == pfx:
            return key
    return None


def vbp_strip_comments(raw):
    """Remove every whole stamp-key line (including its line terminator).

    This is where the source stamp is written back, so it must not be part of the hashed
    .vbp content -- otherwise writing the stamp would change I and the whole scheme would
    be self-referential.  Both the sanctioned VersionComments key and the legacy invalid
    Comments key are stripped, so migrating between them cannot move I either.  Every
    other byte is preserved exactly."""
    out = []
    for line in raw.splitlines(True):
        if _stamp_key_of(line) is None:
            out.append(line)
    return b"".join(out)


#: The four .vbp version keys whose VALUES are build-dependent.  With AutoIncrementVer=1 a
#: /make WRITES a new RevisionVer back into the .vbp (spec sec 10.2), and the .vbp is part of
#: the input set - so without this mask a zero-code-change rebuild changes I.  The KEYS stay in
#: the hashed bytes; only the values are replaced, so every OTHER edit still moves I.
VBP_VERSION_KEYS = ("MajorVer", "MinorVer", "RevisionVer", "AutoIncrementVer")
#: the same keys as BYTES: the .vbp is hashed as raw bytes, so the comparison must be bytes.
#: (Comparing b"MajorVer" against the str tuple is silently always-False - the first version of
#: this mask did exactly that and the selftest caught it: case 6 failed.)
VBP_VERSION_KEYS_B = tuple(k.encode("ascii") for k in VBP_VERSION_KEYS)
VBP_VERSION_PLACEHOLDER = b"<version-independent>"


def vbp_mask_version_values(raw):
    """Replace the VALUE of each version key with a fixed placeholder (spec sec 11.3.2).

    Line endings are preserved byte-for-byte, and a key that is absent stays absent (masking
    is not insertion).  Applied to the .vbp only, before hashing.
    """
    out = []
    for line in raw.split(b"\n"):
        body = line[:-1] if line.endswith(b"\r") else line
        eol = b"\r" if line.endswith(b"\r") else b""
        key, sep, _val = body.partition(b"=")
        if sep and key.strip() in VBP_VERSION_KEYS_B:
            out.append(key + sep + VBP_VERSION_PLACEHOLDER + eol)
        else:
            out.append(line)
    return b"\n".join(out)


def parse_vbp_components(raw):
    """[(key, relative file name)] for every component line in the .vbp."""
    comps = []
    seen = set()
    for line in raw.decode("latin1").splitlines():
        s = line.strip()
        if not s or s[0] in ";'":
            continue
        if "=" not in s:
            continue
        key, val = s.split("=", 1)
        key = key.strip()
        if key not in VBP_COMPONENT_KEYS:
            continue
        val = val.strip()
        fname = val.split(";")[-1].strip() if ";" in val else val
        fname = fname.strip().strip('"').replace("\\", "/")
        if not fname:
            continue
        if fname.lower() in seen:
            continue
        seen.add(fname.lower())
        comps.append((key, fname))
    return comps


def normalize_semantic(raw, ext):
    if ext.lower() not in SEMANTIC_TEXT_EXTS:
        return raw
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    raw = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    raw = raw.rstrip(b"\n")
    if raw:
        raw += b"\n"
    return raw


# --------------------------------------------------------------------------------------
# input set + indicator
# --------------------------------------------------------------------------------------
def canonical_rel(tree, vbp_rel):
    r"""(manifest_rel, abspath) - the .vbp path as it must appear IN THE MANIFEST.

    The manifest records a path STRING per input file and that string is hashed, so recording
    the caller's spelling made I depend on HOW the tool was invoked: --vbp MCP_VB6.vbp (CWD =
    tree root) and --vbp D:\...\MCP_VB6.vbp produced different I for byte-identical trees
    (spec sec 10.8).  Every entry is therefore recorded tree-root-relative, with
    forward slashes, so the two invocations (and any CWD) agree.

    Only the path FORM is canonicalised - the tree's own location is never part of the entry,
    and no file name is case-folded (a genuinely different name must stay a different input).
    """
    tree_abs = os.path.abspath(tree)
    vbp_abs = os.path.abspath(os.path.join(tree_abs, vbp_rel.replace("/", os.sep)))
    return os.path.relpath(vbp_abs, tree_abs).replace(os.sep, "/"), vbp_abs


def collect_input_tree(tree, vbp_rel, mode):
    """(entries, problems).  entries = [{rel, abspath, raw, hashed}] (unsorted)."""
    entries = []
    problems = []
    rel_vbp, vbp_abs = canonical_rel(tree, vbp_rel)
    if not os.path.isfile(vbp_abs):
        problems.append("project file not found: %s" % vbp_rel)
        return entries, problems
    vbp_raw = read_bytes(vbp_abs)
    stripped = vbp_strip_comments(vbp_raw)
    stripped = vbp_mask_version_values(stripped)   # sec 11.3.2 - version values never move I
    hashed = stripped if mode == "strict" else normalize_semantic(stripped, ".vbp")
    entries.append(dict(rel=rel_vbp, abspath=vbp_abs, raw=vbp_raw, hashed=hashed))

    components = parse_vbp_components(vbp_raw)
    # Two different things, deliberately kept apart:
    #   required   -- files the .vbp explicitly lists.  Missing => the tree cannot be
    #                 built, so this is a real exhaustiveness failure and we refuse;
    #   companions -- .frx/.res/.ctx/.dsr sidecars.  The spec says "counted only when
    #                 present": a VB6 project normally has a .frx or two and NO
    #                 .ctx/.dsr at all, so a missing companion is simply "not an input",
    #                 never an error.
    required = []
    companions = []
    for _key, fname in components:
        if is_excluded_rel(fname):
            # e.g. BuildStamp.bas: the stamp's own output, excluded BEFORE the
            # existence check so it may be absent or present without changing I.
            continue
        if fname not in required:
            required.append(fname)
        stem, ext = os.path.splitext(fname)
        for cext in COMPANION_EXTS:
            if cext == ext:
                continue
            cand = stem + cext
            if is_excluded_rel(cand) or cand in companions or cand in required:
                continue
            if os.path.isfile(os.path.join(tree, cand.replace("/", os.sep))):
                companions.append(cand)
    for rel in required + companions:
        ap = os.path.join(tree, rel.replace("/", os.sep))
        if not os.path.isfile(ap):
            problems.append("component file listed by the .vbp is missing: %s" % rel)
            continue
        raw = read_bytes(ap)
        ext = os.path.splitext(rel)[1]
        hashed = raw if mode == "strict" else normalize_semantic(raw, ext)
        entries.append(dict(rel=rel, abspath=ap, raw=raw, hashed=hashed))
    return entries, problems


def manifest_payload(entries, mode):
    lines = []
    for e in sorted(entries, key=lambda x: x["rel"].encode("utf-8")):
        lines.append("%s\t%d\t%s" % (e["rel"], len(e["hashed"]), sha256_hex(e["hashed"])))
    text = ALGO + "\n" + mode + "\n" + "\n".join(lines) + "\n"
    return text.encode("utf-8")


def manifest_index(payload):
    """{rel: (size, sha)} from a manifest payload."""
    idx = {}
    text = payload.decode("utf-8", "replace")
    for line in text.split("\n"):
        parts = line.split("\t")
        if len(parts) == 3 and len(parts[2]) == 64:
            try:
                idx[parts[0]] = (int(parts[1]), parts[2])
            except ValueError:
                continue
    return idx


def compute_stamp(tree, vbp_rel, mode):
    entries, problems = collect_input_tree(tree, vbp_rel, mode)
    if problems:
        return None, None, None, problems
    payload = manifest_payload(entries, mode)
    return sha256_hex(payload), payload, entries, []


# --------------------------------------------------------------------------------------
# PE / resource / version-resource reading (no string scanning)
# --------------------------------------------------------------------------------------
def _u16(b, o):
    return struct.unpack_from("<H", b, o)[0]


def _u32(b, o):
    return struct.unpack_from("<I", b, o)[0]


def parse_pe(b):
    if len(b) < 0x40 or _u16(b, 0) != 0x5A4D:
        raise ValueError("not an MZ image")
    e_lfanew = _u32(b, 0x3C)
    if b[e_lfanew:e_lfanew + 4] != b"PE\x00\x00":
        raise ValueError("no PE signature at e_lfanew=%d" % e_lfanew)
    coff = e_lfanew + 4
    machine = _u16(b, coff)
    nsec = _u16(b, coff + 2)
    timestamp = _u32(b, coff + 4)
    size_opt = _u16(b, coff + 16)
    opt = coff + 20
    magic = _u16(b, opt)
    if magic == 0x10B:
        ndirs_off, dirs_off = opt + 92, opt + 96
    elif magic == 0x20B:
        ndirs_off, dirs_off = opt + 108, opt + 112
    else:
        raise ValueError("unknown optional-header magic 0x%04X" % magic)
    ndirs = _u32(b, ndirs_off)
    dirs = [(i, _u32(b, dirs_off + i * 8), _u32(b, dirs_off + i * 8 + 4))
            for i in range(min(ndirs, 16))]
    sec = coff + 20 + size_opt
    sections = []
    for i in range(nsec):
        o = sec + i * 40
        sections.append(dict(name=b[o:o + 8].rstrip(b"\x00").decode("latin1"),
                             vsize=_u32(b, o + 8), va=_u32(b, o + 12),
                             rawsize=_u32(b, o + 16), raw=_u32(b, o + 20),
                             sec_off=o))
    return dict(e_lfanew=e_lfanew, machine=machine, nsec=nsec, timestamp=timestamp,
                magic=magic, dirs=dirs, sections=sections)


def rva_to_off(pe, rva):
    for s in pe["sections"]:
        if s["va"] <= rva < s["va"] + max(s["vsize"], s["rawsize"]):
            return s["raw"] + (rva - s["va"])
    return None


def walk_resources(b, pe, res_rva):
    """[(path, data_rva, data_size, codepage, data_entry_off)]"""
    base = rva_to_off(pe, res_rva)
    if base is None:
        return []
    found = []

    def walk(doff, path):
        n_named, n_id = _u16(b, doff + 12), _u16(b, doff + 14)
        for i in range(n_named + n_id):
            e = doff + 16 + i * 8
            nm, off = _u32(b, e), _u32(b, e + 4)
            if nm & 0x80000000:
                no = base + (nm & 0x7FFFFFFF)
                ln = _u16(b, no)
                label = b[no + 2:no + 2 + ln * 2].decode("utf-16-le", "replace")
            else:
                label = nm
            if off & 0x80000000:
                walk(base + (off & 0x7FFFFFFF), path + [label])
            else:
                de = base + off
                found.append((path + [label], _u32(b, de), _u32(b, de + 4),
                              _u32(b, de + 8), de))

    walk(base, [])
    return found


def _strip_utf16_nuls(b):
    """Strip trailing NUL *WCHARs* (two bytes at a time).

    Byte-wise trimming of 0x00 bytes is WRONG here: it also eats the high byte of the
    last character (b"P" + NUL + NUL + NUL -> b"P"), leaving an odd-length buffer.
    VB6 terminates a string value with one NUL WCHAR and then pads to 4 bytes."""
    while len(b) >= 2 and b[-2:] == b"\x00\x00":
        b = b[:-2]
    if len(b) % 2:
        b = b[:-1]
    return b


def _versionize_at(blob, start):
    """Parse the VS_VERSIONINFO node starting at blob[start].

    Returns (key, value_bytes, child_off, node_len) where child_off is an ABSOLUTE
    offset into blob (or None when the node carries text and therefore no children).
    Keeping every offset absolute avoids the classic slice-relative/absolute mix-up
    when walking nested node levels.

    Two traps handled here, both observed on a real VB6 DLL:
      * VS_VERSION_INFO carries BOTH a fixed value (the 52-byte VS_FIXEDFILEINFO) and
        children (StringFileInfo / VarFileInfo); a reader that returns "it has a value,
        so it has no children" never finds a single string;
      * VB6 writes wValueLength in BYTES for string nodes while the Win32
        documentation describes characters, so a text value is taken as the bytes from
        the value start to the end of THIS node (bounded by wLength) instead of
        trusting wValueLength."""
    if len(blob) - start < 6:
        raise ValueError("truncated version node at %d" % start)
    node_len = _u16(blob, start)
    value_len = _u16(blob, start + 2)
    wtype = _u16(blob, start + 4)
    if node_len == 0:
        node_len = len(blob) - start
    p = start + 6
    while p + 1 < len(blob) and _u16(blob, p) != 0:
        p += 2
    key = blob[start + 6:p].decode("utf-16-le", "replace")
    after = (p + 2 + 3) & ~3
    end = min(start + node_len, len(blob))
    # wType alone does NOT say whether a node has children: on a real VB6 DLL the
    # containers StringFileInfo and StringTable are written with wType=1 but
    # wValueLength=0.  The reliable discriminator is wValueLength:
    #   value_len>0 and wType==1 -> a text leaf (no children)
    #   value_len>0 and wType==0 -> a binary value FOLLOWED BY children (the root)
    #   value_len==0              -> a container; its children start right after the key
    if value_len > 0 and wtype == 1:
        return key, _strip_utf16_nuls(blob[after:end]), None, node_len
    if value_len > 0:
        return key, blob[after:after + value_len], (after + value_len + 3) & ~3, node_len
    return key, b"", after, node_len


def read_version_strings_python(b):
    """(dict(strings=..., string_table=...), error) from RT_VERSION, pure python."""
    pe = parse_pe(b)
    res_rva = 0
    for i, rva, _sz in pe["dirs"]:
        if i == 2:
            res_rva = rva
    if not res_rva:
        return None, "no resource directory"
    blob = None
    for path, drva, dsz, _cp, _de in walk_resources(b, pe, res_rva):
        if path and path[0] == 16:
            off = rva_to_off(pe, drva)
            if off is not None:
                blob = b[off:off + dsz]
            break
    if blob is None:
        return None, "no RT_VERSION resource"
    try:
        key, _v, child, node_len = _versionize_at(blob, 0)
        if key != "VS_VERSION_INFO":
            return None, "unexpected root version key %r" % key
        strings = {}
        table_key = None
        end = min(node_len, len(blob))
        c = child
        while c is not None and c < end:
            ck, _cv, cchild, clen = _versionize_at(blob, c)
            if clen == 0:
                break
            if ck == "StringFileInfo":
                sub = cchild
                sub_end = min(c + clen, end)
                while sub is not None and sub < sub_end:
                    sk, _sv, schild, slen = _versionize_at(blob, sub)
                    if slen == 0:
                        break
                    if table_key is None:
                        table_key = sk
                    q = schild
                    q_end = min(sub + slen, sub_end)
                    while q is not None and q < q_end:
                        vk, vv, _vchild, vlen = _versionize_at(blob, q)
                        if vlen == 0:
                            break
                        if vk:
                            strings[vk] = vv.decode("utf-16-le", "replace")
                        q = (q + vlen + 3) & ~3
                    sub = (sub + slen + 3) & ~3
            c = (c + clen + 3) & ~3
        return dict(strings=strings, string_table=table_key), None
    except Exception as exc:  # noqa: BLE001 - reported, never guessed
        return None, "version resource parse failed: %r" % (exc,)


def read_version_strings_winapi(path, names, table_key):
    """Read the same strings through the documented API (GetFileVersionInfoW +
    VerQueryValueW).  Returns (dict, error).  Windows only."""
    if os.name != "nt":
        return None, "GetFileVersionInfo is Windows-only (running on %s)" % os.name
    if not table_key:
        return None, "no StringFileInfo language/codepage key available to query"
    try:
        import ctypes
        import ctypes.wintypes as wt
        ver = ctypes.WinDLL("version")
        ver.GetFileVersionInfoSizeW.argtypes = [wt.LPCWSTR, ctypes.POINTER(wt.DWORD)]
        ver.GetFileVersionInfoSizeW.restype = wt.DWORD
        ver.GetFileVersionInfoW.argtypes = [wt.LPCWSTR, wt.DWORD, wt.DWORD,
                                            ctypes.c_void_p]
        ver.GetFileVersionInfoW.restype = wt.BOOL
        ver.VerQueryValueW.argtypes = [ctypes.c_void_p, wt.LPCWSTR,
                                       ctypes.POINTER(ctypes.c_void_p),
                                       ctypes.POINTER(wt.UINT)]
        ver.VerQueryValueW.restype = wt.BOOL
        fn = os.path.abspath(path)
        size = ver.GetFileVersionInfoSizeW(fn, None)
        if not size:
            return None, "GetFileVersionInfoSizeW returned 0 (no version resource)"
        buf = ctypes.create_string_buffer(size)
        if not ver.GetFileVersionInfoW(fn, 0, size, buf):
            return None, "GetFileVersionInfoW failed"
        out = {}
        for name in names:
            sub = "\\StringFileInfo\\" + table_key + "\\" + name
            p = ctypes.c_void_p()
            ln = wt.UINT()
            if ver.VerQueryValueW(buf, sub, ctypes.byref(p), ctypes.byref(ln)) and p.value:
                out[name] = ctypes.wstring_at(p.value)
        return out, None
    except Exception as exc:  # noqa: BLE001
        return None, "winapi read failed: %r" % (exc,)


def parse_stamp_text(text):
    """'algo=vb6src/2;src=<hex>' -> {'algo':..., 'src':...} (missing keys absent)."""
    out = {}
    if not text:
        return out
    for k, v in re.findall(r"([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^;\s]+)", text):
        out[k.lower()] = v
    return out


def read_stamp_from_artifact(path):
    """Read the stamp out of the artifact's version resource, without string scanning.

    Returns dict(ok, stamp, comments, readpoints, notes, error)."""
    res = dict(ok=False, stamp={}, comments=None, readpoints={}, notes=[], error=None)
    try:
        b = read_bytes(path)
    except OSError as exc:
        res["error"] = "cannot read artifact: %r" % (exc,)
        return res
    py, py_err = None, None
    try:
        py, py_err = read_version_strings_python(b)
    except Exception as exc:  # noqa: BLE001
        py_err = "PE/version parse failed: %r" % (exc,)
    if py_err:
        res["notes"].append("pure-python version read: " + py_err)
    win, win_err = None, None
    if py and py["strings"]:
        win, win_err = read_version_strings_winapi(path, sorted(py["strings"]),
                                                   py["string_table"])
        if win_err:
            res["notes"].append("GetFileVersionInfo read: " + win_err)
    if win:
        res["readpoints"]["winapi"] = win
    if py and py["strings"]:
        res["readpoints"]["python"] = py["strings"]
    if not res["readpoints"]:
        res["error"] = "no version-resource strings could be read"
        return res
    comments = None
    # Cross-check the two readers on every string they both saw; a disagreement means
    # the reader is unreliable, so it is an error rather than something to smooth over.
    if win and py and py["strings"]:
        for k in sorted(set(win) & set(py["strings"])):
            if win[k] != py["strings"][k]:
                res["error"] = ("the two version-resource readers disagree on %s: "
                                "GetFileVersionInfo=%r python=%r"
                                % (k, win[k], py["strings"][k]))
                return res
    for channel in ("winapi", "python"):
        got = res["readpoints"].get(channel) or {}
        if "Comments" in got:
            comments = got["Comments"]
            res["comments_channel"] = channel
            break
    res["comments"] = comments
    res["stamp"] = parse_stamp_text(comments.strip() if comments else None)
    res["ok"] = True
    return res


# --------------------------------------------------------------------------------------
# mcp_info read point (structured, no string scanning)
# --------------------------------------------------------------------------------------
def _http_post(url, body, headers, timeout=60):
    req = _UrlRequest(url, data=json.dumps(body).encode("utf-8"), headers=headers)
    with _urlopen(req, timeout=timeout) as r:
        raw = r.read().decode("utf-8-sig")
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("data:"):
            line = line[5:].strip()
        if line.startswith("{"):
            return json.loads(line)
    return json.loads(raw) if raw.strip() else None


def _find_key(obj, wanted):
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == wanted:
                return v
        for v in obj.values():
            got = _find_key(v, wanted)
            if got is not None:
                return got
    elif isinstance(obj, list):
        for v in obj:
            got = _find_key(v, wanted)
            if got is not None:
                return got
    return None


def read_stamp_from_mcp(url):
    """Call mcp_info over MCP and pull build.sourceHash / build.sourceAlgo out of the
    structured result.  Port 8766 is reserved and is refused without sending anything."""
    res = dict(ok=False, stamp={}, error=None)
    parts = _urlsplit(url)
    port = parts.port or (443 if parts.scheme == "https" else 80)
    if port == RESERVED_PORT:
        res["error"] = ("REFUSED: port %d is reserved; nothing was sent" % RESERVED_PORT)
        res["refused"] = True
        return res
    if port < MIN_SELF_OWNED_PORT:
        res["error"] = ("REFUSED: port %d is below the self-owned range (>= %d); "
                        "nothing was sent" % (port, MIN_SELF_OWNED_PORT))
        res["refused"] = True
        return res
    try:
        init = _http_post(url, {"jsonrpc": "2.0", "id": 1, "method": "initialize",
                                "params": {"protocolVersion": "2025-06-18",
                                           "capabilities": {},
                                           "clientInfo": {"name": "srcstamp",
                                                          "version": "1"}}},
                          {"Content-Type": "application/json",
                           "Accept": "application/json, text/event-stream"})
        proto = (init.get("result") or {}).get("protocolVersion") or "2025-06-18"
        call = _http_post(url, {"jsonrpc": "2.0", "id": 2, "method": "tools/call",
                                "params": {"name": "mcp_info", "arguments": {}}},
                          {"Content-Type": "application/json",
                           "Accept": "application/json, text/event-stream",
                           "MCP-Protocol-Version": proto})
    except Exception as exc:  # noqa: BLE001
        res["error"] = "mcp_info call failed: %r" % (exc,)
        return res
    result = call.get("result") if isinstance(call, dict) else None
    payload = None
    if isinstance(result, dict):
        if result.get("structuredContent") is not None:
            payload = result["structuredContent"]
        for item in (result.get("content") or []):
            if isinstance(item, dict) and item.get("type") == "text":
                try:
                    payload = json.loads(item.get("text") or "")
                    break
                except Exception:  # noqa: BLE001
                    continue
    if payload is None:
        res["error"] = "mcp_info returned no structured payload"
        return res
    src = _find_key(payload, "sourceHash")
    algo = _find_key(payload, "sourceAlgo")
    if src is None:
        res["error"] = "mcp_info result has no build.sourceHash (the add-in reading the "
        res["error"] += " constants is the VB side of this batch)"
        return res
    res["stamp"] = {"src": str(src), "algo": str(algo) if algo is not None else None}
    res["ok"] = True
    return res


# --------------------------------------------------------------------------------------
# .vbp Comments= write-back
# --------------------------------------------------------------------------------------
def patch_vbp_comments(vbp_path, value):
    """Idempotently set the .vbp VersionComments= line.  Only that line changes.

    A legacy invalid "Comments=" line is migrated away at the same time (VB6 refuses to
    compile a .vbp containing it).  Returns (state, new_bytes)."""
    raw = read_bytes(vbp_path)
    newline = b"\r\n" if b"\r\n" in raw else b"\n"
    replacement = (VBP_STAMP_KEY + "=" + value).encode("ascii", "replace")
    lines = raw.splitlines(True)
    out, replaced = [], False
    for line in lines:
        if _stamp_key_of(line) is not None:
            if not replaced:
                if line.endswith(b"\r\n"):
                    term = b"\r\n"
                elif line.endswith(b"\n"):
                    term = b"\n"
                else:
                    term = newline
                out.append(replacement + term)
                replaced = True
            continue
        out.append(line)
    if not replaced:
        insert_at = len(out)
        for i, line in enumerate(out):
            if line.lstrip(b" \t")[:1] == b"[":
                insert_at = i
                break
        out.insert(insert_at, replacement + newline)
    new_raw = b"".join(out)
    ev = write_with_evidence(vbp_path, new_raw)
    return ev["state"], new_raw


def render_buildstamp(algo, stamp_hash, mode):
    lines = [
        'Attribute VB_Name = "BuildStamp"',
        "Option Explicit",
        "",
        "' GENERATED by tools/srcstamp.py - DO NOT EDIT. algo=%s mode=%s" % (algo, mode),
        'Public Const VB6MCP_SOURCE_ALGO As String = "%s"' % algo,
        'Public Const VB6MCP_SOURCE_HASH As String = "%s"' % stamp_hash,
    ]
    # VB project files are ANSI without BOM, CRLF terminated.
    return ("\r\n".join(lines) + "\r\n").encode("latin1")


def comments_value(algo, stamp_hash):
    return "algo=%s;src=%s" % (algo, stamp_hash)


# --------------------------------------------------------------------------------------
# version as a SEPARATE comparison item (spec sec 11.3.4)
#
# The stamp answers "is the SOURCE the same".  A version number is a different dimension:
# AutoIncrementVer=1 makes every /make produce a different FileVersion from identical source,
# so a version difference must be LISTED and must NEVER move the verdict.
# --------------------------------------------------------------------------------------
def vbp_version_values(tree, vbp_rel):
    """{key: int|None} for the four version keys, read straight from the .vbp."""
    out = {}
    try:
        raw = read_bytes(os.path.join(tree, vbp_rel.replace("/", os.sep)))
    except OSError:
        return out
    for line in raw.split(b"\n"):
        key, sep, val = line.partition(b"=")
        k = key.strip().decode("latin1", "replace")
        if sep and k in VBP_VERSION_KEYS:
            v = val.strip().decode("latin1", "replace")
            try:
                out[k] = int(v)
            except ValueError:
                out[k] = None
    return out


def artifact_versions(art):
    """(fileVersion, productVersion, channel) from the version-resource read points."""
    for ch in ("winapi", "python"):
        got = ((art.get("readpoints") or {}).get(ch)) or {}
        if got.get("FileVersion") or got.get("ProductVersion"):
            return got.get("FileVersion"), got.get("ProductVersion"), ch
    return None, None, None


def version_tuple(text):
    """'1.10.0092' -> (1, 10, 92).  None when nothing numeric is present."""
    if not text:
        return None
    parts = []
    for chunk in str(text).replace(",", ".").split("."):
        digits = "".join(c for c in chunk if c.isdigit())
        if digits:
            parts.append(int(digits))
    return tuple(parts) if parts else None


def version_record_path(mdir, stamp_hash):
    return os.path.join(mdir, "srcstamp", (stamp_hash or "none") + ".srcversion.txt")


def read_version_record(path):
    """{k: v} or None.  This is a RECORD, never an input: it is not hashed into I."""
    if not path or not os.path.isfile(path):
        return None
    out = {}
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            for line in fh:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    out[k] = v
    except OSError:
        return None
    return out or None


def write_version_record(path, algo, stamp, artifact, fv, pv):
    d = os.path.dirname(path)
    if d and not os.path.isdir(d):
        os.makedirs(d)
    body = ("algo=%s\nsrc=%s\nartifact=%s\nfileVersion=%s\nproductVersion=%s\n"
            % (algo, stamp, os.path.abspath(artifact or ""), fv or "", pv or ""))
    return write_if_changed(path, body.encode("utf-8"))


def report_versions(tree, vbp_rel, art, report, mdir, artifact):
    """Print the version dimension and compare it.  NEVER affects the verdict."""
    fv, pv, ch = artifact_versions(art)
    tv = vbp_version_values(tree, vbp_rel)
    print("ARTIFACT_VERSION FileVersion=%s ProductVersion=%s (channel=%s)"
          % (fv or "none", pv or "none", ch or "none"))
    print("VBP_VERSION MajorVer=%s MinorVer=%s RevisionVer=%s AutoIncrementVer=%s"
          % (tv.get("MajorVer"), tv.get("MinorVer"), tv.get("RevisionVer"),
             tv.get("AutoIncrementVer")))
    ft = version_tuple(fv)
    tt = (tv.get("MajorVer"), tv.get("MinorVer"), tv.get("RevisionVer"))
    state = "unknown"
    detail = "one side is unreadable; reported as unknown rather than guessed"
    if ft and all(x is not None for x in tt):
        # "1.00" and "1.0.0" are the same version: pad to a common length before comparing,
        # otherwise a missing trailing component reads as a false disagreement.
        n = max(len(ft), len(tt))
        fp = tuple(ft) + (0,) * (n - len(ft))
        tp = tuple(tt) + (0,) * (n - len(tt))
        if fp == tp:
            state, detail = "yes", "the artifact's FileVersion equals the .vbp version fields"
        elif fp[:-1] == tp[:-1] and fp[-1] + 1 == tp[-1]:
            state = "auto-increment"
            detail = ("expected with AutoIncrementVer=1: /make stamps the artifact with the "
                      "PRE-increment revision and then writes the NEW revision back into the "
                      ".vbp (spec sec 10.2/11.2)")
        else:
            state, detail = "no", "artifact and .vbp disagree on the version"
    print("VERSION_COMPARE FileVersion artifact=%s vbp=%s match=%s (%s)"
          % (fv or "none", ".".join(str(x) for x in tt) if all(x is not None for x in tt)
             else "unknown", state, detail))
    print("VERSION_NOTE version is a SEPARATE dimension: a difference here never changes "
          "VERDICT (spec sec 11.3.4)")
    state_d = dict(artifactFileVersion=fv, artifactProductVersion=pv, channel=ch,
                   vbp=dict(tv), match=state, detail=detail, recordPath=None,
                   recorded=None, recordedChanged=None)
    report["version"] = state_d
    return state_d


def compare_version_record(args, stamp_hash, state_d, write=False):
    """List this artifact's version against the record for the SAME stamp.

    The record is keyed by the stamp, so two builds of one source (identical I) land on the
    same record - which is what lets a rebuild show 'same source, different version, PASS'."""
    fv = state_d.get("artifactFileVersion")
    pv = state_d.get("artifactProductVersion")
    path = version_record_path(manifest_dir_for(args), stamp_hash)
    state_d["recordPath"] = path
    rec = read_version_record(path)
    if rec:
        state_d["recorded"] = dict(rec)
        fv_changed = rec.get("fileVersion", "") != (fv or "")
        pv_changed = rec.get("productVersion", "") != (pv or "")
        state_d["recordedChanged"] = bool(fv_changed or pv_changed)
        print("VERSION_RECORDED FileVersion=%s ProductVersion=%s (from %s)"
              % (rec.get("fileVersion") or "none", rec.get("productVersion") or "none", path))
        print("VERSION_RECORDED_DELTA fileVersion=%s productVersion=%s "
              "(a difference here is EXPECTED when AutoIncrementVer=1 and is NOT a failure)"
              % ("changed" if fv_changed else "same", "changed" if pv_changed else "same"))
    else:
        print("VERSION_RECORDED none (%s)" % path)
    if write and (fv or pv):
        write_version_record(path, ALGO, stamp_hash, args.artifact, fv, pv)
        print("VERSION_RECORD_WRITTEN " + path)
    return state_d


def manifest_dir_for(args):
    return args.manifest_dir or os.path.join(os.getcwd(), "dist")


def manifest_targets(mdir, tree, stamp_hash):
    """The two manifest paths a stamp owns: the generation file and the named file."""
    store = os.path.join(mdir, "srcstamp")
    base = os.path.basename(os.path.abspath(tree)) or "tree"
    return [os.path.join(store, stamp_hash + ".srcmanifest.txt"),
            os.path.join(mdir, base + ".srcmanifest.txt")]


def record_manifest(mdir, tree, payload, stamp_hash):
    """Write both manifests and return their write EVIDENCE (not just a state word).

    This tool WROTE two files; the receipt now says which ones, whether each was
    created / rewritten / already byte-identical, and each file's own old->new sha256,
    so a caller can verify the claim instead of trusting it."""
    return [write_with_evidence(p, payload)
            for p in manifest_targets(mdir, tree, stamp_hash)]


def find_baseline(args, stamp_hash):
    """Locate a recorded manifest for the artifact's stamp so a per-file diff is
    possible.  Never guesses: returns (path_or_None, searched[])."""
    searched = []
    if args.baseline:
        searched.append(args.baseline)
        return (args.baseline if os.path.isfile(args.baseline) else None), searched
    mdir = manifest_dir_for(args)
    cands = []
    if args.artifact:
        stem = os.path.splitext(os.path.basename(args.artifact))[0]
        cands.append(os.path.join(mdir, "srcstamp", stem + ".srcmanifest.txt"))
    cands.append(os.path.join(mdir, "srcstamp", stamp_hash + ".srcmanifest.txt"))
    for c in cands:
        searched.append(c)
        if os.path.isfile(c):
            return c, searched
    return None, searched


def print_diff(old_payload, new_payload):
    old = manifest_index(old_payload)
    new = manifest_index(new_payload)
    added = sorted(set(new) - set(old))
    removed = sorted(set(old) - set(new))
    changed = sorted(k for k in set(old) & set(new) if old[k][1] != new[k][1])
    print("DIFF added=%d removed=%d changed=%d" % (len(added), len(removed), len(changed)))
    for k in added:
        print("  + %s  size=%d sha=%s" % (k, new[k][0], new[k][1][:16]))
    for k in removed:
        print("  - %s  size=%d sha=%s" % (k, old[k][0], old[k][1][:16]))
    for k in changed:
        print("  ~ %s  size %d->%d  sha %s -> %s"
              % (k, old[k][0], new[k][0], old[k][1][:16], new[k][1][:16]))
    if not (added or removed or changed):
        print("  (no per-file difference could be attributed; the recorded manifest "
              "belongs to a different input set)")
    return dict(added=added, removed=removed, changed=changed)


# --------------------------------------------------------------------------------------
# subcommands
# --------------------------------------------------------------------------------------
def cmd_compute(args):
    stamp, payload, entries, problems = compute_stamp(args.tree, args.vbp, args.mode)
    if problems:
        eprint("srcstamp: the input set is not exhaustive; refusing to report a stamp:")
        for p in problems:
            eprint("  - " + p)
        return 2
    if args.manifest:
        targets = [args.manifest]
    else:
        targets = manifest_targets(manifest_dir_for(args), args.tree, stamp)
    if args.no_write:
        # READ-ONLY form.  compute WRITES by default, and a check tool that calls that
# form silently modifies the working tree (: it minted a "generation"
        # manifest that no build had ever produced).  --no-write reports what WOULD
        # happen and writes nothing.
        recorded = [plan_write(p, payload) for p in targets]
        write_mode = "verified-only (nothing was written)"
    else:
        recorded = [write_with_evidence(p, payload) for p in targets]
        write_mode = "wrote %d file(s)" % len([e for e in recorded if e["touched"]])
    out = dict(algo=ALGO, mode=args.mode, source_hash=stamp,
               files=len(entries), payload_bytes=len(payload),
               manifest_sha256_matches=(sha256_hex(payload) == stamp),
               write_mode=write_mode,
               recorded=recorded,
               tree=os.path.abspath(args.tree), vbp=args.vbp)
    if args.json:
        print(json.dumps(out, indent=1, sort_keys=True))
    else:
        print("ALGO=" + ALGO)
        print("MODE=" + args.mode)
        print("SOURCE_HASH=" + stamp)
        print("FILES=%d" % len(entries))
        for e in recorded:
            print(format_write_evidence("MANIFEST", e))
        print("WRITE_MODE=" + write_mode)
    return 0


def cmd_emit(args):
    stamp, payload, entries, problems = compute_stamp(args.tree, args.vbp, args.mode)
    if problems:
        eprint("srcstamp: the input set is not exhaustive; refusing to emit:")
        for p in problems:
            eprint("  - " + p)
        return 2
    data = render_buildstamp(ALGO, stamp, args.mode)
    result = dict(algo=ALGO, mode=args.mode, source_hash=stamp, files=len(entries),
                  buildstamp=None, vbp=None, manifest=None,
                  tree=os.path.abspath(args.tree))
    if args.out:
        # the state is the OUTCOME of a byte comparison, reported together
        # with the file's own old->new sha256 - never as a bare word.
        result["buildstamp"] = write_with_evidence(args.out, data)
    vbp_value = comments_value(ALGO, stamp)
    # The .vbp Comments= line is the offline read point (spec section 5.2) and is
    # written BY THIS TOOL so that both write points provably come from one process and
    # one value.  It is therefore ON BY DEFAULT -- the plain
    #     srcstamp.py emit . --vbp MCP_VB6.vbp --out BuildStamp.bas
    # sets it.  --vbp-patch retargets it, --no-vbp-patch disables it (dry run).
    if args.vbp_patch:
        vbp_target = args.vbp_patch
    elif args.no_vbp_patch:
        vbp_target = None
    else:
        vbp_target = os.path.join(args.tree, args.vbp.replace("/", os.sep))
    if vbp_target:
        state, _ = patch_vbp_comments(vbp_target, vbp_value)
        result["vbp"] = dict(path=vbp_target, state=state, value=vbp_value)
    recorded = record_manifest(manifest_dir_for(args), args.tree, payload, stamp)
    result["manifest"] = recorded
    n_written = len([e for e in recorded if e["touched"]])
    if result["buildstamp"] and result["buildstamp"]["touched"]:
        n_written += 1
    result["write_mode"] = "wrote %d file(s)" % n_written
    if args.json:
        print(json.dumps(result, indent=1, sort_keys=True))
    else:
        print("ALGO=" + ALGO)
        print("MODE=" + args.mode)
        print("SOURCE_HASH=" + stamp)
        print("FILES=%d" % len(entries))
        if result["buildstamp"]:
            print(format_write_evidence("BUILDSTAMP", result["buildstamp"]))
        if result["vbp"]:
            print("VBP=%s (%s) %s=%s" % (result["vbp"]["path"], result["vbp"]["state"],
                                         VBP_STAMP_KEY, vbp_value))
        else:
            print("VBP_STAMP_LINE=%s=%s (NOT WRITTEN: --no-vbp-patch)"
                  % (VBP_STAMP_KEY, vbp_value))
        for m in result["manifest"]:
            print(format_write_evidence("MANIFEST", m))
        print("WRITE_MODE=" + result["write_mode"])
    return 0


def cmd_check(args):
    stamp, payload, entries, problems = compute_stamp(args.tree, args.vbp, args.mode)
    if problems:
        eprint("srcstamp: the input set is not exhaustive; cannot verify:")
        for p in problems:
            eprint("  - " + p)
        return 2
    report = dict(algo=ALGO, mode=args.mode, tree_hash=stamp, artifact=args.artifact,
                  read_points={}, failures=[], notes=[])
    if not args.json:
        print("SOURCE_TREE_HASH=" + stamp)
    mcp = None
    if args.mcp_url:
        mcp = read_stamp_from_mcp(args.mcp_url)
        report["read_points"]["mcp_info"] = dict(ok=mcp.get("ok"), stamp=mcp.get("stamp"),
                                                 error=mcp.get("error"))
        if mcp.get("error"):
            print("READPOINT mcp_info: " + str(mcp["error"]))
    art = read_stamp_from_artifact(args.artifact)
    report["read_points"]["version_resource"] = dict(
        ok=art.get("ok"), comments=art.get("comments"),
        channel=art.get("comments_channel"), error=art.get("error"),
        notes=art.get("notes"), strings=sorted((art.get("readpoints", {}).get("python")
                                                or {}).keys()))
    for n in art.get("notes") or []:
        report["notes"].append(n)
        print("NOTE " + n)
    if art.get("error"):
        print("READPOINT version_resource: " + str(art["error"]))
    # version is its own dimension (spec sec 11.3.4): always listed, never fatal
    ver_state = report_versions(args.tree, args.vbp, art, report,
                                manifest_dir_for(args), args.artifact)

    got = []
    if mcp and mcp.get("ok"):
        got.append(("mcp_info", mcp["stamp"]))
    if art.get("comments"):
        got.append(("version_resource", art["stamp"]))
    if not got:
        report["verdict"] = "FAIL"
        report["failures"].append("artifact_carries_no_source_stamp")
        report["failures"].append("not_this_tree_or_stale_stamp")
        print("ARTIFACT_STAMP=none")
        print("VERDICT=FAIL (the artifact carries no source stamp: no mcp_info "
              "build.sourceHash and no version-resource Comments=algo=...;src=...)")
        report["exit"] = 1
        if args.json:
            print(json.dumps(report, indent=1, sort_keys=True))
        return 1
    for name, st in got:
        print("ARTIFACT_STAMP[%s]=%s" % (name, json.dumps(st, sort_keys=True)))
    if len(got) == 2 and got[0][1].get("src") != got[1][1].get("src"):
        report["failures"].append("read_points_disagree")
        print("READPOINT_CONFLICT %s=%s vs %s=%s"
              % (got[0][0], got[0][1].get("src"), got[1][0], got[1][1].get("src")))
    _name, st = got[0]
    art_src = st.get("src")
    art_algo = st.get("algo")
    if art_algo and art_algo != ALGO:
        report["failures"].append("algorithm_mismatch")
        print("ALGO_MISMATCH artifact=%s expected=%s (recompute using the artifact's "
              "algorithm identifier)" % (art_algo, ALGO))
    compare_version_record(args, st.get("src") or stamp, ver_state, write=False)
    if art_src == stamp and not report["failures"]:
        # write the record WITHOUT re-printing the comparison (it was printed above)
        if ver_state.get("artifactFileVersion") or ver_state.get("artifactProductVersion"):
            p = write_version_record(version_record_path(manifest_dir_for(args), art_src),
                                     ALGO, art_src, args.artifact,
                                     ver_state.get("artifactFileVersion"),
                                     ver_state.get("artifactProductVersion"))
            print("VERSION_RECORD_WRITTEN " + p)
        report["verdict"] = "PASS"
        print("VERDICT=PASS (the artifact's stamp equals this tree's stamp): the "
              "artifact is consistent with this input set (modulo a sha256 collision)")
        print("EXIT=0")
        report["exit"] = 0
        if args.json:
            print(json.dumps(report, indent=1, sort_keys=True))
        return 0
    if art_src != stamp:
        report["failures"].append("source_hash_mismatch")
        print("HASH_MISMATCH artifact=%s tree=%s" % (art_src, stamp))
        base, searched = find_baseline(args, art_src or "")
        if base:
            print("BASELINE=" + base)
            try:
                report["diff"] = print_diff(read_bytes(base), payload)
            except OSError as exc:
                report["notes"].append("baseline unreadable: %r" % (exc,))
        else:
            print("BASELINE=none (no recorded manifest for the artifact's stamp)")
            print("  searched: " + ", ".join(searched))
            print("  per-file localisation needs the manifest recorded at build time "
                  "(compute/emit write it).  Without it only the single hash can be "
                  "compared; that is reported honestly instead of guessed.")
    report["verdict"] = "FAIL"
    report["failures"].append("not_this_tree_or_stale_stamp")
    print("VERDICT=FAIL (" + ",".join(report["failures"]) + "): the artifact is not "
          "from this tree, or its stamp is stale")
    print("EXIT=1")
    report["exit"] = 1
    if args.json:
        print(json.dumps(report, indent=1, sort_keys=True))
    return 1


def cmd_selftest(args):
    """Prove every property that can be proven without a VB6 build."""
    ok = True
    stamp, payload, entries, problems = compute_stamp(args.tree, args.vbp, args.mode)
    if problems:
        for p in problems:
            eprint("selftest: " + p)
        return 2

    def say(name, good, detail=""):
        nonlocal_ok = None  # noqa: F841 - keep py2/py3 parsing simple
        print("%-58s %s %s" % (name, "PASS" if good else "FAIL", detail))
        return good

    ok = say("1. sha256(manifest bytes) == I", sha256_hex(payload) == stamp, stamp[:16])
    tmp = tempfile.mkdtemp(prefix="srcstamp_selftest_")
    try:
        src_rel = [e["rel"] for e in entries]
        for rel in src_rel:
            dst = os.path.join(tmp, rel.replace("/", os.sep))
            d = os.path.dirname(dst)
            if d and not os.path.isdir(d):
                os.makedirs(d)
            shutil.copyfile(os.path.join(args.tree, rel.replace("/", os.sep)), dst)
        # The tmp tree must be addressed with the CANONICAL rel: passing args.vbp through
        # would let an ABSOLUTE --vbp escape the temp tree (os.path.join discards the earlier
        # component when the later one is absolute) and the "append the stamp key" step below
# would then write into the REAL project file. Same family as, so it is fixed
        # here instead of only in the CLI path - and guarded.
        rel_vbp = canonical_rel(args.tree, args.vbp)[0]
        vbp_abs = os.path.join(tmp, rel_vbp.replace("/", os.sep))

        def _inside(root, p):
            pa = os.path.abspath(p)
            return pa == root or pa.startswith(root + os.sep)

        def _inside_tmp(p):
            return _inside(tmp, p)

        if not _inside_tmp(vbp_abs):
            eprint("selftest: REFUSED - the temp-tree .vbp resolved outside the temp tree: %s"
                   % vbp_abs)
            return 2
        target = None
        for rel in src_rel:
            if rel.lower().endswith(".bas"):
                target = os.path.join(tmp, rel.replace("/", os.sep))
                break
        if target:
            with open(target, "ab") as f:
                f.write(b"'")
            s1, _p1, _e1, pr1 = compute_stamp(tmp, rel_vbp, args.mode)
            ok = say("2. I is a function of the input bytes (1 extra byte => new I)",
                     (not pr1) and s1 != stamp,
                     "%s -> %s" % (stamp[:16], (s1 or "ERR")[:16])) and ok
            with open(target, "r+b") as f:
                f.seek(-1, os.SEEK_END)
                f.truncate()
        with open(os.path.join(tmp, STAMP_BASENAME), "wb") as f:
            f.write(render_buildstamp(ALGO, stamp, args.mode))
        if not _inside_tmp(vbp_abs):
            eprint("selftest: REFUSED - refusing to append the stamp key outside the temp tree")
            return 2
        with open(vbp_abs, "ab") as f:
            f.write(b"%s=algo=%s;src=%s\r\n"
                    % (VBP_STAMP_KEY.encode(), ALGO.encode(), stamp.encode()))
        s2, _p2, _e2, pr2 = compute_stamp(tmp, rel_vbp, args.mode)
        ok = say("3. the stamp's own outputs (BuildStamp.bas, %s=) do not change I"
                 % VBP_STAMP_KEY,
                 (not pr2) and s2 == stamp, "%s" % ((s2 or "ERR")[:16],)) and ok
        for noise in ("VB6MCP.dll", "build.status", "Thumbs.db"):
            with open(os.path.join(tmp, noise), "wb") as f:
                f.write(b"noise")
        s3, _p3, _e3, pr3 = compute_stamp(tmp, rel_vbp, args.mode)
        ok = say("4. unrelated build products in the tree do not change I",
                 (not pr3) and s3 == stamp, "%s" % ((s3 or "ERR")[:16],)) and ok
        if target:
            with open(target, "rb") as f:
                body = f.read()
            with open(target, "wb") as f:
                f.write(body.replace(b"\r\n", b"\n"))
            ss, _ps, _es, prs = compute_stamp(tmp, rel_vbp, "semantic")
            base_s, _b, _c, _d = compute_stamp(args.tree, args.vbp, "semantic")
            ok = say("5. semantic mode is stable across CRLF -> LF normalisation",
                     (not prs) and ss == base_s, "%s" % ((ss or "ERR")[:16],)) and ok
# ---- version independence (spec sec 11.3.2 - the point of)
        def _rewrite(path, fn):
            raw = read_bytes(path)
            out = fn(raw)
            if out is None:
                return False
            with open(path, "wb") as f:
                f.write(out)
            return True

        def _bump_version_values(raw):
            seen = []
            out = []
            for line in raw.split(b"\n"):
                body = line[:-1] if line.endswith(b"\r") else line
                eol = b"\r" if line.endswith(b"\r") else b""
                key, sep, val = body.partition(b"=")
                k = key.strip().decode("latin1", "replace")
                if sep and k in VBP_VERSION_KEYS:
                    try:
                        nv = str(int(val.strip()) + 7).encode()
                    except ValueError:
                        nv = b"7"
                    out.append(key + sep + nv + eol)
                    seen.append(k)
                else:
                    out.append(line)
            return b"\n".join(out), seen

        def _force_version(raw, key, value):
            out = []
            hit = False
            for line in raw.split(b"\n"):
                body = line[:-1] if line.endswith(b"\r") else line
                eol = b"\r" if line.endswith(b"\r") else b""
                k, sep, _v = body.partition(b"=")
                if sep and k.strip() == key.encode():
                    out.append(k + sep + value.encode() + eol)
                    hit = True
                else:
                    out.append(line)
            return b"\n".join(out), hit

        # a FRESH baseline for this block: earlier cases deliberately mutate the tree (case 5
        # converts a .bas to LF and does not restore it), so comparing against the original
        # `stamp` here would measure those mutations instead of the version mask.
        _s_base, _pb, _eb, _pr_base = compute_stamp(tmp, rel_vbp, args.mode)
        _pre = read_bytes(vbp_abs)
        seen_keys = [k for k in VBP_VERSION_KEYS
                     if any(l.split(b"=")[0].strip() == k.encode()
                            for l in _pre.split(b"\n") if b"=" in l)]
        bump_raw, bumped = _bump_version_values(_pre)
        with open(vbp_abs, "wb") as f:
            f.write(bump_raw)
        s_ver, _pv4, _ev4, pr_ver = compute_stamp(tmp, rel_vbp, args.mode)
        ok = say("6. a version-field VALUE change does NOT move I (sec 11.3.2)",
                 (not pr_ver) and s_ver == _s_base,
                 "bumped=%s I=%s" % (",".join(bumped) or "none",
                                     (s_ver or "ERR")[:16])) and ok
        with open(vbp_abs, "wb") as f:
            f.write(bump_raw)
        if "AutoIncrementVer" in seen_keys:
            both_same = True
            det = []
            for val in ("0", "1"):
                r2, hit2 = _force_version(bump_raw, "AutoIncrementVer", val)
                if not hit2:
                    continue
                with open(vbp_abs, "wb") as f:
                    f.write(r2)
                sx, _px, _ex, prx = compute_stamp(tmp, rel_vbp, args.mode)
                det.append(val + "->" + (sx or "ERR")[:16])
                both_same = both_same and (not prx) and sx == _s_base
            ok = say("6b. AutoIncrementVer=0 and =1 give the SAME I", both_same,
                     " ".join(det)) and ok
            with open(vbp_abs, "wb") as f:
                f.write(bump_raw)
        else:
            ok = say("6b. AutoIncrementVer absent in this .vbp (nothing to compare)", True,
                     "keys=%s" % (",".join(seen_keys) or "none")) and ok
        with open(vbp_abs, "wb") as f:
            f.write(_pre)
        with open(vbp_abs, "ab") as f:
            f.write(b"LegalTrademarks=srcstamp-non-version-probe\r\n")
        s_nv, _pn, _en, pr_nv = compute_stamp(tmp, rel_vbp, args.mode)
        ok = say("7. a NON-version .vbp edit DOES move I (masking is not exclusion)",
                 (not pr_nv) and s_nv is not None and s_nv != _s_base,
                 "%s -> %s" % ((_s_base or "ERR")[:16], (s_nv or "ERR")[:16])) and ok
        with open(vbp_abs, "wb") as f:
            f.write(bump_raw)
        # 8 is a UNIT assertion on the mask itself, so it is meaningful whatever the .vbp has:
        # keys must survive, values must be replaced, line endings and other lines untouched,
        # and a .vbp with NO version key must come back BYTE-IDENTICAL (masking is never
        # insertion - a key someone else adds still moves I, which case 7 shows end to end).
        mini_in = b"Type=OleDll\r\nRevisionVer=9\r\nExeName32=x.dll\r\n"
        mini_out = vbp_mask_version_values(mini_in)
        no_keys_in = b"Type=OleDll\r\nExeName32=x.dll\r\n"
        ok = say("8. the mask keeps keys, replaces values, never inserts (unit)",
                 mini_out == b"Type=OleDll\r\nRevisionVer=<version-independent>\r\nExeName32=x.dll\r\n"
                 and vbp_mask_version_values(no_keys_in) == no_keys_in
                 and vbp_mask_version_values(b"MajorVer=1\r\n") == b"MajorVer=<version-independent>\r\n"
                 and vbp_mask_version_values(b"MajorVer=1\n") == b"MajorVer=<version-independent>\n",
                 "present=%s" % (",".join(seen_keys) or "none")) and ok
        # 9: HOW the tool is invoked must not change I (spec sec 10.8).  The .vbp is
        #    recorded tree-root-relative, so the caller's spelling and the tree's spelling are
        #    both irrelevant.  This is the regression guard for that trap.
        _tree_abs = os.path.abspath(args.tree)
        _vbp_abs = os.path.abspath(os.path.join(_tree_abs, args.vbp.replace("/", os.sep)))
        s_rel, _p9, _e9, pr9 = compute_stamp(args.tree, args.vbp, args.mode)
        s_abs, _p9b, _e9b, pr9b = compute_stamp(args.tree, _vbp_abs, args.mode)
        s_ta, _p9c, _e9c, pr9c = compute_stamp(_tree_abs, _vbp_abs, args.mode)
        ok = say("9. relative / absolute / absolute-tree invocation give the SAME I",
                 (not pr9) and (not pr9b) and (not pr9c) and s_rel == s_abs == s_ta,
                 "%s / %s / %s" % ((s_rel or "ERR")[:16], (s_abs or "ERR")[:16],
                                   (s_ta or "ERR")[:16])) and ok
        # ---- 10.x  the 12-class change matrix (->) ------------------------
        # PROVES THE REVERSE DIRECTION: "source changed => I changes".  The forward direction
        # ("same source => same I") is cases 1-9 above.  Runs in its OWN temp copy so the
        # mutations of the earlier cases cannot pollute the baseline, and EVERY class is
        # restored and re-checked against I0 before the next one starts.  Writes only inside
        # the copy (guarded by _inside()).
        MCLASS = []
        mtmp = tempfile.mkdtemp(prefix="srcstamp_matrix_")
        try:
            for _e in entries:
                _dst = os.path.join(mtmp, _e["rel"].replace("/", os.sep))
                _d = os.path.dirname(_dst)
                if _d and not os.path.isdir(_d):
                    os.makedirs(_d)
                shutil.copyfile(os.path.join(args.tree, _e["rel"].replace("/", os.sep)), _dst)

            def _mp(rel):
                return os.path.join(mtmp, rel.replace("/", os.sep))

            def _mrd(rel):
                f = open(_mp(rel), "rb"); b = f.read(); f.close(); return b

            def _mwr(rel, data):
                if not _inside(mtmp, _mp(rel)):
                    raise RuntimeError("matrix write escaped the temp copy: %s" % rel)
                f = open(_mp(rel), "wb"); f.write(data); f.close()

            def _mi(mode):
                st, _pl, _en, pr = compute_stamp(mtmp, rel_vbp, mode)
                if pr:
                    raise RuntimeError("matrix input set broken in %s: %r" % (mode, pr))
                return st

            M0S, M0M = _mi("strict"), _mi("semantic")
            _bas = [e["rel"] for e in entries if e["rel"].lower().endswith(".bas")][0]
            _cls = [e["rel"] for e in entries if e["rel"].lower().endswith(".cls")][0]
            _frm = [e["rel"] for e in entries if e["rel"].lower().endswith(".frm")][0]
            _bas0, _cls0, _frm0, _vbp0 = _mrd(_bas), _mrd(_cls), _mrd(_frm), _mrd(rel_vbp)
            _sep = b"\r\n" if b"\r\n" in _bas0 else b"\n"
            _newfile = "Zz_unreferenced_selftest.bas"

            def _flip(rel):
                b = bytearray(_mrd(rel))
                for _i in range(len(b)):
                    if 65 <= b[_i] <= 90:
                        b[_i] = b[_i] + 32
                        break
                _mwr(rel, bytes(b))

            def _mclass(n, label, exp_s, exp_m, mutate, restore, watch=None, tag="10"):
                _b_s, _b_m = _mi("strict"), _mi("semantic")
                mutate()
                _a_s, _a_m = _mi("strict"), _mi("semantic")
                ch_s, ch_m = (_a_s != _b_s), (_a_m != _b_m)
                restore()
                _r_s, _r_m = _mi("strict"), _mi("semantic")
                good = (ch_s == exp_s) and (ch_m == exp_m) and _r_s == M0S and _r_m == M0M
                MCLASS.append((n, good, ch_s, ch_m))
                _w = watch() if watch else None
                return say((tag + ".%-30s strict=%-8s semantic=%-8s %s%s")
                           % (label, "CHANGED" if ch_s else "same",
                              "CHANGED" if ch_m else "same",
                              "restored=I0" , (" watch=%s" % _w) if _w is not None else ""),
                           good,
                           "%s->%s / %s->%s" % (_b_s[:12], _a_s[:12], _b_m[:12], _a_m[:12]))

            ok = _mclass(1, "%s: line endings only" % _bas, True, False,
                         lambda: _mwr(_bas, _bas0.replace(b"\r\n", b"\n") if _sep == b"\r\n"
                                      else _bas0.replace(b"\n", b"\r\n")),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(2, "%s: prepend UTF-8 BOM" % _bas, True, False,
                         lambda: _mwr(_bas, b"\xef\xbb\xbf" + _bas0),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(3, "%s: final newline removed" % _bas, True, False,
                         lambda: _mwr(_bas, _bas0.rstrip(b"\n")),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(4, "%s: 1 char in a code token" % _bas, True, True,
                         lambda: _flip(_bas), lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(5, "%s: 1 line inserted" % _bas, True, True,
                         lambda: _mwr(_bas, _bas0.replace(_sep, _sep + b"' matrix inserted" + _sep, 1)),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(6, "%s: 1 line deleted" % _bas, True, True,
                         lambda: _mwr(_bas, _sep.join(_bas0.split(_sep)[:-2]) + _sep),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(7, "%s: 1 char in a code token" % _cls, True, True,
                         lambda: _flip(_cls), lambda: _mwr(_cls, _cls0)) and ok
            ok = _mclass(8, "%s: 1 char (form content)" % _frm, True, True,
                         lambda: _flip(_frm), lambda: _mwr(_frm, _frm0)) and ok
            ok = _mclass(9, "new file the .vbp does NOT reference", False, False,
                         lambda: _mwr(_newfile, b"Attribute VB_Name = \"ZzUnref\"\r\nOption Explicit\r\n"),
                         lambda: os.remove(_mp(_newfile))) and ok
            ok = _mclass(10, "%s: touch only (same bytes)" % _bas, False, False,
                         lambda: (os.utime(_mp(_bas), (2000000000, 2000000000)),),
                         lambda: _mwr(_bas, _bas0)) and ok
            ok = _mclass(11, "%s: line endings only" % rel_vbp, True, False,
                         lambda: _mwr(rel_vbp, _vbp0.replace(b"\r\n", b"\n") if b"\r\n" in _vbp0
                                      else _vbp0.replace(b"\n", b"\r\n")),
                         lambda: _mwr(rel_vbp, _vbp0)) and ok
            _f_s, _f_m = _mi("strict"), _mi("semantic")
            _bytes_ok = True
            for _e in entries:
                _fa = open(os.path.join(mtmp, _e["rel"].replace("/", os.sep)), "rb")
                _fb = open(os.path.join(args.tree, _e["rel"].replace("/", os.sep)), "rb")
                if sha256_hex(_fa.read()) != sha256_hex(_fb.read()):
                    _bytes_ok = False
                _fa.close(); _fb.close()
            MCLASS.append((12, (_f_s == M0S) and (_f_m == M0M) and _bytes_ok, False, False))
            ok = say("10.%-30s strict=%-8s semantic=%-8s (bit-exact, all inputs)"
                     % ("12 restore everything -> I0", "same", "same"),
                     (_f_s == M0S) and (_f_m == M0M) and _bytes_ok,
                     "%s / %s bytesIdentical=%s" % (_f_s[:16], _f_m[:16], _bytes_ok)) and ok

            # ---- 11.x  CASE must stay significant (->) ------------------
            # The stamp is a BYTE-level identity: a case-only change IS a source change, in
            # BOTH modes (semantic only absorbs line endings / BOM / trailing newline).  These
            # cases exist so a future "harmless" case fold cannot be introduced silently, and
            # the block ends with a falsification that INJECTS such a fold and requires the
            # cases to stop detecting - i.e. the guard is proven able to fail.
            def _sl(raw):
                return raw.split(b"\r\n") if b"\r\n" in raw else raw.split(b"\n")

            def _jl(parts, raw):
                return (b"\r\n" if b"\r\n" in raw else b"\n").join(parts)

            def _has_alpha(b):
                # bytearray iterates as INTS on both py2 and py3 (a py2 str iterates as
                # 1-char strings where "65 <= c" is silently False - that bug once hid the
                # string/comment cases from my own harness).
                return any(65 <= c <= 122 for c in bytearray(b))

            def _flip_all(b):
                out = bytearray(b)
                for _k in range(len(out)):
                    _c = out[_k]
                    if 97 <= _c <= 122:
                        out[_k] = _c - 32
                    elif 65 <= _c <= 90:
                        out[_k] = _c + 32
                return bytes(out)

            _txt = [e["rel"] for e in entries
                    if e["rel"].lower().endswith((".bas", ".cls", ".frm"))]
            _frms = [e["rel"] for e in entries if e["rel"].lower().endswith(".frm")]
            # BYTES patterns: the lines being matched are bytes, and py3 refuses to mix a
            # str pattern with a bytes subject (py2 tolerated it - hence this bit us here).
            _dcl_re = re.compile(br"^\s*(?:Public |Private |Friend |Static )*"
                                 br"(?:Sub|Function)\s+([A-Za-z_]\w*)")
            _ctl_re = re.compile(br"^\s*Begin\s+VB\.([A-Za-z0-9_.]+)\s+([A-Za-z_]\w*)")

            def _find(rels, pred):
                for _r in rels:
                    for _i, _l in enumerate(_sl(_mrd(_r))):
                        if pred(_l):
                            return _r, _i, _l
                return None, None, None

            def _find_str(rels):
                for _r in rels:
                    for _i, _l in enumerate(_sl(_mrd(_r))):
                        _s = _l.strip()
                        if _s.startswith(b"Attribute") or _s.startswith(b"'"):
                            continue
                        _segs = _l.split(b'"')
                        for _j in range(1, len(_segs), 2):
                            if _has_alpha(_segs[_j]):
                                return _r, _i, _j
                return None, None, None

            _c_specs = []          # (label, rel, idx, mutated_line) for the falsification
            _r1, _i1, _l1 = _find(_txt, lambda ln: bool(_dcl_re.match(ln))
                                  and not ln.strip().startswith(b"Attribute"))
            _r2, _i2, _l2 = _find(_txt, lambda ln: ln.strip().startswith(b"Option Explicit"))
            _r3, _i3, _j3 = _find_str(_txt)
            _r4, _i4, _l4 = _find(_txt, lambda ln: ln.strip().startswith(b"'")
                                  and _has_alpha(ln.strip()[1:]))
            _r5, _i5, _l5 = _find(_frms, lambda ln: bool(_ctl_re.match(ln))
                                  and _ctl_re.match(ln).group(1) != b"Form")
            _found = None not in (_r1, _r2, _i3, _r4, _r5)
            ok = say("11.%-30s constructs located" % "precheck", _found,
                     "decl=%s kw=%s str=%s cmt=%s ctl=%s"
                     % (_r1, _r2, _r3, _r4, _r5)) and ok
            if _found:
                def _line_with(rel, idx, newtok, oldtok):
                    # returns the MUTATED LINE (not the whole file): returning the joined file
                    # here was a real bug - _cclass then assigned it as one line, corrupting the
                    # file, and the case still reported CHANGED (so the bug hid behind a PASS).
                    _p = _sl(_mrd(rel))
                    _p[idx] = _p[idx].replace(oldtok, newtok, 1)
                    return _p[idx]

                _n1 = _line_with(_r1, _i1, _flip_all(_dcl_re.match(_l1).group(1)),
                                 _dcl_re.match(_l1).group(1))
                _n2 = _line_with(_r2, _i2, b"option explicit", b"Option Explicit")
                _sp = _sl(_mrd(_r3)); _sg = _sp[_i3].split(b'"')
                _sg[_j3] = _flip_all(_sg[_j3]); _n3 = b'"'.join(_sg)
                _n4 = _line_with(_r4, _i4, _flip_all(_l4.strip()[1:]), _l4.strip()[1:])
                _n5 = _line_with(_r5, _i5, _flip_all(_ctl_re.match(_l5).group(2)),
                                 _ctl_re.match(_l5).group(2))

                def _cclass(n, label, rel, idx, newline):
                    _orig = _mrd(rel)
                    _p = _sl(_orig)
                    if idx >= len(_p):
                        return say("11.%-30s" % label, False, "line index out of range")
                    _p[idx] = newline
                    _mut = _jl(_p, _orig)
                    # PURITY: this block claims "a case-only change is still a source change",
                    # so the mutation must be proven case-only - same length, actually
                    # different, and equal after lower().  Without this, a corrupting edit
                    # would still report CHANGED and the case would be a lie.
                    _pure = (len(_mut) == len(_orig)) and (_mut != _orig) and \
                            (_mut.lower() == _orig.lower())
                    if not _pure:
                        return say("11.%-30s" % (label + " [purity]"), False,
                                   "len %d->%d changed=%s lowerEqual=%s"
                                   % (len(_orig), len(_mut), _mut != _orig,
                                      _mut.lower() == _orig.lower()))
                    _c_specs.append((label, rel, idx, _mut))
                    return _mclass(n, label, True, True,
                                   lambda: _mwr(rel, _mut), lambda: _mwr(rel, _orig), tag="11")

                _wd = _dcl_re.match(_l1).group(1)
                ok = _cclass(1, "identifier case (%s)" % _wd, _r1, _i1, _n1) and ok
                ok = _cclass(2, "keyword case (Option Explicit)", _r2, _i2, _n2) and ok
                ok = _cclass(3, "string-literal case (%s)" % _r3, _r3, _i3, _n3) and ok
                ok = _cclass(4, "comment case (%s)" % _r4, _r4, _i4, _n4) and ok
                ok = _cclass(5, ".frm control NAME case (%s)" % _r5, _r5, _i5, _n5) and ok

                # FALSIFICATION: inject a case fold into the identity computation and require
                # the five cases above to STOP detecting their change.  If they still detected
                # it, these assertions would be vacuous.
                _real_mp = manifest_payload
                _masked = 0
                try:
                    def _fold_mp(_entries, _mode):
                        _fo = []
                        for _e2 in _entries:
                            _d2 = dict(_e2)
                            _d2["hashed"] = _e2["hashed"].lower()
                            _fo.append(_d2)
                        return _real_mp(_fo, _mode)

                    globals()["manifest_payload"] = _fold_mp
                    for _lab, _rel, _idx, _mut in _c_specs:
                        _orig = _mrd(_rel)
                        _mwr(_rel, _orig)
                        _f_before = _mi("strict")
                        _mwr(_rel, _mut)
                        _f_after = _mi("strict")
                        _mwr(_rel, _orig)
                        if _f_before == _f_after:
                            _masked += 1
                finally:
                    globals()["manifest_payload"] = _real_mp
                ok = say("11.%-30s lower() injected: masked %d/%d"
                         % ("FALSIFICATION", _masked, len(_c_specs)),
                         _masked == len(_c_specs),
                         "with a case fold injected into the identity, all %d case cases stop "
                         "detecting their change - i.e. these assertions CAN fail"
                         % len(_c_specs)) and ok

            # --- the mask's OWN falsification: disable semantic absorption and the
            #     representation-only classes MUST start failing.  A test that cannot fail
            #     proves nothing, so this is asserted, not assumed.
            _real_norm = normalize_semantic
            try:
                globals()["normalize_semantic"] = lambda raw, ext: raw
                _b0 = _mwr(_bas, _bas0) or M0S
                _real_i0s, _real_i0m = _mi("strict"), _mi("semantic")
                _mwr(_bas, _bas0.rstrip(b"\n"))
                _neg_i = _mi("semantic")
                _mwr(_bas, _bas0)
                _detected = (_neg_i != _real_i0m)
            finally:
                globals()["normalize_semantic"] = _real_norm
            ok = say("10.%-30s semantic-sabotage detected" % "FALSIFICATION",
                     _detected,
                     "with normalisation disabled, class 3 DOES move semantic I: %s" % _detected) and ok
            _mc_bad = [c for c in MCLASS if not c[1]]
            ok = say("10.%d 12-class matrix (reverse direction: change => dI)" % 99,
                     not _mc_bad,
                     "classes=%d ok=%d unexpected=%d" % (len(MCLASS), len(MCLASS) - len(_mc_bad),
                                                         len(_mc_bad))) and ok
        except Exception as _mexc:  # noqa: BLE001 - reported, never silently swallowed
            ok = say("10.x 12-class matrix crashed", False, repr(_mexc)) and ok
        finally:
            shutil.rmtree(mtmp, ignore_errors=True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("SELFTEST=" + ("PASS" if ok else "FAIL"))
    print("EXIT=" + ("0" if ok else "1"))
    return 0 if ok else 1


def build_parser():
    p = argparse.ArgumentParser(
        prog="srcstamp.py",
        description="Source stamp (vb6src/2): a stable indicator of a VB6 source "
                    "input set.  Equal indicator => same origin.")
    sub = p.add_subparsers(dest="cmd")

    def common(sp):
        sp.add_argument("tree")
        sp.add_argument("--vbp", default=DEFAULT_VBP)
        sp.add_argument("--mode", choices=MODES, default="strict")
        sp.add_argument("--manifest-dir", default=None,
                        help="where manifests are recorded (default: ./dist)")
        sp.add_argument("--json", action="store_true")

    c = sub.add_parser("compute", help="compute I and record the manifest")
    c.description = ("Compute the source stamp I for a tree and record the manifest "
                     "(I == sha256 of the manifest file).  THIS WRITES FILES by "
                     "default; a read-only check must pass --no-write.")
    common(c)
    c.add_argument("--manifest", default=None, help="write the manifest to this path")
    c.add_argument("--no-write", action="store_true",
                   help="compute I and REPORT what would be written (created / "
                        "rewritten / byte-identical) but write NOTHING; the read-only "
                        "form every check_* caller must use")

    e = sub.add_parser("emit", help="write BuildStamp.bas and the .vbp Comments= line")
    e.description = ("Write BuildStamp.bas (--out) and idempotently set the .vbp "
                     "Comments= line (--vbp-patch).  Both are excluded from the input "
                     "set, so emitting never changes I.  Run this AFTER the .vbp lists "
                     "Module=BuildStamp; BuildStamp.bas and BEFORE the build.")
    common(e)
    e.add_argument("--out", default=None,
                   help="BuildStamp.bas path (omitted: print the line instead of writing)")
    e.add_argument("--vbp-patch", default=None, metavar="VBP",
                   help="idempotently set Comments= in THIS .vbp instead of the hashed "
                        "project file (default: the --vbp file inside the tree)")
    e.add_argument("--no-vbp-patch", action="store_true",
                   help="do not touch any .vbp (dry run); the Comments= line is then "
                        "only printed")

    k = sub.add_parser("check", help="assert that the artifact comes from this tree")
    k.description = ("Read I out of the artifact (mcp_info when --mcp-url is given, "
                     "otherwise the version resource via GetFileVersionInfo) and compare "
                     "it with I recomputed from the tree.  Equal => same origin; a "
                     "missing stamp is a FAIL, never a skip.")
    k.add_argument("artifact")
    k.add_argument("tree")
    k.add_argument("--vbp", default=DEFAULT_VBP)
    k.add_argument("--mode", choices=MODES, default="strict")
    k.add_argument("--mcp-url", default=None)
    k.add_argument("--baseline", default=None,
                   help="manifest to diff against when the hashes differ")
    k.add_argument("--manifest-dir", default=None)
    k.add_argument("--json", action="store_true")

    s = sub.add_parser("selftest", help="prove the invariants that need no VB6 build")
    s.description = ("Copy the input set to a temporary tree and check every property "
                     "that does not need a VB6 build: I is a function of the input "
                     "bytes, and the stamp's own outputs / build products never move I.")
    common(s)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    if not args.cmd:
        build_parser().print_help()
        return 2
    if args.cmd == "compute":
        return cmd_compute(args)
    if args.cmd == "emit":
        return cmd_emit(args)
    if args.cmd == "check":
        return cmd_check(args)
    if args.cmd == "selftest":
        return cmd_selftest(args)
    return 2


if __name__ == "__main__":
    sys.exit(main())
