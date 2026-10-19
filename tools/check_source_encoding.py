#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""task-154 - build-source encoding guard for the VB6 project.

WHY
    VB6 reads every .bas / .cls / .frm as the SYSTEM ANSI CODE PAGE (ACP; 936/GBK
    on a Chinese Windows).  It does NOT look for a UTF-8 BOM and it is not aware of
    UTF-8 at all.  A source file stored as UTF-8 is therefore read as GBK:
      * a non-ASCII STRING LITERAL compiles to MOJIBAKE inside the artifact
        (measured: CodeTools.cls "debug" text reached the shipped VB6MCP.dll as
        U+748B U+51AD U+75EF and the tool face served it to every MCP client), and
      * a UTF-8 BOM makes VB6 refuse the whole project
        (docs/REPRODUCIBLE_BUILDS.md section 8.3: /make rc=1, zero artifact,
         "cannot load, line 0: expected: expression").
    Comments are not compiled, so non-ASCII in a comment is a DISPLAY issue only
    (it shows as mojibake in the IDE) - but it is the same class of drift, so it is
    reported and the file must still be ACP-encoded.

CONVENTION ENFORCED (docs/SOURCE_ENCODING.md; CONTRIBUTING.md already said ANSI/GBK)
    C1  every file listed by MCP_VB6.vbp is stored as 7-bit ASCII or as the system
        ANSI code page (GBK) WITHOUT a BOM.  UTF-8 is refused for build components.
    C2  inside a build component, non-ASCII is allowed in COMMENTS only.  A string
        literal must be 7-bit ASCII; user-visible non-ASCII text is built with
        ChrW$(&H....) so the artifact does not depend on the build machine's ACP.
    C3  non-ASCII outside strings and comments ("code position") is always an error;
        VB6 identifiers are ASCII.

FINDING KINDS
    ERROR (exit 1)
      BOM_PRESENT      leading UTF-8/UTF-16 BOM - VB6 refuses to load the project
      UNDECODABLE      neither valid UTF-8 nor valid GBK
      NON_ACP_SOURCE   valid UTF-8 *with non-ASCII* -> VB6 will misread these bytes
      CODE_NON_ASCII   non-ASCII in code position (identifier / stray character)
    WARN  (ERROR under --strict)
      LITERAL_NON_ASCII  non-ASCII inside a string literal of an ACP file.  It DOES
                         compile correctly on this build machine (ACP=936), but the
                         compiled text then depends on the build machine's ACP.

USAGE
    python tools/check_source_encoding.py [--root .] [--vbp MCP_VB6.vbp]
                                          [--strict] [--json OUT.json]
                                          [--paths a.bas,b.cls]   (audit extra files)
EXIT
    0 = no ERROR (and no WARN when --strict)
    1 = violations
    2 = cannot evaluate (vbp / component missing)
"""
from __future__ import print_function

import codecs
import glob
import io
import json
import os
import re
import subprocess
import sys

IDCH = re.compile(r"[A-Za-z0-9_]")
BOMS = ((codecs.BOM_UTF8, "utf-8"), (codecs.BOM_UTF16_LE, "utf-16le"),
        (codecs.BOM_UTF16_BE, "utf-16be"))


def parse_arg(name, default=None):
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return default


def read_bytes(path):
    with open(path, "rb") as fh:
        return fh.read()


def classify(raw):
    """-> (encoding, bom).  encoding in ascii|utf-8|gbk|unknown|utf-16*"""
    for bom, name in BOMS:
        if raw.startswith(bom):
            return name, True
    if len(raw) == 0:
        return "ascii", False
    if all(c < 0x80 for c in raw):
        return "ascii", False
    try:
        raw.decode("utf-8")
        return "utf-8", False
    except UnicodeDecodeError:
        pass
    try:
        raw.decode("cp936")
        return "gbk", False
    except UnicodeDecodeError:
        return "unknown", False


def decode(raw, enc):
    if enc == "utf-16le":
        return raw.decode("utf-16-le", "replace")
    if enc == "utf-16be":
        return raw.decode("utf-16-be", "replace")
    if enc == "utf-8":
        return raw.decode("utf-8", "replace")
    if enc == "gbk":
        return raw.decode("cp936", "replace")
    return raw.decode("latin-1", "replace")


def scan_positions(text):
    """Yield dicts for every non-ASCII character with its region."""
    for ln, line in enumerate(text.split("\n"), 1):
        line = line.rstrip("\r")
        i, n, in_str = 0, len(line), False
        while i < n:
            ch = line[i]
            if in_str:
                if ch == '"':
                    if i + 1 < n and line[i + 1] == '"':
                        i += 2
                        continue
                    in_str = False
                    i += 1
                    continue
                if ord(ch) > 127:
                    yield {"line": ln, "col": i + 1, "region": "literal", "ch": ch}
                i += 1
                continue
            if ch == '"':
                in_str = True
                i += 1
                continue
            if ch == "'":
                for j in range(i, n):
                    if ord(line[j]) > 127:
                        yield {"line": ln, "col": j + 1, "region": "comment",
                               "ch": line[j]}
                break
            if line[i:i + 3].lower() == "rem" and (i + 3 >= n or line[i + 3] in " \t"):
                pre = line[:i].strip()
                if pre == "" or pre.endswith(":"):
                    for j in range(i, n):
                        if ord(line[j]) > 127:
                            yield {"line": ln, "col": j + 1, "region": "comment",
                                   "ch": line[j]}
                    break
            if ord(ch) > 127:
                left = i > 0 and bool(IDCH.match(line[i - 1]))
                right = i + 1 < n and bool(IDCH.match(line[i + 1]))
                yield {"line": ln, "col": i + 1, "region": "code", "ch": ch,
                       "in_identifier": bool(left or right)}
            i += 1


def vbp_components(vbp_path):
    comps = []
    with io.open(vbp_path, "r", encoding="latin-1") as fh:
        for line in fh.read().splitlines():
            m = re.match(r"^(Form|Class|Module|UserControl|PropertyPage|UserDocument|"
                         r"Designer|RelatedDoc|ResFile32)\s*=\s*(.+)$", line)
            if not m:
                continue
            val = m.group(2).strip()
            src = val.split(";", 1)[1].strip() if ";" in val else val
            if src:
                comps.append({"kind": m.group(1), "file": src})
    return comps


def audit_file(path, relname):
    raw = read_bytes(path)
    enc, bom = classify(raw)
    findings = []
    if bom:
        findings.append({"kind": "BOM_PRESENT", "severity": "ERROR",
                         "encoding": enc, "detail": "leading BOM of %s" % enc})
    if enc == "unknown":
        findings.append({"kind": "UNDECODABLE", "severity": "ERROR", "encoding": enc,
                         "detail": "neither valid UTF-8 nor valid GBK"})
        return {"file": relname, "bytes": len(raw), "encoding": enc, "bom": bom,
                "findings": findings, "non_ascii": {}}
    if enc == "utf-8":
        findings.append({"kind": "NON_ACP_SOURCE", "severity": "ERROR",
                         "encoding": enc,
                         "detail": "valid UTF-8 with non-ASCII bytes; VB6 reads these "
                                   "bytes as ACP(936) and misreads them"})
    text = decode(raw, enc)
    counts = {"comment": 0, "literal": 0, "code": 0}
    samples = {"code": [], "literal": []}
    for h in scan_positions(text):
        counts[h["region"]] += 1
        if h["region"] == "code":
            findings.append({"kind": "CODE_NON_ASCII", "severity": "ERROR",
                             "line": h["line"], "col": h["col"], "ch": h["ch"],
                             "in_identifier": h["in_identifier"]})
        elif h["region"] == "literal":
            if len(samples["literal"]) < 10:
                samples["literal"].append({"line": h["line"], "col": h["col"],
                                           "ch": h["ch"]})
    if counts["literal"]:
        findings.append({"kind": "LITERAL_NON_ASCII", "severity": "WARN",
                         "count": counts["literal"],
                         "detail": "non-ASCII inside a string literal; compiles "
                                   "correctly on ACP=936 but is ACP-dependent",
                         "where": samples["literal"]})
    return {"file": relname, "bytes": len(raw), "encoding": enc, "bom": bom,
            "non_ascii": counts, "findings": findings}


# --- the MCP-side sources are Python/Node, NOT VB build components ------------
# Why this exists (measured 2026-10-17): the v1.11.8 internal-identifier cleanup
# rewrote bridge/vb6mcp_bridge.py and pulled the _send docstring line back to
# column 0.  The VB project still built, every gate still passed (no gate ever
# compiled a .py or checked a .mjs), and the only symptom was "the middle service
# cannot start" - reported by the project side, not by us.  So this gate now
# compiles every shipped .py, runs node --check on every shipped .mjs, and
# refuses a UTF-8 BOM on them (a BOM on markdown is a WARN: harmless, reported).
def audit_mcp_sources(root):
    findings = []

    def err(rel, detail, sev="ERROR"):
        findings.append({"kind": "SYNTAX" if sev == "ERROR" else "BOM_MARKDOWN",
                         "severity": sev, "file": rel, "detail": detail})

    for pat in ("bridge/*.py", "tools/*.py"):
        for p in sorted(glob.glob(os.path.join(root, pat))):
            rel = os.path.relpath(p, root)
            raw = read_bytes(p)
            if raw.startswith(codecs.BOM_UTF8):
                err(rel, "UTF-8 BOM on a Python source (strip it)")
                raw = raw[len(codecs.BOM_UTF8):]
            try:
                compile(raw, rel, "exec")
            except SyntaxError as ex:
                err(rel, "line %s: %s" % (ex.lineno, ex.msg))

    for p in sorted(glob.glob(os.path.join(root, "bridge/*.mjs"))):
        rel = os.path.relpath(p, root)
        if read_bytes(p).startswith(codecs.BOM_UTF8):
            err(rel, "UTF-8 BOM on an ES module (strip it)")
        try:
            cp = subprocess.run(["node", "--check", p],
                                capture_output=True, text=True)
        except OSError:
            findings.append({"kind": "SYNTAX", "severity": "WARN",
                             "file": rel, "detail": "node missing - not checked"})
            continue
        if cp.returncode != 0:
            tail = (cp.stderr or "").strip().splitlines()
            err(rel, tail[-1] if tail else "node --check failed")

    for pat in ("skills/*/SKILL.md", "docs/*.md"):
        for p in sorted(glob.glob(os.path.join(root, pat))):
            rel = os.path.relpath(p, root)
            if read_bytes(p).startswith(codecs.BOM_UTF8):
                err(rel, "UTF-8 BOM on markdown", sev="WARN")
    return findings


def main():
    root = os.path.abspath(parse_arg("--root", "."))
    vbp = parse_arg("--vbp", "MCP_VB6.vbp")
    strict = "--strict" in sys.argv
    json_out = parse_arg("--json")
    extra = [p for p in (parse_arg("--paths", "") or "").split(",") if p.strip()]

    vbp_path = os.path.join(root, vbp)
    if not os.path.exists(vbp_path):
        print("FAIL cannot evaluate: %s not found" % vbp_path)
        return 2
    comps = vbp_components(vbp_path)
    if not comps:
        print("FAIL cannot evaluate: %s declares no components" % vbp_path)
        return 2

    targets = []
    missing = []
    for c in comps:
        p = os.path.join(root, c["file"])
        if os.path.exists(p):
            targets.append((p, c["file"], c["kind"]))
        else:
            missing.append(c["file"])
    for e in extra:
        p = e if os.path.isabs(e) else os.path.join(root, e)
        targets.append((p, e, "extra"))

    rows = []
    errors = 0
    warns = 0
    for p, rel, kind in targets:
        try:
            r = audit_file(p, rel)
        except Exception as ex:
            r = {"file": rel, "bytes": -1, "encoding": "?", "bom": False,
                 "findings": [{"kind": "UNREADABLE", "severity": "ERROR",
                               "detail": repr(ex)}], "non_ascii": {}}
        r["kind"] = kind
        rows.append(r)
        for f in r["findings"]:
            sev = f["severity"]
            if sev == "WARN" and strict:
                f["severity"] = "ERROR(strict)"
                sev = "ERROR"
            if sev == "ERROR":
                errors += 1
            else:
                warns += 1

    for r in rows:
        for f in r["findings"]:
            loc = ""
            if "line" in f:
                loc = " @%d:%d" % (f["line"], f["col"])
            elif f.get("where"):
                loc = " @%s" % ",".join(str(w["line"]) for w in f["where"][:6])
            extra_txt = ""
            if "count" in f:
                extra_txt = " x%d" % f["count"]
            print("%-9s %-18s %-22s%s%s" % (f["severity"], r["file"], f["kind"],
                                            extra_txt, loc))

    print()
    print("FILES=%d  (components=%d extras=%d missing=%d)"
          % (len(rows), len(comps), len(extra), len(missing)))
    if missing:
        print("MISSING=" + ",".join(missing))
    tally = {}
    for r in rows:
        for f in r["findings"]:
            tally[f["kind"]] = tally.get(f["kind"], 0) + 1
    print("ENCODINGS=" + ",".join(sorted(set(r["encoding"] for r in rows))))
    print("ERRORS=%d WARNS=%d%s" % (errors, warns,
                                    " (strict)" if strict else ""))
    for k in sorted(tally):
        print("  %-18s %d" % (k, tally[k]))
    if json_out:
        with io.open(json_out, "w", encoding="utf-8") as fh:
            json.dump({"root": root, "vbp": vbp, "strict": strict,
                       "components": len(comps), "missing": missing,
                       "errors": errors, "warns": warns, "files": rows},
                      fh, ensure_ascii=False, indent=1)
    syn = audit_mcp_sources(root)
    syn_errors = 0
    for f in syn:
        print("%-9s %-18s %-22s %s" % (f["severity"], f["file"], f["kind"], f["detail"]))
        if f["severity"] == "ERROR":
            syn_errors += 1
            errors += 1
        else:
            warns += 1
    print("MCP_SOURCE_FINDINGS=%d  SYNTAX_ERRORS=%d" % (len(syn), syn_errors))
    print("VERDICT=%s" % ("PASS" if errors == 0 else "FAIL"))
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
