#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lessons (experience library / 经验库) -- authoritative Python implementation.

This is a 1:1 port of the VB add-in module Lessons.bas (593 lines) so that the
MCP hub can own the module after the add-in drops it.  Storage, schema, merge
semantics, key names and the receipt shapes are unchanged:

    mcp_lessons/defaults.json   shipped, READ-ONLY
                                {"_meta":{...},"instructions":{"intro":..,"rules":[..]},"lessons":[..]}
    mcp_lessons/user.json       writable: {"lessons":[..]}

Semantics ported from Lessons.bas
---------------------------------
* MergedLessons   : defaults first, then user.  An entry whose id already exists
                    (case-insensitive) OVERWRITES the keys of the earlier entry
                    (except the "id" key itself) and KEEPS the earlier position;
                    order is first-appearance order.  Entries without a usable
                    id are skipped.  Disabled entries (enabled == false, with
                    JsonGetBool semantics) are filtered unless includeDisabled.
* Instructions    : intro + CRLF + "N) rule" lines joined by CRLF, from
                    defaults.json only.  It renders the INTRO + NUMBERED RULES
                    ONLY -- the lessons are NOT part of this string (verified
                    against re_work/t1102/instructions_full_golden.txt: the
                    rendered text equals golden[2121:] byte for byte).
* AddJson         : scenario/trap/do all required (non-empty); id defaults to the
                    first free custom-N (N < 1000, else custom-x); an existing id
                    is REPLACED IN PLACE (same index, first match only); the user
                    store is .bak'd and rewritten atomically.
* DeleteJson      : removes every user entry with that id; if the id still exists
                    in defaults.json a minimal {"id":..,"enabled":false} entry is
                    APPENDED (action "disabled(default)"), else "removed"; an
                    unknown id is "not_found".
* ExportMd        : exact markdown format of Lessons.ExportMd, CRLF line ends,
                    坑/应 markers (ChrW$(&H5751)/ChrW$(&H5E94)).
* ImportMdJson    : '## [id] scenario' headers + 坑:/trap: and 应:/do: lines;
                    ids already in the MERGED library are never overwritten
                    (they fall back to custom-N and count as a conflict).
* Store safety (TASK-P1-038): a user.json that exists but is empty or does not
                    parse is NEVER silently replaced by an empty store -- the file
                    is left untouched, a timestamped .bak is kept and the call
                    fails with error="lesson_store_unreadable".

CLI (frozen contract -- do not rename verbs/options):
    python tools/lessons.py instructions
    python tools/lessons.py list    [--include-disabled]
    python tools/lessons.py add     --id X --scenario S --trap T --do D [--scope mcp|vb6|all] [--enabled 0|1]
    python tools/lessons.py delete  --id X
    python tools/lessons.py export  [--include-disabled]
    python tools/lessons.py import-md --path P [--dry-run]
    python tools/lessons.py validate
    [--data-dir DIR]  global, for tests: read/write a COPY of the two stores

stdout contract: exactly ONE compact JSON object, UTF-8, no BOM, one trailing
newline; nothing else is ever written to stdout (diagnostics go to stderr).
Failures print {"ok":false,"error":"<code>","message":"..."} and exit non-zero.
"""
import argparse
import io
import json
import os
import shutil
import sys
import time

DIR_NAME = "mcp_lessons"
DEFAULTS_LEAF = "defaults.json"
USER_LEAF = "user.json"
EOL = "\r\n"                 # vbCrLf
MARK_TRAP = "\u5751"             # 坑 = ChrW$(&H5751)
MARK_DO = "\u5e94"               # 应 = ChrW$(&H5E94)
DEFAULT_INTRO = "VB6MCP usage notes (read before calling tools):"
EXPORT_TITLE = "# VB6MCP lessons (re-import with import_lessons_md)"
IMPORT_FORMAT_HINT = ("no lessons parsed: expected '## [id] scenario' headers plus "
                      "trap/do lines (run export_lessons to see the exact format)")


class LessonError(Exception):
    """A failure the CLI turns into {"ok":false,"error":code,"message":...}."""

    def __init__(self, code, message, **extra):
        Exception.__init__(self, message)
        self.code = code
        self.message = message
        self.extra = extra


# ---------------------------------------------------------------------------
# paths / files  (Lessons.LessonsDir / FilePath / EnsureDir / ReadText / WriteText)
# ---------------------------------------------------------------------------

def default_data_dir():
    """<repo root>/mcp_lessons  (the add-in used App.Path, i.e. next to the DLL)."""
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.dirname(here), DIR_NAME)


def resolve_data_dir(data_dir):
    if data_dir:
        return os.path.abspath(data_dir)
    return default_data_dir()


def file_path(data_dir, leaf):
    return os.path.join(data_dir, leaf)


def read_text(path):
    """Lenient reader: VB used Encoding.DetectEncoding + DecodeBytes.

    Missing/unreadable -> "" (exactly like Lessons.ReadText).  A UTF-8 BOM is
    stripped; UTF-16 BOMs are honoured; otherwise UTF-8, then GBK, then latin-1.
    """
    try:
        with open(path, "rb") as fh:
            raw = fh.read()
    except (IOError, OSError):
        return ""
    if not raw:
        return ""
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", "replace")
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le", "replace")
    if raw.startswith(b"\xfe\xff"):
        return raw[2:].decode("utf-16-be", "replace")
    for enc in ("utf-8", "cp936", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def write_text(path, text):
    """Atomic write (Lessons.WriteText): utf-8, no BOM, tmp file then rename.

    Returns False when nothing could be written.  JSON is serialised without a
    trailing newline, matching Lessons.JsonStringify / the shipped user.json.
    """
    directory = os.path.dirname(path)
    if directory and not os.path.isdir(directory):
        try:
            os.makedirs(directory)
        except (IOError, OSError):
            pass
    data = text.encode("utf-8")
    tmp = path + ".tmp"
    try:
        with open(tmp, "wb") as fh:
            fh.write(data)
    except (IOError, OSError):
        return False
    try:
        os.replace(tmp, path)
        return True
    except (IOError, OSError):
        pass
    # Rename refused (target locked/absent): fall back to a direct write so the
    # update is not lost, exactly like Lessons.WriteText.
    try:
        with open(path, "wb") as fh:
            fh.write(data)
        try:
            os.unlink(tmp)
        except (IOError, OSError):
            pass
        return True
    except (IOError, OSError):
        try:
            os.unlink(tmp)
        except (IOError, OSError):
            pass
        return False


def backup(path):
    """Lessons.Backup: copy to <path>.bak_<yyyymmdd_hhnnss>.  "" when absent.

    VB's FileCopy silently fails when the second-granularity name already exists,
    which would leave a write without a backup; this port adds a numeric suffix
    instead so that a backup ALWAYS exists for every write.
    """
    if not os.path.exists(path):
        return ""
    base = path + ".bak_" + time.strftime("%Y%m%d_%H%M%S")
    dest = base
    n = 0
    while os.path.exists(dest) and n < 100:
        n += 1
        dest = base + "_" + str(n)
    try:
        shutil.copyfile(path, dest)
        return dest
    except (IOError, OSError):
        return ""


# ---------------------------------------------------------------------------
# VB JsonLib value semantics (JsonGetStr / JsonGetBool / CStr / JsonStringify)
# ---------------------------------------------------------------------------

def vb_str(value, default=""):
    """JsonGetStr semantics: CStr(v) for any non-null scalar, else the default."""
    if value is None:
        return default
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return default
        if value.is_integer():
            return str(int(value))
        out = repr(value)
        if out.startswith("0."):
            out = out[1:]          # VB6 CStr(0.1) == ".1"
        elif out.startswith("-0."):
            out = "-" + out[2:]
        return out
    return default                 # dict / list / anything else


def vb_bool(value, default=False):
    """JsonGetBool semantics: real booleans pass through, else "true"/"1" wins."""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (dict, list)):
        return default
    return vb_str(value, "").lower() in ("true", "1")


def json_dumps(value):
    """Lessons.JsonStringify: compact, raw UTF-8, no trailing newline."""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"),
                      allow_nan=False)


# ---------------------------------------------------------------------------
# store loading (LoadDoc / LoadUserDocStrict / StoreUnreadableJson)
# ---------------------------------------------------------------------------

def _parse_object(text):
    if text.strip() == "":
        return None
    try:
        doc = json.loads(text)
    except ValueError:
        return None
    if isinstance(doc, dict):
        return doc
    return None


def load_doc(data_dir, leaf):
    """Lessons.LoadDoc: an object document, else None (never raises)."""
    return _parse_object(read_text(file_path(data_dir, leaf)))


def load_user_doc_strict(data_dir):
    """Lessons.LoadUserDocStrict -> (doc|None, why, backup_path)."""
    path = file_path(data_dir, USER_LEAF)
    if not os.path.exists(path):
        return {}, "missing", ""
    text = read_text(path)
    if text.strip() == "":
        return None, "empty", backup(path)
    doc = _parse_object(text)
    if doc is None:
        return None, "unreadable", backup(path)
    return doc, "", ""


def store_unreadable_error(data_dir, why, backup_path):
    path = file_path(data_dir, USER_LEAF)
    return LessonError(
        "lesson_store_unreadable",
        "the lesson store exists but could not be parsed (%s); it was left UNCHANGED" % why,
        reason="lesson_store_unreadable",
        why=why,
        path=path,
        backupPath=backup_path,
        advice=("repair or remove %s (a timestamped backup was kept when possible), "
                "then retry; nothing was written" % path),
    )


# ---------------------------------------------------------------------------
# merging (CloneDict / MergeInto / MergeList / MergedLessons)
# ---------------------------------------------------------------------------

def _merge_into(dst, src):
    for key, value in src.items():
        if vb_str(key, "").lower() != "id":
            dst[key] = value


def _merge_list(doc, order, by_id):
    if not isinstance(doc, dict):
        return
    arr = doc.get("lessons")
    if not isinstance(arr, list):
        return
    for item in arr:
        if not isinstance(item, dict):
            continue
        key = vb_str(item.get("id"), "").lower()
        if not key:
            continue
        if key in by_id:
            _merge_into(by_id[key], item)
        else:
            by_id[key] = dict(item)
            order.append(key)


def merged_lessons(data_dir, include_disabled):
    """Lessons.MergedLessons: see the module docstring for the merge rules."""
    order = []
    by_id = {}
    _merge_list(load_doc(data_dir, DEFAULTS_LEAF), order, by_id)
    _merge_list(load_doc(data_dir, USER_LEAF), order, by_id)
    out = []
    for key in order:
        item = by_id[key]
        if include_disabled or vb_bool(item.get("enabled"), True):
            out.append(item)
    return out


def merged_id_set(data_dir):
    """Lessons.MergedIdSet: ids of the merged library, disabled included."""
    seen = set()
    for item in merged_lessons(data_dir, True):
        key = vb_str(item.get("id"), "").lower()
        if key:
            seen.add(key)
    return seen


# ---------------------------------------------------------------------------
# instructions (RulesText / Instructions)
# ---------------------------------------------------------------------------

def _rules_text(inst):
    if not isinstance(inst, dict) or "rules" not in inst:
        return ""
    rules = inst["rules"]
    if isinstance(rules, list):
        parts = []
        for i, rule in enumerate(rules, 1):
            parts.append("%d) %s" % (i, vb_str(rule, "")))
        return EOL.join(parts)
    if rules is None:
        return ""
    return vb_str(rules, "")


def instructions_text(data_dir):
    """Lessons.Instructions: intro + numbered rules.  Lessons are NOT included."""
    intro = DEFAULT_INTRO
    body = ""
    doc = load_doc(data_dir, DEFAULTS_LEAF)
    if isinstance(doc, dict):
        inst = doc.get("instructions")
        if isinstance(inst, dict):
            intro = vb_str(inst.get("intro"), intro)
            body = _rules_text(inst)
    if body:
        return intro + EOL + body
    return intro


# ---------------------------------------------------------------------------
# list / add / delete  (ListJson / AddJson / DeleteJson)
# ---------------------------------------------------------------------------

def list_json(data_dir, include_disabled):
    col = merged_lessons(data_dir, include_disabled)
    lessons = []
    for item in col:
        lessons.append({
            "id": vb_str(item.get("id"), ""),
            "scenario": vb_str(item.get("scenario"), ""),
            "trap": vb_str(item.get("trap"), ""),
            "do": vb_str(item.get("do"), ""),
            "scope": vb_str(item.get("scope"), "all"),
            "enabled": vb_bool(item.get("enabled"), True),
        })
    return {"ok": True, "count": len(lessons), "lessons": lessons}


def _user_lessons(doc):
    arr = doc.get("lessons")
    if isinstance(arr, list):
        return arr
    arr = []
    doc["lessons"] = arr
    return arr


def next_custom_id(lessons):
    """Lessons.NextCustomId: custom-1 .. custom-999, else custom-x."""
    n = 1
    while n < 1000:
        cand = "custom-%d" % n
        used = False
        for item in lessons:
            if isinstance(item, dict) and \
                    vb_str(item.get("id"), "").lower() == cand.lower():
                used = True
                break
        if not used:
            return cand
        n += 1
    return "custom-x"


def add_json(data_dir, lesson_id, scenario, trap, do_text, scope, enabled):
    if not scenario or not trap or not do_text:
        raise LessonError("lesson_fields_required",
                          "scenario, trap and do are all required")
    doc, why, backup_path = load_user_doc_strict(data_dir)
    if doc is None:
        raise store_unreadable_error(data_dir, why, backup_path)
    lessons = _user_lessons(doc)
    if not lesson_id:
        lesson_id = next_custom_id(lessons)
    item = {"id": lesson_id}
    item["scope"] = scope if scope else "all"
    item["enabled"] = bool(enabled)
    item["scenario"] = scenario
    item["trap"] = trap
    item["do"] = do_text
    action = "added"
    for i, existing in enumerate(lessons):
        if isinstance(existing, dict) and \
                vb_str(existing.get("id"), "").lower() == lesson_id.lower():
            lessons[i] = item            # VB: Remove i then Add item before i
            action = "updated"
            break
    else:
        lessons.append(item)
    backup(file_path(data_dir, USER_LEAF))
    if not write_text(file_path(data_dir, USER_LEAF), json_dumps(doc)):
        raise LessonError("lesson_write_failed", "cannot write user.json")
    return {"ok": True, "id": lesson_id, "action": action,
            "lessonsTotal": len(merged_lessons(data_dir, False))}


def in_defaults(data_dir, lesson_id):
    doc = load_doc(data_dir, DEFAULTS_LEAF)
    if not isinstance(doc, dict):
        return False
    arr = doc.get("lessons")
    if not isinstance(arr, list):
        return False
    for item in arr:
        if isinstance(item, dict) and \
                vb_str(item.get("id"), "").lower() == lesson_id.lower():
            return True
    return False


class LessonNotFound(LessonError):
    """"not_found" keeps the VB receipt keys (LessonError only has code/message)."""

    def __init__(self, payload):
        LessonError.__init__(self, payload["error"], payload["message"])
        self.payload = payload


def delete_json(data_dir, lesson_id):
    if not lesson_id:
        raise LessonError("lesson_id_required", "id is required")
    doc, why, backup_path = load_user_doc_strict(data_dir)
    if doc is None:
        raise store_unreadable_error(data_dir, why, backup_path)
    lessons = _user_lessons(doc)
    removed = False
    for i in range(len(lessons) - 1, -1, -1):
        item = lessons[i]
        if isinstance(item, dict) and \
                vb_str(item.get("id"), "").lower() == lesson_id.lower():
            del lessons[i]
            removed = True
    if in_defaults(data_dir, lesson_id):
        lessons.append({"id": lesson_id, "enabled": False})
        action = "disabled(default)"
    elif removed:
        action = "removed"
    else:
        action = "not_found"
    backup(file_path(data_dir, USER_LEAF))
    if not write_text(file_path(data_dir, USER_LEAF), json_dumps(doc)):
        raise LessonError("lesson_write_failed", "cannot write user.json")
    result = {"ok": action != "not_found", "id": lesson_id, "action": action,
              "lessonsTotal": len(merged_lessons(data_dir, False))}
    if action == "not_found":
        # Lessons.DeleteJson returned ok:false here (with no error field).  The CLI
        # contract additionally requires error+message on a failure path, so the VB
        # keys (id/action/lessonsTotal) are kept AND the failure envelope is filled
        # in; the exit code is non-zero.
        result["error"] = "lesson_not_found"
        result["message"] = ("no lesson with id %r exists in the user layer or in "
                             "defaults; nothing was removed" % lesson_id)
        raise LessonNotFound(result)
    return result


# ---------------------------------------------------------------------------
# export (ExportMd)
# ---------------------------------------------------------------------------

def export_md(data_dir, include_disabled):
    col = merged_lessons(data_dir, include_disabled)
    s = EXPORT_TITLE + EOL
    for item in col:
        s += EOL + "## [" + vb_str(item.get("id"), "") + "] " + \
             vb_str(item.get("scenario"), "") + EOL
        s += MARK_TRAP + ": " + vb_str(item.get("trap"), "") + EOL
        s += MARK_DO + ": " + vb_str(item.get("do"), "") + EOL
        s += "scope: " + vb_str(item.get("scope"), "all") + EOL
        if not vb_bool(item.get("enabled"), True):
            s += "enabled: false" + EOL
    return s


# ---------------------------------------------------------------------------
# import (SplitMdHead / ValidLessonId / ParseMd / ImportMdJson)
# ---------------------------------------------------------------------------

def valid_lesson_id(s):
    """Lessons.ValidLessonId: 1..120 chars, no control chars, no [ or ]."""
    if not s or len(s) > 120:
        return False
    for ch in s:
        c = ord(ch)
        if 1 <= c <= 32:
            return False
        if c == 91 or c == 93:      # [ ]
            return False
    return True


def split_md_head(rest):
    """'## [id] scenario' -> (id, scenario); malformed heads yield no id."""
    id_out = ""
    scen_out = rest
    if not rest.startswith("["):
        return id_out, scen_out
    q = rest.find("]")
    if q < 1:                        # VB: InStr(rest,"]") < 2 (0 or 1)
        return id_out, scen_out
    cand = rest[1:q].strip()
    if valid_lesson_id(cand):
        id_out = cand
    scen_out = rest[q + 1:].strip()
    if not scen_out:
        scen_out = id_out if id_out else rest
    return id_out, scen_out


def parse_md(text):
    """Lessons.ParseMd -> list of [id, scenario, trap, do] (in file order)."""
    out = []
    cur = -1
    for raw in text.replace("\r\n", "\n").split("\n"):
        ln = raw.strip()
        if ln[:2] == "##":
            rest = ln[2:].strip()
            hid, hscen = split_md_head(rest)
            out.append([hid, hscen, "", ""])
            cur = len(out) - 1
        elif cur >= 0:
            # trap is tested first, exactly like Lessons.ParseMd.
            if ln[:2] == MARK_TRAP + ":" or ln[:4].lower() == "trap":
                out[cur][2] = ln[ln.find(":") + 1:].strip()
            elif ln[:2] == MARK_DO + ":" or ln[:2].lower() == "do":
                out[cur][3] = ln[ln.find(":") + 1:].strip()
    return out


def import_md_json(data_dir, path, dry_run):
    text = read_text(path)
    if not text:
        raise LessonError("lesson_import_read_failed", "cannot read: " + vb_str(path, ""))
    rows = parse_md(text)
    parsed = len(rows)
    if parsed == 0:
        raise LessonError("lesson_import_parsed_none", IMPORT_FORMAT_HINT)
    seen = merged_id_set(data_dir)
    added = 0
    skipped = 0
    conflicts = 0
    ids_used = 0
    for hid, scen, trap, do_text in rows:
        trap = trap.strip()
        do_text = do_text.strip()
        if trap and do_text:
            want_id = hid
            use_id = ""
            if want_id:
                low = want_id.lower()
                if low in seen:
                    conflicts += 1
                else:
                    use_id = want_id
                    seen.add(low)
                    ids_used += 1
            if not dry_run:
                try:
                    add_json(data_dir, use_id, scen, trap, do_text, "all", True)
                except LessonNotFound:
                    raise
                except LessonError as exc:
                    if exc.code == "lesson_store_unreadable":
                        raise            # Lessons.ImportMdJson returns this as-is
                    # Lessons counted every other AddJson failure as "added" (only
                    # the unreadable store was special-cased); kept for receipt
                    # fidelity -- see the module docstring.
                    added += 1
                    continue
                added += 1
        else:
            skipped += 1
    return {"ok": True, "parsed": parsed, "added": added, "skipped": skipped,
            "idsUsed": ids_used, "conflicts": conflicts, "dryRun": bool(dry_run)}


# ---------------------------------------------------------------------------
# validate
# ---------------------------------------------------------------------------

def _validate_file(data_dir, leaf, errors):
    path = file_path(data_dir, leaf)
    if not os.path.exists(path):
        if leaf == DEFAULTS_LEAF:
            errors.append("%s: file is missing (%s)" % (leaf, path))
        return 0
    text = read_text(path)
    if text.strip() == "":
        errors.append("%s: file exists but is EMPTY (%s); writers refuse to "
                      "replace it" % (leaf, path))
        return 0
    try:
        doc = json.loads(text)
    except ValueError as exc:
        errors.append("%s: not valid JSON (%s)" % (leaf, exc))
        return 0
    if not isinstance(doc, dict):
        errors.append("%s: top level is not a JSON object" % leaf)
        return 0
    arr = doc.get("lessons")
    if not isinstance(arr, list):
        errors.append("%s: 'lessons' is missing or not an array" % leaf)
        return 0
    seen = set()
    for i, item in enumerate(arr):
        where = "%s[%d]" % (leaf, i)
        if not isinstance(item, dict):
            errors.append("%s: entry is not an object (skipped by the merge)" % where)
            continue
        lid = vb_str(item.get("id"), "")
        if not lid:
            errors.append('%s: no usable "id" (entry is skipped by the merge)' % where)
            continue
        if not valid_lesson_id(lid):
            errors.append("%s: id %r is not a valid lesson id (1..120 chars, no "
                          "control/[/] characters)" % (where, lid))
        key = lid.lower()
        if key in seen:
            errors.append("%s: duplicate id %r in the same file (the later entry "
                          "merges into the earlier one)" % (where, lid))
        seen.add(key)
        for field in ("scenario", "trap", "do", "scope"):
            if field in item and item[field] is not None and \
                    not isinstance(item[field], str):
                errors.append('%s: "%s" is not a string' % (where, field))
        if "enabled" in item and item["enabled"] is not None and \
                not isinstance(item["enabled"], bool):
            errors.append('%s: "enabled" is not a boolean' % where)
    return len(arr)


def validate(data_dir):
    """Report the schema state of both stores.

    ok:true means the scan COMPLETED; errors[] is the verdict (a corrupt store
    still returns ok:true so the caller can print every problem at once).
    """
    errors = []
    n_defaults = _validate_file(data_dir, DEFAULTS_LEAF, errors)
    n_user = _validate_file(data_dir, USER_LEAF, errors)
    return {"ok": True, "defaults": n_defaults, "user": n_user, "errors": errors}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def emit(payload, exit_code=0):
    """Exactly one compact UTF-8 JSON object + LF on stdout, no BOM."""
    body = json_dumps(payload).encode("utf-8") + b"\n"
    stream = getattr(sys.stdout, "buffer", None)
    if stream is None:                       # exotic/embedded stdout
        sys.stdout.write(body.decode("utf-8"))
        sys.stdout.flush()
        return exit_code
    stream.write(body)
    stream.flush()
    return exit_code


def fail(code, message, exit_code=1, **extra):
    payload = {"ok": False, "error": code, "message": message}
    for key, value in extra.items():
        payload[key] = value
    return emit(payload, exit_code)


class Parser(argparse.ArgumentParser):
    """argparse that never prints chatter to stdout: usage errors are JSON."""

    def error(self, message):
        emit({"ok": False, "error": "usage", "message": message}, 2)
        raise SystemExit(2)


def bool_value(text):
    if isinstance(text, bool):
        return text
    return str(text).strip().lower() in ("1", "true", "yes", "on", "y")


def add_bool_flag(parser, name, help_text):
    parser.add_argument(name, dest=name.strip("-").replace("-", "_"),
                        nargs="?", const=True, default=False, metavar="0|1",
                        help=help_text)


def build_parser():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--data-dir", dest="data_dir", default=None, metavar="DIR",
                        help="read/write this directory instead of <repo>/%s "
                             "(for tests on copies)" % DIR_NAME)
    parser = Parser(
        prog="lessons.py",
        parents=[common],
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Lessons (experience library): authoritative Python port of "
                    "the VB add-in module Lessons.bas.",
        epilog='Every command prints exactly ONE compact JSON object (UTF-8, no '
               "BOM, trailing newline) and nothing else on stdout.  Failures are "
               '{"ok":false,"error":"<code>","message":"..."} with a non-zero exit '
               "code.",
    )
    sub = parser.add_subparsers(dest="command", metavar="<command>")

    p = sub.add_parser("instructions", parents=[common],
                       help="the initialize.instructions lesson section "
                            "(intro + numbered rules)")
    p.set_defaults(func=cmd_instructions)

    p = sub.add_parser("list", parents=[common],
                       help="list the merged lessons (defaults + user)")
    add_bool_flag(p, "--include-disabled", "also list disabled lessons (default off)")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("add", parents=[common],
                       help="add or overwrite a lesson in the user layer")
    p.add_argument("--id", dest="lesson_id", default="", metavar="X",
                   help="stable lesson id (optional; a free custom-N is minted)")
    p.add_argument("--scenario", dest="scenario", default="", metavar="S",
                   help="when this applies (required)")
    p.add_argument("--trap", dest="trap", default="", metavar="T",
                   help="the pitfall (required)")
    p.add_argument("--do", dest="do_text", default="", metavar="D",
                   help="what to do instead (required)")
    p.add_argument("--scope", dest="scope", default="all", metavar="mcp|vb6|all",
                   help="all | mcp | vb6 (default all)")
    p.add_argument("--enabled", dest="enabled", default="1", metavar="0|1",
                   help="1 = enabled (default), 0 = disabled")
    p.set_defaults(func=cmd_add)

    p = sub.add_parser("delete", parents=[common],
                       help="remove a user lesson, or disable a shipped one")
    p.add_argument("--id", dest="lesson_id", default="", metavar="X",
                   help="lesson id (required)")
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser("export", parents=[common],
                       help="export the merged library as re-importable markdown")
    add_bool_flag(p, "--include-disabled", "also export disabled lessons (default off)")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("import-md", parents=[common],
                       help="import lessons from a markdown file")
    p.add_argument("--path", dest="path", default="", metavar="P",
                   help="markdown file ('## [id] scenario' + 坑:/trap: + 应:/do:)")
    add_bool_flag(p, "--dry-run", "parse only, write nothing (default off)")
    p.set_defaults(func=cmd_import_md)

    p = sub.add_parser("validate", parents=[common],
                       help="check both stores against the schema")
    p.set_defaults(func=cmd_validate)
    return parser


def cmd_instructions(args, data_dir):
    return {"ok": True, "text": instructions_text(data_dir)}


def cmd_list(args, data_dir):
    return list_json(data_dir, bool_value(args.include_disabled))


def cmd_add(args, data_dir):
    return add_json(data_dir, args.lesson_id, args.scenario, args.trap,
                    args.do_text, args.scope, bool_value(args.enabled))


def cmd_delete(args, data_dir):
    return delete_json(data_dir, args.lesson_id)


def cmd_export(args, data_dir):
    return {"ok": True,
            "markdown": export_md(data_dir, bool_value(args.include_disabled))}


def cmd_import_md(args, data_dir):
    return import_md_json(data_dir, args.path, bool_value(args.dry_run))


def cmd_validate(args, data_dir):
    return validate(data_dir)


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "command", None):
        return fail("usage",
                    "no command given (expected one of: instructions, list, add, "
                    "delete, export, import-md, validate); try --help",
                    2)
    data_dir = resolve_data_dir(getattr(args, "data_dir", None))
    try:
        payload = args.func(args, data_dir)
    except LessonNotFound as exc:
        return emit(exc.payload, 1)
    except LessonError as exc:
        return fail(exc.code, exc.message, 1, **exc.extra)
    except SystemExit:
        raise
    except Exception as exc:            # never leak a traceback as stdout
        return fail("internal_error", "%s: %s" % (type(exc).__name__, exc), 3)
    return emit(payload, 0 if payload.get("ok") else 1)


if __name__ == "__main__":
    raise SystemExit(main())
