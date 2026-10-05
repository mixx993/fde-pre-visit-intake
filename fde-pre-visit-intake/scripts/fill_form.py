#!/usr/bin/env python3
"""Fill the pre-visit form template from answers.json. Standard library only.

Usage: python3 fill_form.py --answers answers.json --out filled.xlsx
"""
import argparse, json, os, re, sys, zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(HERE, "..", "template")


def flatten(obj, prefix=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from flatten(v, f"{prefix}{k}.")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from flatten(v, f"{prefix}{i}.")
    elif obj is not None and str(obj).strip() != "":
        yield prefix[:-1], str(obj).strip()


def set_cell(xml, cell, text):
    pattern = re.compile(r'<c r="%s"((?: s="\d+")?)[^>]*?(?:/>|>.*?</c>)' % cell, re.S)
    value = '<c r="%s"\\1 t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
        cell, escape(text).replace("\\", "\\\\"))
    new, n = pattern.subn(value, xml, count=1)
    return new, n == 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--answers", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    meta = json.load(open(os.path.join(TEMPLATE_DIR, "fields.json"), encoding="utf-8"))
    where = {f["id"]: (f["sheet_index"], f["cell"]) for f in meta["fields"]}
    answers = json.load(open(args.answers, encoding="utf-8"))

    edits, unknown = {}, []
    for key, val in flatten(answers):
        if key.startswith("_"):
            continue
        if key in where:
            sheet, cell = where[key]
            edits.setdefault(sheet, []).append((cell, val))
        else:
            unknown.append(key)

    src = zipfile.ZipFile(os.path.join(TEMPLATE_DIR, meta["template"]))
    out = zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED)
    filled, missed = 0, []
    for item in src.infolist():
        data = src.read(item.filename)
        m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
        if m and int(m.group(1)) in edits:
            xml = data.decode("utf-8")
            for cell, val in edits[int(m.group(1))]:
                xml, ok = set_cell(xml, cell, val)
                filled += ok
                if not ok:
                    missed.append(cell)
            data = xml.encode("utf-8")
        out.writestr(item, data)
    out.close()

    print(json.dumps({"out": os.path.abspath(args.out), "filled": filled,
                      "unknown_keys": unknown, "cells_not_found": missed}, ensure_ascii=False))
    return 1 if missed else 0


if __name__ == "__main__":
    sys.exit(main())
