"""
DSM Report Comparator - HTML Edition
Version 2.0

Input: 2 Generated DSM Reports (.txt)
Output: Standalone interactive HTML report

Requirements:
    Python 3.9+
    No external Python packages required.

Run:
    python DSM_Comparator_HTML.py
or double-click:
    DSM_Comparator_HTML.pyw
"""

import os
import re
import sys
import json
import html
import traceback
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Excel output is optional. HTML generation does not require openpyxl.
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


APP_NAME = "DSM Report Comparator"
VERSION = "2.0"


# ============================================================
# DSM PARSER
# ============================================================

def read_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def find_marker_ci(text, marker, start=0):
    """Find a DSM marker case-insensitively."""
    m = re.search(re.escape(marker), text[start:], re.IGNORECASE)
    return (start + m.start()) if m else -1


def get_section(text, start_marker, end_marker=None):
    start = find_marker_ci(text, start_marker)
    if start < 0:
        raise ValueError(f"Could not find required DSM section: {start_marker}")
    end = len(text)
    if end_marker:
        pos = find_marker_ci(text, end_marker, start + len(start_marker))
        if pos >= 0:
            end = pos
    return text[start:end]


def parse_field(block, label):
    """Return a DSM field value from the SAME physical line only.

    DSM reports are line-oriented.  Do not use ``\\s`` here because it
    can consume newlines when a field is empty.  Splitting into physical
    lines makes an empty field unambiguous and prevents the next field from
    being accidentally captured.
    """
    prefix = label + ":"
    for line in block.splitlines():
        stripped = line.strip()
        if stripped.startswith(prefix):
            return stripped[len(prefix):].strip()
    return ""


def parse_blocks(section_text, prefix_regex, flags=0):
    lines = section_text.splitlines()
    rx = re.compile(prefix_regex, flags)
    starts = [i for i, line in enumerate(lines) if rx.match(line.strip())]

    result = []
    for idx, start in enumerate(starts):
        end = starts[idx + 1] if idx + 1 < len(starts) else len(lines)
        result.append((lines[start].strip(), "\n".join(lines[start:end])))
    return result


def description_or_na(value):
    """Normalize DSM descriptions without consuming neighboring fields.

    Some DSM exports append the metadata token ``LongName_EN`` directly to
    the end of the description, e.g. ``...catalyst heatingLongName_EN``.
    That token is export metadata, not part of the description.
    """
    value = (value or "").strip()

    # Remove the generated metadata marker wherever it is appended to the
    # description.  Keep the actual description text intact.
    value = re.sub(r"LongName_EN\s*$", "", value, flags=re.IGNORECASE).strip()

    if not value or value.lower() in {"null", "none", "n/a", "na"}:
        return "Not Available"
    return value


def parse_dfc(text):
    sec = get_section(text, "DFC_Name", "DSQ_Name")
    result = {}

    for name, block in parse_blocks(sec, r"^DFC_\S+\s*$", flags=re.IGNORECASE):
        if name.lower() in {"dfc_name", "dfc_number"}:
            continue
        result[name] = {
            "number": parse_field(block, "Element Number"),
            "mp": parse_field(block, "Collection MP"),
            "elements": parse_field(block, "Elements"),
            "method": parse_field(block, "Debounce Method"),
            "param": parse_field(block, "Debounce Parameter"),
            "description": description_or_na(parse_field(block, "Description")),
            "readiness": parse_field(block, "Readiness Group assigned"),
            "multiple": parse_field(
                block,
                "Multiple Debounce (Supported - 1, Not Supported - 0)"
            ),
        }
    return result


def parse_dsq(text):
    sec = get_section(text, "DSQ_Name", "DFP_Name")
    result = {}

    for name, block in parse_blocks(sec, r"^(?:DSQ|DFQ)_\S+\s*$", flags=re.IGNORECASE):
        if name.lower() in {"dsq_name", "dsq_number", "dfq_name", "dfq_number"}:
            continue
        result[name] = {
            "number": parse_field(block, "Element Number"),
            "mp": parse_field(block, "Collection MP"),
            "elements": parse_field(block, "Elements"),
            "description": description_or_na(parse_field(block, "Description")),
        }
    return result


def parse_fid(text):
    start = find_marker_ci(text, "FId_Name")
    if start < 0:
        raise ValueError("Could not find required DSM section: FId_Name")

    end_match = re.search(
        r"^(?:DSQ|DFQ)_\S+\s+Assigned inhibits:\s*$",
        text[start:],
        re.MULTILINE | re.IGNORECASE
    )
    sec = text[start:start + end_match.start()] if end_match else text[start:]

    result = {}
    for name, block in parse_blocks(sec, r"^FID_\S+\s*$", flags=re.IGNORECASE):
        if name.lower() in {"fid_name", "fid_number"}:
            continue

        mappings = []
        for line in block.splitlines():
            m = re.match(
                r"^\s+(DFC_\S+)\s+\(([^)]*)\)\s*\{([^}]*)\}",
                line, re.IGNORECASE
            )
            if m:
                mappings.append({
                    "dfc": m.group(1),
                    "debounce": m.group(2).strip(),
                    "category": m.group(3).strip(),
                })

        result[name] = {
            "number": parse_field(block, "FId Number"),
            "mp": parse_field(block, "Collection MP"),
            "description": description_or_na(parse_field(block, "Description")),
            "iumpr": parse_field(block, "Assigned IUMPR Group"),
            "scheduling": parse_field(block, "Scheduling Mode"),
            "mappings": mappings,
        }
    return result


def parse_report(path):
    text = read_text(path)

    if not re.search(r"Overview over generated elements for DSM", text, re.IGNORECASE):
        raise ValueError(
            f"{os.path.basename(path)} does not appear to be a generated DSM report."
        )

    m = re.search(r"DSM_xTabVersion:\s*(\d+)", text)

    return {
        "filename": os.path.basename(path),
        "table_version": m.group(1) if m else "",
        "DFC": parse_dfc(text),
        "FID": parse_fid(text),
        "DSQ": parse_dsq(text),
    }


# ============================================================
# COMPARISON
# ============================================================

def compare_entities(old, new, fields):
    added = sorted(set(new) - set(old))
    removed = sorted(set(old) - set(new))
    modified = []

    for name in sorted(set(old) & set(new)):
        changes = [f for f in fields if old[name].get(f, "") != new[name].get(f, "")]
        if changes:
            modified.append((name, changes))

    return added, removed, modified


def mapping_set(report):
    return {
        (fid, item["dfc"])
        for fid, data in report["FID"].items()
        for item in data["mappings"]
    }


def compare_reports(old, new):
    da, dr, dm = compare_entities(
        old["DFC"], new["DFC"],
        ["description", "method", "param", "readiness", "multiple"]
    )
    fa, fr, fm = compare_entities(
        old["FID"], new["FID"],
        ["description", "iumpr", "scheduling"]
    )

    # A FID is also functionally modified when its assigned DFC set changes.
    old_fid_map = {fid: {item["dfc"] for item in data["mappings"]} for fid, data in old["FID"].items()}
    new_fid_map = {fid: {item["dfc"] for item in data["mappings"]} for fid, data in new["FID"].items()}
    fm_dict = {name: list(changes) for name, changes in fm}
    for fid in sorted(set(old["FID"]) & set(new["FID"])):
        if old_fid_map.get(fid, set()) != new_fid_map.get(fid, set()):
            if "mappings" not in fm_dict.setdefault(fid, []):
                fm_dict[fid].append("mappings")
    fm = sorted((name, changes) for name, changes in fm_dict.items())
    qa, qr, qm = compare_entities(
        old["DSQ"], new["DSQ"],
        ["description"]
    )

    om = mapping_set(old)
    nm = mapping_set(new)

    return {
        "dfc_added": da, "dfc_removed": dr, "dfc_modified": dm,
        "fid_added": fa, "fid_removed": fr, "fid_modified": fm,
        "dsq_added": qa, "dsq_removed": qr, "dsq_modified": qm,
        "old_map": om, "new_map": nm,
        "map_added": sorted(nm - om),
        "map_removed": sorted(om - nm),
    }


# ============================================================
# HTML GENERATOR
# ============================================================

def esc(value):
    return html.escape(str(value or ""))


def js_data(value):
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")


def badge(status):
    return f'<span class="badge {status.lower()}">{esc(status)}</span>'


def table_empty(message, colspan):
    return f'<tr><td colspan="{colspan}" class="empty">{esc(message)}</td></tr>'


def make_dfc_rows(old, new, cmp):
    rows = []

    for name in cmp["dfc_added"]:
        n = new["DFC"][name]
        rows.append({
            "status": "ADDED", "name": name,
            "description": n["description"],
            "old_method": "", "new_method": n["method"],
            "old_param": "", "new_param": n["param"],
            "old_readiness": "", "new_readiness": n["readiness"],
            "old_multiple": "", "new_multiple": n["multiple"],
            "changed": ["New DFC"],
            "structural": False
        })

    for name in cmp["dfc_removed"]:
        o = old["DFC"][name]
        rows.append({
            "status": "REMOVED", "name": name,
            "description": o["description"],
            "old_method": o["method"], "new_method": "",
            "old_param": o["param"], "new_param": "",
            "old_readiness": o["readiness"], "new_readiness": "",
            "old_multiple": o["multiple"], "new_multiple": "",
            "changed": ["DFC removed"],
            "structural": False
        })

    for name, changes in cmp["dfc_modified"]:
        o, n = old["DFC"][name], new["DFC"][name]
        structural = (o["number"], o["mp"]) != (n["number"], n["mp"])
        rows.append({
            "status": "MODIFIED", "name": name,
            "description": n["description"],
            "old_method": o["method"], "new_method": n["method"],
            "old_param": o["param"], "new_param": n["param"],
            "old_readiness": o["readiness"], "new_readiness": n["readiness"],
            "old_multiple": o["multiple"], "new_multiple": n["multiple"],
            "changed": changes,
            "structural": structural
        })

    return rows


def make_fid_rows(old, new, cmp):
    rows = []

    for name in cmp["fid_added"]:
        n = new["FID"][name]
        rows.append({
            "status": "ADDED", "name": name, "description": n["description"],
            "old_iumpr": "", "new_iumpr": n["iumpr"],
            "old_sched": "", "new_sched": n["scheduling"],
            "changed": ["New FID"]
        })

    for name in cmp["fid_removed"]:
        o = old["FID"][name]
        rows.append({
            "status": "REMOVED", "name": name, "description": o["description"],
            "old_iumpr": o["iumpr"], "new_iumpr": "",
            "old_sched": o["scheduling"], "new_sched": "",
            "changed": ["FID removed"]
        })

    for name, changes in cmp["fid_modified"]:
        o, n = old["FID"][name], new["FID"][name]
        rows.append({
            "status": "MODIFIED", "name": name, "description": n["description"],
            "old_iumpr": o["iumpr"], "new_iumpr": n["iumpr"],
            "old_sched": o["scheduling"], "new_sched": n["scheduling"],
            "changed": changes
        })

    return rows


def make_dsq_rows(old, new, cmp):
    rows = []

    for name in cmp["dsq_added"]:
        n = new["DSQ"][name]
        rows.append({
            "status": "ADDED", "name": name,
            "description": n["description"],
            "old_number": "", "new_number": n["number"],
            "changed": ["New DSQ"]
        })

    for name in cmp["dsq_removed"]:
        o = old["DSQ"][name]
        rows.append({
            "status": "REMOVED", "name": name,
            "description": o["description"],
            "old_number": o["number"], "new_number": "",
            "changed": ["DSQ removed"]
        })

    for name, changes in cmp["dsq_modified"]:
        o, n = old["DSQ"][name], new["DSQ"][name]
        rows.append({
            "status": "MODIFIED", "name": name,
            "description": n["description"],
            "old_number": o["number"], "new_number": n["number"],
            "changed": changes
        })

    return rows


def build_html(old, new, cmp, output_path):
    dfc_rows = make_dfc_rows(old, new, cmp)
    fid_rows = make_fid_rows(old, new, cmp)
    dsq_rows = make_dsq_rows(old, new, cmp)

    # Complete DFC debounce list
    debounce_rows = []
    for name in sorted(set(old["DFC"]) | set(new["DFC"])):
        o, n = old["DFC"].get(name), new["DFC"].get(name)

        if o and n:
            functional = any(
                o[k] != n[k]
                for k in ["description", "method", "param", "readiness", "multiple"]
            )
            structural = (o["number"], o["mp"]) != (n["number"], n["mp"])
            status = "MODIFIED" if functional else "UNCHANGED"
            debounce_rows.append({
                "name": name,
                "description": n["description"],
                "old_method": o["method"], "new_method": n["method"],
                "old_param": o["param"], "new_param": n["param"],
                "old_readiness": o["readiness"], "new_readiness": n["readiness"],
                "structural": structural,
                "status": status
            })
        elif n:
            debounce_rows.append({
                "name": name, "description": n["description"],
                "old_method": "", "new_method": n["method"],
                "old_param": "", "new_param": n["param"],
                "old_readiness": "", "new_readiness": n["readiness"],
                "structural": True, "status": "ADDED"
            })
        else:
            debounce_rows.append({
                "name": name, "description": o["description"],
                "old_method": o["method"], "new_method": "",
                "old_param": o["param"], "new_param": "",
                "old_readiness": o["readiness"], "new_readiness": "",
                "structural": True, "status": "REMOVED"
            })

    old_dfcs = set(old["DFC"])
    new_dfcs = set(new["DFC"])
    old_refs = {d for _, d in cmp["old_map"]}
    new_refs = {d for _, d in cmp["new_map"]}

    validation = [
        ("Old DFCs not referenced by any FID", "INFO",
         len(old_dfcs - old_refs), sorted(old_dfcs - old_refs)),
        ("New DFCs not referenced by any FID", "INFO",
         len(new_dfcs - new_refs), sorted(new_dfcs - new_refs)),
        ("Old FIDs with no DFC mapping", "INFO",
         sum(not x["mappings"] for x in old["FID"].values()),
         sorted(k for k, x in old["FID"].items() if not x["mappings"])),
        ("New FIDs with no DFC mapping", "INFO",
         sum(not x["mappings"] for x in new["FID"].values()),
         sorted(k for k, x in new["FID"].items() if not x["mappings"])),
    ]

    # Mapping rows
    old_lookup = {
        (fid, item["dfc"]): item
        for fid, data in old["FID"].items()
        for item in data["mappings"]
    }
    new_lookup = {
        (fid, item["dfc"]): item
        for fid, data in new["FID"].items()
        for item in data["mappings"]
    }

    mapping_rows = []
    for fid, dfc in cmp["map_added"]:
        n = new_lookup[(fid, dfc)]
        mapping_rows.append({
            "status": "ADDED", "fid": fid, "dfc": dfc,
            "old_info": "", "new_info": f'{n["debounce"]} / {n["category"]}'
        })
    for fid, dfc in cmp["map_removed"]:
        o = old_lookup[(fid, dfc)]
        mapping_rows.append({
            "status": "REMOVED", "fid": fid, "dfc": dfc,
            "old_info": f'{o["debounce"]} / {o["category"]}', "new_info": ""
        })

    # Dependency / impact analysis for modified DFCs
    impacted = {}
    for row in dfc_rows:
        if row["status"] == "MODIFIED":
            refs = []
            for fid, data in new["FID"].items():
                if row["name"] in {x["dfc"] for x in data["mappings"]}:
                    refs.append(fid)
            impacted[row["name"]] = refs

    summary = {
        "dfc_old": len(old["DFC"]), "dfc_new": len(new["DFC"]),
        "dfc_added": len(cmp["dfc_added"]), "dfc_removed": len(cmp["dfc_removed"]),
        "dfc_modified": len(cmp["dfc_modified"]),
        "fid_old": len(old["FID"]), "fid_new": len(new["FID"]),
        "fid_added": len(cmp["fid_added"]), "fid_removed": len(cmp["fid_removed"]),
        "fid_modified": len(cmp["fid_modified"]),
        "dsq_old": len(old["DSQ"]), "dsq_new": len(new["DSQ"]),
        "dsq_added": len(cmp["dsq_added"]), "dsq_removed": len(cmp["dsq_removed"]),
        "dsq_modified": len(cmp["dsq_modified"]),
        "map_old": len(cmp["old_map"]), "map_new": len(cmp["new_map"]),
        "map_added": len(cmp["map_added"]), "map_removed": len(cmp["map_removed"]),
    }

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------
    css = r"""
:root{
  --bg:#f5f7fb; --card:#ffffff; --text:#172033; --muted:#697386;
  --line:#e4e8ef; --blue:#2563eb; --blue2:#dbeafe;
  --green:#16803c; --greenbg:#dcfce7; --red:#b42318; --redbg:#fee2e2;
  --amber:#a15c00; --amberbg:#fef3c7; --graybg:#f1f5f9;
  --shadow:0 8px 28px rgba(15,23,42,.07);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);font-family:Segoe UI,Arial,sans-serif;color:var(--text)}
.container{max-width:1500px;margin:auto;padding:28px}
.hero{background:linear-gradient(135deg,#17365d,#2563eb);color:#fff;border-radius:20px;padding:28px 32px;box-shadow:var(--shadow)}
.hero h1{margin:0 0 8px;font-size:28px}
.hero .sub{opacity:.88;font-size:14px}
.fileline{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}
.filepill{background:rgba(255,255,255,.13);padding:9px 12px;border-radius:10px;font-size:13px}
.tabs{position:sticky;top:0;z-index:20;background:rgba(245,247,251,.95);backdrop-filter:blur(10px);display:flex;gap:4px;padding:14px 0}
.tab{border:0;background:transparent;padding:11px 15px;border-radius:10px;cursor:pointer;font-weight:600;color:var(--muted)}
.tab:hover,.tab.active{background:#fff;color:var(--blue);box-shadow:var(--shadow)}
.panel{display:none}.panel.active{display:block}
.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px;box-shadow:var(--shadow)}
.card .label{font-size:12px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.04em}
.card .count{font-size:26px;font-weight:800;margin:7px 0}
.delta{font-size:13px}
.add{color:var(--green)}.remove{color:var(--red)}.modify{color:var(--amber)}
.section{background:#fff;border:1px solid var(--line);border-radius:16px;margin-top:18px;padding:20px;box-shadow:var(--shadow)}
.section h2{margin:0 0 14px;font-size:18px}
.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0}
input,select{border:1px solid #cbd5e1;background:#fff;border-radius:9px;padding:10px 12px;font-size:13px}
input{min-width:280px;flex:1}
button.action{border:1px solid #cbd5e1;background:#fff;border-radius:9px;padding:10px 13px;cursor:pointer}
button.action:hover{border-color:var(--blue);color:var(--blue)}
.tablewrap{overflow:auto;border:1px solid var(--line);border-radius:12px}
table{border-collapse:collapse;width:100%;font-size:13px;background:#fff}
th{position:sticky;top:0;background:#f8fafc;text-align:left;color:#475569;font-size:11px;text-transform:uppercase;letter-spacing:.04em;padding:11px;border-bottom:1px solid var(--line);white-space:nowrap}
td{padding:11px;border-bottom:1px solid #edf0f4;vertical-align:top}
tr:hover td{background:#f8fbff}
.badge{display:inline-block;border-radius:999px;padding:4px 8px;font-size:10px;font-weight:800}
.badge.added{background:var(--greenbg);color:var(--green)}
.badge.removed{background:var(--redbg);color:var(--red)}
.badge.modified{background:var(--amberbg);color:var(--amber)}
.badge.unchanged{background:var(--graybg);color:#475569}
.change{background:#fff7d6;border-radius:5px;padding:3px 6px;margin:2px;display:inline-block}
.old{color:#9b2c2c;text-decoration:line-through}.new{color:#166534;font-weight:650}
.arrow{color:#94a3b8;padding:0 5px}
.mono{font-family:Consolas,monospace;font-size:12px}
.muted{color:var(--muted)}
.empty{text-align:center;color:var(--muted);padding:35px}
details{border:1px solid var(--line);border-radius:10px;margin:7px 0;padding:10px;background:#fff}
summary{cursor:pointer;font-weight:700}
.kpi{display:flex;gap:8px;align-items:baseline}
.kpi strong{font-size:22px}
.warning{background:#fff8e1;border-left:4px solid #eab308;padding:12px;border-radius:8px}
.footer{color:var(--muted);font-size:12px;padding:25px 0}
@media(max-width:1000px){.cards{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){.container{padding:12px}.cards{grid-template-columns:1fr}.hero h1{font-size:22px}}
"""

    # JS dataset includes all rows for client-side filtering.
    datasets = {
        "dfc": dfc_rows,
        "fid": fid_rows,
        "dsq": dsq_rows,
        "debounce": debounce_rows,
        "mapping": mapping_rows,
        "impact": impacted,
        "validation": validation,
    }

    html_parts = []
    html_parts.append("<!doctype html><html><head><meta charset='utf-8'>")
    html_parts.append(f"<title>DSM Comparison — {esc(old['filename'])} vs {esc(new['filename'])}</title>")
    html_parts.append("<style>" + css + "</style></head><body>")
    html_parts.append("<div class='container'>")

    html_parts.append("""
    <div class="hero">
      <h1>Diagnostics Comparison Report</h1>
      <div class="sub">Comparison of Diagnostic Changes Across PVER Releases</div>
      <div class="fileline">
    """)
    html_parts.append(f"<div class='filepill'>Baseline DSM Report: {esc(old['filename'])}</div>")
    html_parts.append(f"<div class='filepill'>Target DSM Report: {esc(new['filename'])}</div>")
    html_parts.append(f"<div class='filepill'>DSM Table · {esc(old['table_version'])} → {esc(new['table_version'])}</div>")
    html_parts.append("</div></div>")

    # Tabs
    tabs = [
        ("overview", "Overview"),
        ("dfc", "DFC"),
        ("fid", "FID"),
        ("dsq", "DSQ"),
        ("mapping", "FID → DFC"),
        ("debounce", "Debounce"),
        ("validation", "Validation"),
    ]
    html_parts.append("<nav class='tabs'>")
    for i, (key, label) in enumerate(tabs):
        html_parts.append(f"<button class='tab {'active' if i==0 else ''}' onclick=\"showTab('{key}',this)\">{label}</button>")
    html_parts.append("</nav>")

    # Overview
    html_parts.append("<section id='overview' class='panel active'>")
    html_parts.append("<div class='cards'>")

    def card(title, oldv, newv, added, removed, modified):
        return f"""
        <div class="card">
          <div class="label">{title}</div>
          <div class="count">{oldv:,} <span class="muted">→</span> {newv:,}</div>
          <div class="delta">
            <span class="add">+{added:,} added</span> ·
            <span class="remove">-{removed:,} removed</span>
            {f' · <span class="modify">~{modified:,} modified</span>' if modified is not None else ''}
          </div>
        </div>"""

    html_parts.append(card("DFC", summary["dfc_old"], summary["dfc_new"],
                           summary["dfc_added"], summary["dfc_removed"], summary["dfc_modified"]))
    html_parts.append(card("FID", summary["fid_old"], summary["fid_new"],
                           summary["fid_added"], summary["fid_removed"], summary["fid_modified"]))
    html_parts.append(card("DSQ", summary["dsq_old"], summary["dsq_new"],
                           summary["dsq_added"], summary["dsq_removed"], summary["dsq_modified"]))
    html_parts.append(card("FID → DFC Mappings", summary["map_old"], summary["map_new"],
                           summary["map_added"], summary["map_removed"], None))
    html_parts.append("</div>")

    html_parts.append("""
    <div class="section">
      <h2>Change Navigator</h2>
      <div class="toolbar">
        <button class="action" onclick="showTab('dfc',document.querySelectorAll('.tab')[1]); filterTable('dfc','MODIFIED')">16/changed DFCs</button>
        <button class="action" onclick="showTab('dfc',document.querySelectorAll('.tab')[1]); filterTable('dfc','ADDED')">Added DFCs</button>
        <button class="action" onclick="showTab('dfc',document.querySelectorAll('.tab')[1]); filterTable('dfc','REMOVED')">Removed DFCs</button>
        <button class="action" onclick="showTab('fid',document.querySelectorAll('.tab')[2]); filterTable('fid','MODIFIED')">Modified FIDs</button>
        <button class="action" onclick="showTab('mapping',document.querySelectorAll('.tab')[4]); filterTable('mapping','ADDED')">New mappings</button>
      </div>
      <div class="warning">
        <strong>Comparison rule:</strong> entity names are the primary identity.
        Generated Element Number / Collection MP shifts are treated as structural changes,
        not automatically as functional changes.
      </div>
    </div>
    """.replace("16/changed", f"{summary['dfc_modified']} changed"))

    html_parts.append("</section>")

    # Generic table helpers in HTML
    def toolbar(tabkey):
        return f"""
        <div class="toolbar">
          <input id="{tabkey}_search" placeholder="Search name, description..." oninput="renderTable('{tabkey}')">
          <select id="{tabkey}_status" onchange="renderTable('{tabkey}')">
            <option value="">All statuses</option>
            <option>ADDED</option><option>REMOVED</option><option>MODIFIED</option><option>UNCHANGED</option>
          </select>
          <button class="action" onclick="clearFilter('{tabkey}')">Clear</button>
        </div>"""

    html_parts.append("<section id='dfc' class='panel'><div class='section'><h2>DFC Delta</h2>")
    html_parts.append(toolbar("dfc"))
    html_parts.append("<div class='tablewrap'><table><thead><tr><th>Status</th><th>DFC</th><th>Description</th><th>Debounce Method</th><th>Debounce Parameter</th><th>Readiness</th><th>Changes</th><th>Impact</th></tr></thead><tbody id='dfc_body'></tbody></table></div></div></section>")

    html_parts.append("<section id='fid' class='panel'><div class='section'><h2>FID Delta</h2>")
    html_parts.append(toolbar("fid"))
    html_parts.append("<div class='tablewrap'><table><thead><tr><th>Status</th><th>FID</th><th>Description</th><th>Scheduling</th><th>IUMPR</th><th>Changes</th></tr></thead><tbody id='fid_body'></tbody></table></div></div></section>")

    html_parts.append("<section id='dsq' class='panel'><div class='section'><h2>DSQ Delta</h2>")
    html_parts.append(toolbar("dsq"))
    html_parts.append("<div class='tablewrap'><table><thead><tr><th>Status</th><th>DSQ</th><th>Description</th><th>Old Element</th><th>New Element</th><th>Changes</th></tr></thead><tbody id='dsq_body'></tbody></table></div></div></section>")

    html_parts.append("<section id='mapping' class='panel'><div class='section'><h2>FID → DFC Mapping Delta</h2>")
    html_parts.append(toolbar("mapping"))
    html_parts.append("<div class='tablewrap'><table><thead><tr><th>Status</th><th>FID</th><th>DFC</th><th>Old</th><th>New</th></tr></thead><tbody id='mapping_body'></tbody></table></div></div></section>")

    html_parts.append("<section id='debounce' class='panel'><div class='section'><h2>Complete DFC Debounce View</h2>")
    html_parts.append(toolbar("debounce"))
    html_parts.append("<div class='tablewrap'><table><thead><tr><th>Status</th><th>DFC</th><th>Description</th><th>Debounce Method</th><th>Debounce Parameter</th><th>Readiness</th><th>Structural Change</th></tr></thead><tbody id='debounce_body'></tbody></table></div></div></section>")

    html_parts.append("<section id='validation' class='panel'><div class='section'><h2>Validation / Review Checks</h2>")
    html_parts.append("<div id='validation_content'></div></div></section>")

    html_parts.append("<div class='footer'>Generated by DSM Report Comparator v2.0 · Built by MS/EPK-PS.</div>")
    html_parts.append("</div>")

    # JavaScript
    data_json = js_data(datasets)

    js = r"""
const DATA = __DATA__;

function showTab(id, btn){
  document.querySelectorAll('.panel').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  if(btn) btn.classList.add('active');
  if(id==='dfc') renderTable('dfc');
  if(id==='fid') renderTable('fid');
  if(id==='dsq') renderTable('dsq');
  if(id==='mapping') renderTable('mapping');
  if(id==='debounce') renderTable('debounce');
}

function badge(s){
  return '<span class="badge '+String(s).toLowerCase()+'">'+esc(s)+'</span>';
}
function esc(s){
  return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
}
function changeList(a){
  if(!a || !a.length) return '<span class="muted">—</span>';
  return a.map(x=>'<span class="change">'+esc(x)+'</span>').join(' ');
}
function arrow(o,n){
  if(o===n) return esc(n);
  return '<span class="old">'+esc(o||'—')+'</span><span class="arrow">→</span><span class="new">'+esc(n||'—')+'</span>';
}

function getRows(key){
  let rows=DATA[key]||[];
  const search=(document.getElementById(key+'_search')?.value||'').toLowerCase();
  const status=document.getElementById(key+'_status')?.value||'';
  if(search){
    rows=rows.filter(r=>JSON.stringify(r).toLowerCase().includes(search));
  }
  if(status) rows=rows.filter(r=>r.status===status);
  return rows;
}

function renderTable(key){
  const rows=getRows(key);
  const body=document.getElementById(key+'_body');
  if(!body) return;

  if(!rows.length){
    body.innerHTML='<tr><td colspan="8" class="empty">No matching records.</td></tr>';
    return;
  }

  if(key==='dfc'){
    body.innerHTML=rows.map(r=>`<tr>
      <td>${badge(r.status)}</td>
      <td class="mono"><strong>${esc(r.name)}</strong></td>
      <td>${esc(r.description)}</td>
      <td>${arrow(r.old_method,r.new_method)}</td>
      <td>${arrow(r.old_param,r.new_param)}</td>
      <td>${arrow(r.old_readiness,r.new_readiness)}</td>
      <td>${changeList(r.changed)}</td>
      <td>${r.status==='MODIFIED' && DATA.impact[r.name] ?
        '<details><summary>'+DATA.impact[r.name].length+' FID(s)</summary>'+
        DATA.impact[r.name].map(x=>'<div class="mono">'+esc(x)+'</div>').join('')+
        '</details>' : '<span class="muted">—</span>'}</td>
    </tr>`).join('');
  }

  if(key==='fid'){
    body.innerHTML=rows.map(r=>`<tr>
      <td>${badge(r.status)}</td><td class="mono"><strong>${esc(r.name)}</strong></td>
      <td>${esc(r.description)}</td>
      <td>${arrow(r.old_sched,r.new_sched)}</td>
      <td>${arrow(r.old_iumpr,r.new_iumpr)}</td>
      <td>${changeList(r.changed)}</td>
    </tr>`).join('');
  }

  if(key==='dsq'){
    body.innerHTML=rows.map(r=>`<tr>
      <td>${badge(r.status)}</td><td class="mono"><strong>${esc(r.name)}</strong></td>
      <td>${esc(r.description)}</td><td>${esc(r.old_number||'—')}</td>
      <td>${esc(r.new_number||'—')}</td><td>${changeList(r.changed)}</td>
    </tr>`).join('');
  }

  if(key==='mapping'){
    body.innerHTML=rows.map(r=>`<tr>
      <td>${badge(r.status)}</td><td class="mono">${esc(r.fid)}</td>
      <td class="mono"><strong>${esc(r.dfc)}</strong></td>
      <td>${esc(r.old_info||'—')}</td><td>${esc(r.new_info||'—')}</td>
    </tr>`).join('');
  }

  if(key==='debounce'){
    body.innerHTML=rows.map(r=>`<tr>
      <td>${badge(r.status)}</td><td class="mono"><strong>${esc(r.name)}</strong></td>
      <td>${esc(r.description)}</td>
      <td>${arrow(r.old_method,r.new_method)}</td>
      <td>${arrow(r.old_param,r.new_param)}</td>
      <td>${arrow(r.old_readiness,r.new_readiness)}</td>
      <td>${r.structural ? '<span class="change">YES</span>' : 'NO'}</td>
    </tr>`).join('');
  }
}

function clearFilter(key){
  document.getElementById(key+'_search').value='';
  document.getElementById(key+'_status').value='';
  renderTable(key);
}

function filterTable(key,status){
  const el=document.getElementById(key+'_status');
  if(el){el.value=status;renderTable(key);}
}

function renderValidation(){
  const root=document.getElementById('validation_content');
  root.innerHTML=DATA.validation.map(v=>`
    <details>
      <summary>${esc(v[0])} — ${v[2]}</summary>
      <div style="margin-top:10px">${v[3].length ? v[3].map(x=>'<div class="mono">'+esc(x)+'</div>').join('') : '<span class="muted">None</span>'}</div>
    </details>
  `).join('');
}

document.addEventListener('DOMContentLoaded',()=>{
  renderValidation();
  renderTable('dfc');
});
"""

    js = js.replace("__DATA__", data_json)
    html_parts.append("<script>" + js + "</script>")
    html_parts.append("</body></html>")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("".join(html_parts))


# ============================================================
# EXCEL GENERATOR (optional)
# ============================================================

# ---------- Excel generation ----------

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
TITLE_FILL = PatternFill("solid", fgColor="17365D")
SUB_FILL = PatternFill("solid", fgColor="D9EAF7")
ADDED_FILL = PatternFill("solid", fgColor="E2F0D9")
REMOVED_FILL = PatternFill("solid", fgColor="FCE4D6")
MODIFIED_FILL = PatternFill("solid", fgColor="FFF2CC")
INFO_FILL = PatternFill("solid", fgColor="DDEBF7")

WHITE_BOLD = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFF")
THIN = Side(style="thin", color="D9E1F2")


def style_header(ws, row, count):
    for c in range(1, count + 1):
        cell = ws.cell(row, c)
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )


def autosize(ws, max_width=60):
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        width = 10
        for cell in list(col_cells)[:800]:
            width = max(width, len(str(cell.value or "")) + 2)
        ws.column_dimensions[letter].width = min(width, max_width)


def add_table(ws, name):
    if ws.max_row < 2:
        return
    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)


def apply_common(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def status_fill(status):
    return {
        "ADDED": ADDED_FILL,
        "REMOVED": REMOVED_FILL,
        "MODIFIED": MODIFIED_FILL,
    }.get(status)


def build_excel(old, new, cmp, output_path):
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("Excel output requires the openpyxl package. Please run RUN_DSM_COMPARATOR.bat or install it with: python -m pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    ws.title = "01_Summary"

    # Summary
    ws.merge_cells("A1:F1")
    ws["A1"] = "DSM REPORT COMPARISON"
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")

    ws["A3"] = "Old Report"
    ws["B3"] = old["filename"]
    ws["A4"] = "New Report"
    ws["B4"] = new["filename"]
    ws["A5"] = "Old DSM Table Version"
    ws["B5"] = old["table_version"]
    ws["A6"] = "New DSM Table Version"
    ws["B6"] = new["table_version"]

    headers = ["Category", "Old Count", "New Count", "Added", "Removed", "Modified"]
    for c, value in enumerate(headers, 1):
        ws.cell(8, c, value)
    style_header(ws, 8, 6)

    summary_rows = [
        ("DFC", len(old["DFC"]), len(new["DFC"]),
         len(cmp["dfc_added"]), len(cmp["dfc_removed"]), len(cmp["dfc_modified"])),
        ("FID", len(old["FID"]), len(new["FID"]),
         len(cmp["fid_added"]), len(cmp["fid_removed"]), len(cmp["fid_modified"])),
        ("DSQ", len(old["DSQ"]), len(new["DSQ"]),
         len(cmp["dsq_added"]), len(cmp["dsq_removed"]), len(cmp["dsq_modified"])),
        ("FID → DFC Mapping", len(cmp["old_map"]), len(cmp["new_map"]),
         len(cmp["map_added"]), len(cmp["map_removed"]), 0),
    ]

    for r, row in enumerate(summary_rows, 9):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            ws.cell(r, c).border = Border(bottom=THIN)

    ws["A15"] = "Comparison logic"
    ws["A15"].font = Font(bold=True)
    ws["B15"] = (
        "Entity names are the primary comparison keys. Functional changes include "
        "description/debounce/readiness/scheduling changes. Generated Element Number "
        "and Collection MP shifts are reported separately because insertion/removal "
        "of entries can legitimately shift generated indices."
    )
    ws.merge_cells("B15:F17")
    ws["B15"].alignment = Alignment(wrap_text=True, vertical="top")

    ws["A19"] = "Reference"
    ws["B19"] = "DFC definitions include debounce method, debounce parameter and description."
    ws.merge_cells("B19:F19")

    apply_common(ws)
    autosize(ws, 70)

    # DFC delta
    ws = wb.create_sheet("02_DFC_Delta")
    cols = [
        "Status", "DFC", "Description",
        "Old Debounce Method", "New Debounce Method",
        "Old Debounce Parameter", "New Debounce Parameter",
        "Old Readiness Group", "New Readiness Group",
        "Old Multiple Debounce", "New Multiple Debounce",
        "Changed Fields"
    ]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    def add_dfc(status, name):
        source = new["DFC"] if status == "ADDED" else old["DFC"]
        rec = source[name]
        if status == "ADDED":
            row = [status, name, rec["description"], "",
                   rec["method"], "", rec["param"], "",
                   rec["readiness"], "", rec["multiple"], "ADDED"]
        elif status == "REMOVED":
            row = [status, name, rec["description"], rec["method"], "",
                   rec["param"], "", rec["readiness"], "", rec["multiple"], "", "REMOVED"]
        ws.append(row)

    for name in cmp["dfc_added"]:
        add_dfc("ADDED", name)
    for name in cmp["dfc_removed"]:
        add_dfc("REMOVED", name)

    for name, changes in cmp["dfc_modified"]:
        o = old["DFC"][name]
        n = new["DFC"][name]
        ws.append([
            "MODIFIED", name, n["description"],
            o["method"], n["method"], o["param"], n["param"],
            o["readiness"], n["readiness"], o["multiple"], n["multiple"],
            ", ".join(changes)
        ])

    for row in ws.iter_rows(min_row=2):
        fill = status_fill(row[0].value)
        if fill:
            for cell in row:
                cell.fill = fill
    add_table(ws, "DFC_Delta_Table")
    apply_common(ws)
    autosize(ws, 60)

    # FID delta
    ws = wb.create_sheet("03_FID_Delta")
    cols = [
        "Status", "FID", "Description",
        "Old IUMPR Group", "New IUMPR Group",
        "Old Scheduling Mode", "New Scheduling Mode",
        "Changed Fields"
    ]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    for name in cmp["fid_added"]:
        n = new["FID"][name]
        ws.append([
            "ADDED", name, n["description"], "", n["iumpr"],
            "", n["scheduling"], "ADDED"
        ])

    for name in cmp["fid_removed"]:
        o = old["FID"][name]
        ws.append([
            "REMOVED", name, o["description"], o["iumpr"], "",
            o["scheduling"], "", "REMOVED"
        ])

    for name, changes in cmp["fid_modified"]:
        o, n = old["FID"][name], new["FID"][name]
        ws.append([
            "MODIFIED", name, n["description"],
            o["iumpr"], n["iumpr"], o["scheduling"], n["scheduling"],
            ", ".join(changes)
        ])

    for row in ws.iter_rows(min_row=2):
        fill = status_fill(row[0].value)
        if fill:
            for cell in row:
                cell.fill = fill
    add_table(ws, "FID_Delta_Table")
    apply_common(ws)
    autosize(ws, 60)

    # DSQ delta
    ws = wb.create_sheet("04_DSQ_Delta")
    cols = ["Status", "DSQ", "Description", "Old Element Number", "New Element Number", "Changed Fields"]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    for name in cmp["dsq_added"]:
        n = new["DSQ"][name]
        ws.append(["ADDED", name, n["description"], "", n["number"], "ADDED"])
    for name in cmp["dsq_removed"]:
        o = old["DSQ"][name]
        ws.append(["REMOVED", name, o["description"], o["number"], "", "REMOVED"])
    for name, changes in cmp["dsq_modified"]:
        o, n = old["DSQ"][name], new["DSQ"][name]
        ws.append(["MODIFIED", name, n["description"], o["number"], n["number"], ", ".join(changes)])

    for row in ws.iter_rows(min_row=2):
        fill = status_fill(row[0].value)
        if fill:
            for cell in row:
                cell.fill = fill
    add_table(ws, "DSQ_Delta_Table")
    apply_common(ws)
    autosize(ws, 60)

    # Complete DFC debounce comparison
    ws = wb.create_sheet("05_DFC_Debounce")
    cols = [
        "DFC", "Description",
        "Old Debounce Method", "New Debounce Method",
        "Old Debounce Parameter", "New Debounce Parameter",
        "Old Readiness Group", "New Readiness Group",
        "Old Multiple Debounce", "New Multiple Debounce",
        "Element/Collection MP Changed", "Functional Status"
    ]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    for name in sorted(set(old["DFC"]) | set(new["DFC"])):
        o = old["DFC"].get(name)
        n = new["DFC"].get(name)

        if o and n:
            functional = any(
                o[k] != n[k]
                for k in ["description", "method", "param", "readiness", "multiple"]
            )
            structural = (o["number"], o["mp"]) != (n["number"], n["mp"])
            status = "MODIFIED" if functional else "UNCHANGED"
            description = n["description"]
            row = [
                name, description,
                o["method"], n["method"], o["param"], n["param"],
                o["readiness"], n["readiness"], o["multiple"], n["multiple"],
                "YES" if structural else "NO", status
            ]
        elif n:
            row = [
                name, n["description"], "", n["method"], "", n["param"],
                "", n["readiness"], "", n["multiple"], "YES", "ADDED"
            ]
        else:
            row = [
                name, o["description"], o["method"], "", o["param"], "",
                o["readiness"], "", o["multiple"], "", "YES", "REMOVED"
            ]

        ws.append(row)

        fill = status_fill(row[-1])
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill

    add_table(ws, "DFC_Debounce_Table")
    apply_common(ws)
    autosize(ws, 65)

    # FID -> DFC mapping delta
    ws = wb.create_sheet("06_FID_DFC_Mapping")
    cols = ["Status", "FID", "DFC", "Old Present", "New Present",
            "Old FID Debounce/Category", "New FID Debounce/Category"]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    old_lookup = {(fid, item["dfc"]): item for fid, data in old["FID"].items() for item in data["mappings"]}
    new_lookup = {(fid, item["dfc"]): item for fid, data in new["FID"].items() for item in data["mappings"]}

    for key in cmp["map_added"]:
        fid, dfc = key
        item = new_lookup[key]
        ws.append(["ADDED", fid, dfc, "NO", "YES",
                   "", f'{item["debounce"]} / {item["category"]}'])
    for key in cmp["map_removed"]:
        fid, dfc = key
        item = old_lookup[key]
        ws.append(["REMOVED", fid, dfc, "YES", "NO",
                   f'{item["debounce"]} / {item["category"]}', ""])

    for row in ws.iter_rows(min_row=2):
        fill = status_fill(row[0].value)
        if fill:
            for cell in row:
                cell.fill = fill

    add_table(ws, "FID_DFC_Mapping_Delta")
    apply_common(ws)
    autosize(ws, 65)

    # FID mapping matrix
    ws = wb.create_sheet("07_FID_Mapping_Matrix")
    cols = ["FID", "Old DFC Count", "New DFC Count", "Added DFCs", "Removed DFCs"]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    for fid in sorted(set(old["FID"]) | set(new["FID"])):
        old_set = {x["dfc"] for x in old["FID"].get(fid, {}).get("mappings", [])}
        new_set = {x["dfc"] for x in new["FID"].get(fid, {}).get("mappings", [])}
        ws.append([
            fid,
            len(old_set),
            len(new_set),
            ", ".join(sorted(new_set - old_set)),
            ", ".join(sorted(old_set - new_set))
        ])

    add_table(ws, "FID_Mapping_Matrix")
    apply_common(ws)
    autosize(ws, 85)

    # Validation
    ws = wb.create_sheet("08_Validation")
    cols = ["Check", "Severity", "Count", "Details"]
    ws.append(cols)
    style_header(ws, 1, len(cols))

    old_dfcs = set(old["DFC"])
    new_dfcs = set(new["DFC"])
    old_refs = {dfc for _, dfc in cmp["old_map"]}
    new_refs = {dfc for _, dfc in cmp["new_map"]}

    validation_rows = [
        (
            "Old DFCs not referenced by any FID",
            "INFO",
            len(old_dfcs - old_refs),
            ", ".join(sorted(old_dfcs - old_refs)[:100])
        ),
        (
            "New DFCs not referenced by any FID",
            "INFO",
            len(new_dfcs - new_refs),
            ", ".join(sorted(new_dfcs - new_refs)[:100])
        ),
        (
            "Old FIDs with no DFC mapping",
            "INFO",
            sum(not v["mappings"] for v in old["FID"].values()),
            ", ".join(sorted(k for k, v in old["FID"].items() if not v["mappings"]))
        ),
        (
            "New FIDs with no DFC mapping",
            "INFO",
            sum(not v["mappings"] for v in new["FID"].values()),
            ", ".join(sorted(k for k, v in new["FID"].items() if not v["mappings"]))
        ),
        (
            "New FID mappings pointing to unknown DFC",
            "CHECK",
            len({dfc for _, dfc in cmp["new_map"] if dfc not in new_dfcs}),
            ", ".join(sorted({dfc for _, dfc in cmp["new_map"] if dfc not in new_dfcs})[:100])
        ),
    ]

    for row in validation_rows:
        ws.append(row)

    add_table(ws, "Validation_Table")
    apply_common(ws)
    autosize(ws, 100)

    # General formatting
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"

    wb.save(output_path)



# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("900x680")
        self.minsize(820, 600)
        self.old_var = tk.StringVar()
        self.new_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.format_var = tk.StringVar(value="HTML")
        self.status_var = tk.StringVar(value="Select two DSM reports and choose the output format.")
        self._ui()

    def _ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Sub.TLabel", font=("Segoe UI", 10))
        style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), padding=10)

        root = ttk.Frame(self, padding=28)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="DSM Report Comparator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(root, text="Compare any two DSM Reports and choose HTML, Excel, or both as output.", style="Sub.TLabel").pack(anchor="w", pady=(4, 22))

        box = ttk.LabelFrame(root, text="DSM Inputs", padding=18)
        box.pack(fill="x")
        self.file_row(box, 0, "OLD / BASELINE DSM Report", self.old_var, self.browse_old)
        self.file_row(box, 1, "NEW / TARGET DSM Report", self.new_var, self.browse_new)

        out = ttk.LabelFrame(root, text="Output", padding=18)
        out.pack(fill="x", pady=16)
        ttk.Label(out, text="Output folder").grid(row=0, column=0, sticky="w", padx=(0,10))
        ttk.Entry(out, textvariable=self.output_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(out, text="Browse...", command=self.browse_output).grid(row=0, column=2, padx=(10,0))
        out.columnconfigure(1, weight=1)

        fmt = ttk.LabelFrame(root, text="Output Format", padding=15)
        fmt.pack(fill="x")
        ttk.Radiobutton(fmt, text="HTML — Interactive report", variable=self.format_var, value="HTML").grid(row=0,column=0,sticky="w",padx=12,pady=5)
        ttk.Radiobutton(fmt, text="Excel — Detailed workbook", variable=self.format_var, value="EXCEL").grid(row=0,column=1,sticky="w",padx=12,pady=5)
        ttk.Radiobutton(fmt, text="Both HTML + Excel", variable=self.format_var, value="BOTH").grid(row=0,column=2,sticky="w",padx=12,pady=5)
        ttk.Label(fmt, text="Excel output requires openpyxl; HTML has no external package dependency.", foreground="#666666").grid(row=1,column=0,columnspan=3,sticky="w",padx=12,pady=(4,0))

        features = ttk.LabelFrame(root, text="Included Analysis", padding=15)
        features.pack(fill="x", pady=(16,0))
        items = [
            "DFC added / removed / modified", "FID added / removed / modified",
            "DSQ added / removed / modified", "All DFC debounce method & parameter comparison",
            "FID → DFC mapping delta", "Element / Collection MP structural-change detection",
            "Validation / orphan checks", "Interactive search and filtering in HTML",
        ]
        for i, item in enumerate(items):
            ttk.Label(features,text="• "+item).grid(row=i//2,column=i%2,sticky="w",padx=10,pady=3)

        actions=ttk.Frame(root); actions.pack(fill="x",pady=(22,0))
        self.compare=ttk.Button(actions,text="GENERATE REPORT",style="Action.TButton",command=self.start)
        self.compare.pack(side="left")
        self.open=ttk.Button(actions,text="Open Output Folder",command=self.open_folder,state="disabled")
        self.open.pack(side="left",padx=12)
        self.progress=ttk.Progressbar(root,mode="indeterminate"); self.progress.pack(fill="x",pady=(22,7))
        ttk.Label(root,textvariable=self.status_var).pack(anchor="w")
        ttk.Label(root,text=f" · {APP_NAME} v{VERSION} · Built by MS/EPK-PS ·").pack(side="bottom",anchor="w",pady=(18,0))

    def file_row(self,parent,row,label,var,cmd):
        ttk.Label(parent,text=label).grid(row=row,column=0,sticky="w",padx=(0,10),pady=8)
        ttk.Entry(parent,textvariable=var).grid(row=row,column=1,sticky="ew",pady=8)
        ttk.Button(parent,text="Browse...",command=cmd).grid(row=row,column=2,padx=(10,0),pady=8)
        parent.columnconfigure(1,weight=1)

    def browse_old(self):
        p=filedialog.askopenfilename(title="Select OLD DSM report",filetypes=[("DSM text report","*.txt"),("All files","*.*")])
        if p:
            self.old_var.set(p)
            if not self.output_var.get(): self.output_var.set(os.path.dirname(p))

    def browse_new(self):
        p=filedialog.askopenfilename(title="Select NEW DSM report",filetypes=[("DSM text report","*.txt"),("All files","*.*")])
        if p:
            self.new_var.set(p)
            if not self.output_var.get(): self.output_var.set(os.path.dirname(p))

    def browse_output(self):
        p=filedialog.askdirectory(title="Select output folder")
        if p: self.output_var.set(p)

    def open_folder(self):
        p=os.path.join(self.output_var.get(), "Output")
        if p and os.path.isdir(p): os.startfile(p)

    def start(self):
        oldp,newp,out=self.old_var.get().strip(),self.new_var.get().strip(),self.output_var.get().strip()
        fmt=self.format_var.get()
        if not oldp or not os.path.isfile(oldp):
            messagebox.showwarning("Input required","Select a valid OLD DSM .txt report."); return
        if not newp or not os.path.isfile(newp):
            messagebox.showwarning("Input required","Select a valid NEW DSM .txt report."); return
        if os.path.abspath(oldp)==os.path.abspath(newp):
            messagebox.showwarning("Input required","OLD and NEW reports must be different."); return
        if fmt in ("EXCEL","BOTH") and not OPENPYXL_AVAILABLE:
            messagebox.showerror("Excel dependency missing","Excel output requires openpyxl.\n\nRun RUN_DSM_COMPARATOR.bat once, or install it with:\npython -m pip install openpyxl"); return
        if not out:
            out=os.path.dirname(newp); self.output_var.set(out)
        out=os.path.join(out, "Output")
        os.makedirs(out,exist_ok=True)
        base=f"DSM_Comparison_{Path(oldp).stem}_vs_{Path(newp).stem}"
        html_output=os.path.join(out,base+".html") if fmt in ("HTML","BOTH") else None
        excel_output=os.path.join(out,base+".xlsx") if fmt in ("EXCEL","BOTH") else None
        self.compare.config(state="disabled"); self.open.config(state="disabled"); self.progress.start(10)
        self.status_var.set("Parsing DSM reports...")
        threading.Thread(target=self.worker,args=(oldp,newp,html_output,excel_output,fmt),daemon=True).start()

    def worker(self,oldp,newp,html_output,excel_output,fmt):
        try:
            old=parse_report(oldp); new=parse_report(newp)
            self.after(0,lambda:self.status_var.set("Comparing DSM entities and mappings..."))
            cmp=compare_reports(old,new)
            outputs=[]
            if html_output:
                self.after(0,lambda:self.status_var.set("Building interactive HTML report..."))
                build_html(old,new,cmp,html_output); outputs.append(html_output)
            if excel_output:
                self.after(0,lambda:self.status_var.set("Building Excel workbook..."))
                build_excel(old,new,cmp,excel_output); outputs.append(excel_output)
            summary=(f"Reports generated successfully.\n\nFormat: {fmt}\n"
                     f"DFC: {len(old['DFC'])} → {len(new['DFC'])} (+{len(cmp['dfc_added'])} / -{len(cmp['dfc_removed'])} / {len(cmp['dfc_modified'])} modified)\n"
                     f"FID: {len(old['FID'])} → {len(new['FID'])} (+{len(cmp['fid_added'])} / -{len(cmp['fid_removed'])} / {len(cmp['fid_modified'])} modified)\n"
                     f"DSQ: {len(old['DSQ'])} → {len(new['DSQ'])} (+{len(cmp['dsq_added'])} / -{len(cmp['dsq_removed'])} / {len(cmp['dsq_modified'])} modified)\n"
                     f"FID→DFC mappings: {len(cmp['old_map'])} → {len(cmp['new_map'])} (+{len(cmp['map_added'])} / -{len(cmp['map_removed'])})\n\n" + "\n".join(outputs))
            self.after(0,lambda:self.success(outputs,summary,html_output))
        except Exception as e:
            detail=traceback.format_exc(); self.after(0,lambda:self.fail(str(e),detail))

    def success(self,outputs,msg,html_output):
        self.progress.stop(); self.compare.config(state="normal"); self.open.config(state="normal"); self.status_var.set("Done.")
        if html_output:
            try: webbrowser.open(Path(html_output).resolve().as_uri())
            except Exception: pass
        messagebox.showinfo("DSM comparison complete",msg)

    def fail(self,error,detail):
        self.progress.stop(); self.compare.config(state="normal"); self.status_var.set("Comparison failed.")
        print(detail); messagebox.showerror("DSM comparison failed",error)

if __name__=="__main__":
    App().mainloop()
