#!/usr/bin/env python3
"""
Generate the synthetic test fixtures for the TPD add-in (issue #164).

    pip install openpyxl
    python tests/fixtures/generate_fixtures.py

Writes, next to this script:

    EQ_List_Sample.xlsx    a Smartsheet-style equipment-list export
    Schedule_Sample.xlsx   a Smartsheet-style project-schedule export

EVERYTHING in the output is fabricated: company names, tags, part numbers,
people, dates, task names. The generator is seeded, so the content is the same
every run (the .xlsx zip container is not byte-identical - it carries write
timestamps).

What IS modelled on real exports is *structure only* - column headings and
order, column widths and which columns are hidden, the shape of the row
hierarchy, fill colours, number formats. No cell content is copied from, or
derived from, a real workbook. The files are built from scratch (no
copy-and-scrub), so they carry no author, path, comment or hidden-sheet
metadata either.

The EQ List mirrors the shape of a real export - 25 columns, a nested tree of
PARENT rows and item rows (Excel row grouping, parents above their children),
row fills by type, centred cells, trailing blank rows. The Schedule mirrors a
real schedule export - 16 columns (6 hidden), a task tree up to 6 levels deep,
real dates and percentages. Both deliberately include the awkward inputs the
add-in has to cope with. See tests/README.md
for what each fixture exercises and what to expect.
"""
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.properties import Outline

OUT_DIR = Path(__file__).resolve().parent
SEED = 164
CREATOR = "TPD Add-in test fixtures"
FIXED_STAMP = dt.datetime(2030, 1, 1, 0, 0, 0)

PARENT_TOP = PatternFill("solid", fgColor="5FB3F9")    # level 0/1 PARENT rows
PARENT_SUB = PatternFill("solid", fgColor="B9DDFC")    # level 2+ PARENT rows
ITEM_FILL = PatternFill("solid", fgColor="C6E7C8")     # item rows
REVIEW_FILL = PatternFill("solid", fgColor="FEFF85")   # "For Approval" rows
GRAY = PatternFill("solid", fgColor="BDBDBD")          # schedule: section rows
RED_FLAG = PatternFill("solid", fgColor="F87E7D")      # schedule: a flagged row
BOLD = Font(bold=True)
PLAIN = Font(bold=False)
CENTER = Alignment(horizontal="center")
LEFT = Alignment(horizontal="left")
LEFT_WRAP = Alignment(horizontal="left", wrap_text=True, vertical="top")


def _stamp(wb):
    wb.properties.creator = CREATOR
    wb.properties.lastModifiedBy = CREATOR
    wb.properties.created = FIXED_STAMP
    wb.properties.modified = FIXED_STAMP
    wb.properties.title = None


# =============================================================================
# EQ List
# =============================================================================

# Same 25 headings, same order, as the real export, with widths to match.
EQ_COLUMNS = [
    ("Status", 15.8), ("OWNER", 18.3), ("Purchased", 13.3), ("INTERNAL ID", 18.0),
    ("CUSTOMER ID", 14.1), ("Location", 22.7), ("TYPE1", 18.1), ("TYPE2", 24.2),
    ("SIZE", 11.6), ("RATING", 12.2), ("CONNECTION TYPE", 13.8), ("Description", 49.4),
    ("Manufacturer", 15.2), ("Vendor", 18.1), ("Model", 31.4), ("Details", 22.3),
    ("Cost", 12.7), ("Lead Time", 14.2), ("PO Number", 13.8), ("PO Issue Date", 12.8),
    ("Ship Date", 12.0), ("Redbox Row 1", 23.4), ("Redbox Row 2", 23.4),
    ("Redbox Row 3", 23.4), ("Filename", 23.4),
]
COL = {name: i + 1 for i, (name, _) in enumerate(EQ_COLUMNS)}

# Shape of the tree (outline level + P=PARENT row / i=item row / B=blank row),
# three top-level sections then trailing blank rows. Structure only.
EQ_TREE = (
    "0P 1P 2i 1P 2i 2i 2i 2i 2i 2i 2i 2i 1P 2P 3i 3i 3i 3i 3i 3i 3i 3i 2P 3P "
    "4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 4i 3P 4i 4i 4i 4i 4i 4i 4i "
    "0P 1P 2i 2i 3i 2i 2i 2i 2i 2i 2i 2i 2i 2i 1P 2P 3i 3i 2P 3i 3i 3i 3i 3i "
    "3i 3i 2P 3i 3i 3i 3i 1P 2i 2i 2i 1P 2i 2i "
    + " ".join(["0B"] * 14)
).split()

# Fictional companies. None of these is a real business.
MANUFACTURERS = [
    "Alder Valve Co", "Birchwood Controls", "Cedar Flow Systems", "Dunmore Pumps",
    "Elmstead Burners", "Fernhill Instruments", "Garnet Motor Works",
    "Hollis Thermal", "Ironbark Strainers", "Juniper Bellows", "Kestrel Fluid Tech",
    "Larkspur Gauges", "Marlow Heaters", "Northgate Fittings",
]
VENDORS = [
    "Alder Valve Co", "Birchwood Controls", "Cedar Flow Systems", "Dunmore Pumps",
    "Fernhill Instruments", "Hollis Thermal", "Juniper Bellows", "Northgate Fittings",
]
# Awkward group-column values, planted on a few rows (see tests/README.md).
VENDOR_BAD_CHARS = "Acme Valve/Fittings [West]: Ltd?"            # illegal in a sheet name
VENDOR_TOO_LONG = "International Combustion & Controls Corporation of America"  # > 31 chars
# Same first 31 characters as VENDOR_TOO_LONG, so the two truncate to the SAME
# sheet name and the add-in's collision suffixing (#57) has to kick in.
VENDOR_TOO_LONG_2 = "International Combustion & Controls Corporation of Asia"
VENDOR_CASE_VARIANT = " alder valve co "                          # normalizes to "Alder Valve Co"

LOCATIONS = [
    "Boiler House", "Pump Room", "Roof Deck", "North Yard", "Fuel Skid",
    "Control Room", "Mezzanine Level", "Process Bay", "Utility Corridor", "Tank Farm",
]
SIZES = ['1/4"', '1/2"', '3/4"', '1"', '1-1/2"', '2"', '2-1/2"', '3"', "3x4"]
RATINGS = ["150#", "290#", "580#", "700#", "800#"]
CONNECTIONS = ["BW", "FLANGED", "NPT", "SW"]
LEAD_TIMES = ["4wk", "6wk", "8wk", "12wk", "16wk"]
OWNERS = ["AX", "BQ", "CZ"]

# TYPE1 -> (tag prefix, [TYPE2 choices], [description templates], needs size/rating/conn)
EQ_TYPES = {
    "Valve": ("V", ["Ball", "Gate", "Globe", "Check", "Relief", "Butterfly"],
              ["{t2} valve, {size}", "{t2} valve, {size}, {rating}"], True),
    "Pump": ("P", ["Centrifugal", "Gear", "Diaphragm"],
             ["{t2} pump, close-coupled", "{t2} transfer pump"], False),
    "Motor": ("M", ["TEFC", "ODP", "Explosion proof"],
              ["{t2} motor, 3-phase", "{t2} motor with VFD duty rating"], False),
    "Instrument": ("I", ["Pressure Transmitter", "Temperature Transmitter",
                         "Flow Switch", "Level Switch", "Pressure Gauge"],
                   ["{t2}", "{t2}, loop powered"], False),
    "Strainer": ("S", ["Y-Type", "Basket"], ["{t2} strainer, {size}"], True),
    "Burner": ("B", ["Gas", "Dual fuel"], ["{t2} burner assembly"], False),
    "Heater": ("H", ["Electric", "Immersion"], ["{t2} heater"], False),
    "Bellows": ("E", ["Expansion Joint"], ["{t2}, {size}"], True),
    "Tank": ("T", ["Expansion", "Day tank"], ["{t2} tank"], False),
    "Purchased Assm": ("A", ["Skid"], ["Packaged skid assembly"], False),
}

PARENT_NAMES_TOP = [
    "Boiler Package A", "Boiler Package B", "Fuel Handling", "Heat Recovery",
    "Combustion Air", "Feedwater System", "Blowdown System", "Controls Package",
]
PARENT_NAMES_SUB = [
    "Burner Assembly", "Fuel Gas Train", "Pilot Skid", "Pump Package", "Instrument Rack",
    "Strainer Group", "Isolation Valves", "Expansion Group", "Vent Group",
]
DETAILS = [
    "Stainless", "Carbon steel", "Brass", "316SS", "Spring return", "Fail closed",
    "Locking handle", "With bracket", "0-100 psi", "NACE", "Cast iron body",
]
# Long, multi-line text for the wrapped-cell cases (#162).
WRAPPED_DETAILS = [
    "Supplier to confirm gasket material.\nSee spec section 3.2 for hydrotest "
    "requirements before shipment.",
    "Field-verify nozzle orientation against the piping layout; revised drawing "
    "to follow from engineering once the skid footprint is fixed.",
    "Includes mounting hardware.\nSpare seal kit to ship loose.\nTag plate to be "
    "stainless steel, wired to the unit.",
]


def _tag(rng, prefix, counters):
    counters[prefix] = counters.get(prefix, 100) + rng.randint(1, 4)
    n = counters[prefix]
    suffix = rng.choice(["", "", "", "A", "B"])
    return f"{prefix}{n}{suffix}"


def _model(rng):
    a = rng.choice(["KV", "DP", "FX", "LT", "HR", "QM", "ZN"])
    return rng.choice([
        f"{a}-{rng.randint(100, 990)}/{rng.randint(2, 9)}",
        f"{a}{rng.randint(10, 99)}-{rng.choice('ABCDEF')}{rng.randint(1, 9)}",
        f"{a}{rng.randint(100, 999)}",
    ])


def build_eq_list():
    rng = random.Random(SEED)
    wb = Workbook()
    _stamp(wb)
    ws = wb.active
    ws.title = "0000_ Sample Site - EQ List"        # <= 31 chars, like the real tab
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False)  # parents above children
    ws.sheet_format.outlineLevelRow = 4

    for name, width in EQ_COLUMNS:
        ws.column_dimensions[ws.cell(1, COL[name]).column_letter].width = width
        c = ws.cell(1, COL[name], name)
        c.font = BOLD

    counters = {}
    items = []                       # (row number) of every item row, in order
    row = 1
    first_parent_in_section = True
    for tok in EQ_TREE:
        row += 1
        level, kind = int(tok[0]), tok[1]
        ws.row_dimensions[row].outlineLevel = level
        vals = {}
        fill, bold = None, False

        if kind == "P":
            vals["Purchased"] = "PARENT"
            if level <= 1:
                vals["Description"] = rng.choice(PARENT_NAMES_TOP) + (
                    "" if level == 0 else f" - {rng.choice(PARENT_NAMES_SUB)}")
                fill, bold = PARENT_TOP, True
            else:
                t1 = rng.choice(list(EQ_TYPES))
                prefix = EQ_TYPES[t1][0]
                tag = _tag(rng, prefix, counters)
                vals.update({
                    "INTERNAL ID": tag, "Location": rng.choice(LOCATIONS),
                    "TYPE1": t1, "Description": rng.choice(PARENT_NAMES_SUB),
                    "Redbox Row 1": tag, "Filename": f"{tag} {rng.choice(PARENT_NAMES_SUB)}",
                })
                if level == 2 and rng.random() < 0.4:
                    vals["Manufacturer"] = rng.choice(MANUFACTURERS)
                if level == 3:
                    sz = rng.choice(SIZES)
                    vals["SIZE"] = sz
                    vals["Redbox Row 3"] = f'{sz}  {rng.choice(CONNECTIONS)}'
                fill = PARENT_SUB
        elif kind == "i":
            items.append(row)
            t1 = rng.choice(list(EQ_TYPES))
            prefix, t2s, descs, hardware = EQ_TYPES[t1]
            t2 = rng.choice(t2s)
            tag = _tag(rng, prefix, counters)
            mfr = rng.choice(MANUFACTURERS)
            model = _model(rng)
            size = rng.choice(SIZES) if (hardware or rng.random() < 0.15) else ""
            rating = rng.choice(RATINGS) if (hardware and rng.random() < 0.3) else ""
            conn = rng.choice(CONNECTIONS) if (hardware and rng.random() < 0.7) else ""
            desc = rng.choice(descs).format(t2=t2, size=size or "NPS", rating=rating or "150#")
            vals.update({
                "Status": "Approved", "OWNER": rng.choice(OWNERS),
                "Purchased": rng.choices(["N/R", "NO"], weights=[55, 45])[0],
                "INTERNAL ID": tag, "Location": rng.choice(LOCATIONS),
                "TYPE1": t1, "TYPE2": t2, "SIZE": size, "RATING": rating,
                "CONNECTION TYPE": conn, "Description": desc, "Manufacturer": mfr,
                "Vendor": rng.choice(VENDORS), "Model": model,
                "Redbox Row 1": tag, "Redbox Row 2": model,
                "Redbox Row 3": "  ".join(x for x in (size, rating, conn) if x),
                "Filename": f"{tag} {mfr} {model}",
            })
            if rng.random() < 0.4:
                vals["Details"] = rng.choice(DETAILS)
            if rng.random() < 0.14:
                vals["CUSTOMER ID"] = f"C{rng.randint(1, 40)}{rng.choice(['', 'A'])}"
            if rng.random() < 0.10:
                vals["Lead Time"] = rng.choice(LEAD_TIMES)
            fill = ITEM_FILL
        # kind == "B": a blank row, no values, no fill

        for name, _ in EQ_COLUMNS:
            c = ws.cell(row, COL[name], vals.get(name) or None)
            c.alignment = LEFT if name == "Description" else (
                Alignment() if name == "CUSTOMER ID" else CENTER)
            if fill is not None:
                c.fill = fill
            c.font = BOLD if bold else PLAIN

    # ---- planted edge cases (tests/README.md lists them) ---------------------
    def put(r, name, value, wrap=False):
        c = ws.cell(r, COL[name])
        c.value = value                 # explicit: ws.cell(..., None) would not clear
        if wrap:
            c.alignment = LEFT_WRAP
        return c

    # two "For Approval" rows (yellow, like the real export's review rows)
    for r in (items[5], items[40]):
        put(r, "Status", "For Approval")
        for name, _ in EQ_COLUMNS:
            ws.cell(r, COL[name]).fill = REVIEW_FILL
    # one INCLUDED row and one cost + one with a lead time and a numeric tag
    put(items[12], "Purchased", "INCLUDED")
    put(items[20], "Cost", 1250.00)
    put(items[30], "INTERNAL ID", 7.2)
    put(items[30], "Redbox Row 1", 7.2)
    # a TYPE2 with an embedded line break (real exports have these)
    put(items[8], "TYPE2", "Pressure Transmitter\n(4-20mA)")
    # wrapped, multi-line Details - exercises row-height autofit (#162)
    for r, text in zip((items[15], items[33], items[50]), WRAPPED_DETAILS):
        put(r, "Details", text, wrap=True)
    # Split-by-Vendor edge cases: illegal sheet-name chars, > 31 chars, a
    # case/whitespace variant of a real vendor, and blank vendors
    put(items[2], "Vendor", VENDOR_BAD_CHARS)
    put(items[18], "Vendor", VENDOR_BAD_CHARS)
    put(items[25], "Vendor", VENDOR_TOO_LONG)
    put(items[44], "Vendor", VENDOR_TOO_LONG_2)
    put(items[10], "Vendor", "Alder Valve Co")
    put(items[11], "Vendor", VENDOR_CASE_VARIANT)
    for r in (items[7], items[36], items[60]):
        put(r, "Vendor", None)

    wb.save(OUT_DIR / "EQ_List_Sample.xlsx")
    return ws.max_row


# =============================================================================
# Schedule
# =============================================================================

# Same 16 headings, same order, as the real export. Six of them are HIDDEN in
# the real file, and stay hidden here: the add-in has to cope with hidden
# source columns (they should still be selectable in the picker).
SCH_COLUMNS = [
    ("Status", 13.8, False), ("Computed Status", 8.0, True), ("% Complete", 14.4, False),
    ("Tasks", 57.0, False), ("Duration", 14.2, False), ("Start Date", 12.8, False),
    ("End Date", 12.0, False), ("Completion Date", 14.2, False), ("Predecessors", 8.0, True),
    ("MANUFACTURER", 25.0, False), ("Assigned To", 23.6, False), ("Baseline Start", 8.0, True),
    ("Baseline Finish", 8.0, True), ("Variance", 8.0, True), ("Ready to Start", 8.0, True),
    ("Notes", 30.5, False),
]
SCOL = {name: i + 1 for i, (name, _, _) in enumerate(SCH_COLUMNS)}

# Outline level of every row under the heading, in order; "B" = a blank row.
# Structure only - 27 summary rows, 108 task rows, 24 trailing blank rows. A row
# is a summary row when the next row is nested deeper (Smartsheet's parent/child).
SCH_LEVELS = (
    "0 1 2 3 3 2 3 3 3 3 3 3 3 3 2 2 2 3 3 3 2 3 3 3 3 3 3 3 4 4 5 5 3 2 3 3 3 3 3 3 3 4 "
    "4 4 1 1 2 3 3 4 4 4 3 4 4 3 4 4 3 4 4 3 4 4 3 4 4 3 4 4 2 3 3 3 3 3 3 3 2 3 3 4 3 2 "
    "3 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 4 2 3 3 3 3 3 3 3 3 3 3 3 2 3 3 3 3 "
    "3 1 1 0 1 1 1 1 1 "
    + " ".join(["B"] * 24)
).split()

PEOPLE = ["Alex Example", "Blake Example", "Casey Example",
          "Drew Example", "Emery Example", "Finley Example"]
SCH_AREAS = ["Boiler Island", "Fuel Handling", "Heat Recovery", "Water Treatment",
             "Controls and Electrical"]
SCH_GROUPS = ["Engineering", "Procurement", "Fabrication", "Delivery", "Site Work",
              "Installation", "Commissioning", "Closeout"]
SCH_SUBGROUPS = ["Mechanical", "Electrical", "Piping", "Instrumentation", "Structural",
                 "Documentation"]
SCH_ITEMS = ["Burner skid", "Fuel gas train", "Feedwater pump set", "Economizer",
             "Control panel", "Stack and damper", "Blowdown tank", "Strainer group",
             "Expansion joints", "Pressure transmitters", "Isolation valves",
             "Pilot assembly", "Day tank", "Deaerator", "Flow meters"]
SCH_VERBS = ["Order", "Receive", "Fabricate", "Inspect", "Install", "Test",
             "Submit documentation for", "Review drawings for", "Deliver", "Commission"]
SCH_FAKE_TEXT_DURATION = "3 days"      # one odd text value, as in the real export


def _workday(d):
    while d.weekday() >= 5:
        d += dt.timedelta(days=1)
    return d


def _add_workdays(d, n):
    """The working day n-1 days after d (n=1 -> d itself); n may be 0/negative."""
    d = _workday(d)
    step = 1 if n >= 0 else -1
    left = abs(n - 1) if n >= 1 else abs(n)
    while left > 0:
        d += dt.timedelta(days=step)
        if d.weekday() < 5:
            left -= 1
    return d


def _workdays_between(a, b):
    """Signed working-day count from a to b (b later => positive)."""
    if a == b:
        return 0
    lo, hi, sign = (a, b, 1) if a < b else (b, a, -1)
    n, d = 0, lo
    while d < hi:
        d += dt.timedelta(days=1)
        if d.weekday() < 5:
            n += 1
    return sign * n


def build_schedule():
    rng = random.Random(SEED + 1)
    wb = Workbook()
    _stamp(wb)
    ws = wb.active
    ws.title = "0000_ Sample Site - Schedule"
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False)
    ws.sheet_format.outlineLevelRow = 5

    for name, width, hidden in SCH_COLUMNS:
        letter = ws.cell(1, SCOL[name]).column_letter
        ws.column_dimensions[letter].width = width
        ws.column_dimensions[letter].hidden = hidden
        ws.cell(1, SCOL[name], name).font = BOLD

    # ---- shape the tree ------------------------------------------------------
    n = len(SCH_LEVELS)
    rows = []                                        # one dict per row under the heading
    for i, tok in enumerate(SCH_LEVELS):
        if tok == "B":
            rows.append({"blank": True, "level": 0, "row": i + 2})
            continue
        level = int(tok)
        nxt = SCH_LEVELS[i + 1] if i + 1 < n else "B"
        summary = nxt != "B" and int(nxt) > level
        rows.append({"blank": False, "level": level, "summary": summary,
                     "row": i + 2})

    # ---- names ---------------------------------------------------------------
    area_i = grp_i = sub_i = 0
    item_pool = SCH_ITEMS[:]
    rng.shuffle(item_pool)
    current_item = rng.choice(SCH_ITEMS)
    for r in rows:
        if r["blank"]:
            continue
        lv = r["level"]
        if r["summary"]:
            if lv == 0:
                r["name"] = SCH_AREAS[area_i % len(SCH_AREAS)]
                area_i += 1
            elif lv == 1:
                r["name"] = SCH_GROUPS[grp_i % len(SCH_GROUPS)]
                grp_i += 1
            elif lv == 2:
                r["name"] = SCH_SUBGROUPS[sub_i % len(SCH_SUBGROUPS)]
                sub_i += 1
            else:
                current_item = item_pool[0]
                item_pool.append(item_pool.pop(0))
                r["name"] = current_item
        else:
            if lv >= 4:
                r["name"] = f"{rng.choice(SCH_VERBS)} {current_item.lower()}"
            else:
                r["name"] = f"{rng.choice(SCH_VERBS)} {rng.choice(SCH_ITEMS).lower()}"

    # ---- dates: leaves first (working days only), then roll summaries up ------
    project_start = dt.date(2030, 1, 7)                # a Monday; far-future, obviously fictional
    cursor = project_start
    leaves = [r for r in rows if not r["blank"] and not r["summary"]]
    for r in rows:
        if r["blank"]:
            continue
        if r["summary"]:
            # each item group starts somewhere inside the project window, so groups
            # overlap like a real schedule instead of running end to end
            if r["level"] >= 2:
                cursor = _add_workdays(project_start, rng.randint(1, 200))
            continue
        unit = rng.choices(["d", "w"], weights=[85, 15])[0]
        count = rng.randint(1, 6) if unit == "d" else rng.randint(1, 3)
        working = count if unit == "d" else count * 5
        r["duration_text"] = f"{count}{unit}"
        r["start"] = _workday(cursor)
        r["end"] = _add_workdays(r["start"], working)
        cursor = _add_workdays(r["end"], rng.choice([1, 1, 2]))
        if rng.random() < 0.3:                          # overlap with the next task
            cursor = _add_workdays(r["start"], max(1, working // 2))
    for i, r in enumerate(rows):
        if r["blank"] or not r["summary"]:
            continue
        kids = []
        for later in rows[i + 1:]:
            if later["blank"] or later["level"] <= r["level"]:
                break
            if not later["summary"]:
                kids.append(later)
        r["start"] = min(k["start"] for k in kids)
        r["end"] = max(k["end"] for k in kids)
        total = _workdays_between(r["start"], r["end"]) + 1
        r["duration_text"] = total if rng.random() < 0.5 else f"{total}d"   # a mix, as in the real file
    for r in rows:
        if r["blank"]:
            continue
        # Baseline = the plan as first agreed. Variance = finish - baseline finish,
        # so a baseline that sits later than today's dates gives a NEGATIVE
        # variance - which is what most rows look like in the real export.
        off_s = rng.choice([0, 0, 2, 5, 8, 12, 15, 20])
        off_f = rng.choice([0, 3, 5, 8, 10, 12, 15, 20])
        if rng.random() < 0.1:
            off_f = -off_f
        r["b_start"] = _add_workdays(r["start"], 1 + off_s)
        r["b_finish"] = _add_workdays(r["end"], 1 + off_f)
        slip = _workdays_between(r["b_finish"], r["end"])
        r["variance"] = f"{slip}d"

    # ---- status: the earliest 14 leaves are done, two more are underway --------
    as_of = None
    for k, r in enumerate(leaves):
        if k < 14:
            r["status"], r["pct"] = "Complete", 1.0
            as_of = r["end"]
        elif k < 16:
            r["status"], r["pct"] = "In Progress", rng.choice([0.12, 0.21, 0.5])
        else:
            r["status"], r["pct"] = "Not Started", 0.0
    as_of = as_of + dt.timedelta(days=1)
    for i, r in enumerate(rows):
        if r["blank"] or not r["summary"]:
            continue
        kids = []
        for later in rows[i + 1:]:
            if later["blank"] or later["level"] <= r["level"]:
                break
            if not later["summary"]:
                kids.append(later)
        pct = round(sum(k["pct"] for k in kids) / len(kids), 2)
        r["pct"] = pct
        r["status"] = "Complete" if pct == 1.0 else ("Not Started" if pct == 0.0 else "In Progress")

    # ---- who / what / dependencies -------------------------------------------
    leaf_rows = [r["row"] for r in leaves]
    mfr_pool = rng.sample(MANUFACTURERS, 11)
    owner_of_parent = {}
    cur_mfr, cur_who = rng.choice(mfr_pool), rng.choice(PEOPLE)
    for i, r in enumerate(rows):
        if r["blank"]:
            continue
        if r["summary"] and r["level"] >= 3:
            cur_mfr, cur_who = rng.choice(mfr_pool), rng.choice(PEOPLE)
        r["mfr"] = cur_mfr if (not r["summary"] and r["level"] >= 3) or (
            r["summary"] and rng.random() < 0.45) else None
        r["who"] = cur_who if (not r["summary"] and r["level"] >= 3) or (
            r["summary"] and rng.random() < 0.5) else (
            rng.choice(PEOPLE) if rng.random() < 0.6 else None)
        r["ready"] = (r["start"] - as_of).days if rng.random() < 0.97 else None
        pred = None
        if not r["summary"] and rng.random() < 0.75:
            earlier = [x for x in leaf_rows if x < r["row"]]
            if earlier:
                prev = earlier[-1]
                roll = rng.random()
                pred = (prev if roll < 0.78 else
                        f"{prev}FS +{rng.randint(1, 5)}d" if roll < 0.9 else
                        f"{prev}FF" if roll < 0.95 else
                        f"{earlier[-2] if len(earlier) > 1 else prev}, {prev}")
        r["pred"] = pred

    # a few completed rows carry a completion date (one is odd text)
    done = [r for r in leaves if r["status"] == "Complete"]
    for r in done[:12]:
        r["completion"] = r["end"]
    if len(done) > 12:
        done[12]["completion"] = SCH_FAKE_TEXT_DURATION

    # ---- write ---------------------------------------------------------------
    gray_used = 0
    red_used = False
    for r in rows:
        rn = r["row"]
        ws.row_dimensions[rn].outlineLevel = r["level"]
        fill, bold = None, False
        if not r["blank"]:
            lv = r["level"]
            if r["summary"]:
                bold = lv != 1 or rng.random() < 0.5
                if lv == 0:
                    fill = PARENT_TOP
                elif lv == 1:
                    if not red_used and r["status"] == "Not Started" and rng.random() < 0.25:
                        fill, red_used = RED_FLAG, True
                    elif rng.random() < 0.55:
                        fill = PARENT_SUB
                elif lv == 2:
                    fill = GRAY
                elif lv == 3:
                    bold = True
            elif lv == 3 and gray_used < 2 and rng.random() < 0.1:
                fill, gray_used = GRAY, gray_used + 1
        values = {} if r["blank"] else {
            "Status": r["status"], "% Complete": r["pct"],
            "Tasks": r["name"], "Duration": r["duration_text"],
            "Start Date": r["start"], "End Date": r["end"],
            "Completion Date": r.get("completion"), "Predecessors": r["pred"],
            "MANUFACTURER": r["mfr"], "Assigned To": r["who"],
            "Baseline Start": r["b_start"], "Baseline Finish": r["b_finish"],
            "Variance": r["variance"], "Ready to Start": r["ready"],
        }
        for cname, _, _ in SCH_COLUMNS:
            v = values.get(cname)
            c = ws.cell(rn, SCOL[cname], None if v in ("", None) else v)
            if cname in ("Start Date", "End Date", "Completion Date",
                         "Baseline Start", "Baseline Finish"):
                c.number_format = "MM/dd/yy"
                c.alignment = CENTER
            elif cname == "% Complete":
                c.number_format = "#,##0%;-#,##0%"
                c.alignment = Alignment(horizontal="right")
            elif cname == "Variance":
                c.alignment = Alignment(horizontal="right")
            elif cname == "Tasks":
                c.alignment = Alignment(horizontal="left", indent=r["level"])
            elif cname in ("Status", "Duration", "MANUFACTURER"):
                c.alignment = CENTER
            c.font = BOLD if (bold or (cname == "Completion Date" and v is not None)) else PLAIN
            if fill is not None:
                c.fill = fill

    wb.save(OUT_DIR / "Schedule_Sample.xlsx")
    return len(rows)


if __name__ == "__main__":
    last_row = build_eq_list()
    n_rows = build_schedule()
    print(f"EQ_List_Sample.xlsx   : {last_row} rows (incl. heading), "
          f"{len(EQ_COLUMNS)} columns")
    print(f"Schedule_Sample.xlsx  : {n_rows + 1} rows (incl. heading), "
          f"{len(SCH_COLUMNS)} columns")
