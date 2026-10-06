#!/usr/bin/env python3
"""
Generate the synthetic test fixtures for the TPD add-in (issue #164).

    pip install openpyxl
    python tests/fixtures/generate_fixtures.py

Writes, next to this script:

    EQ_List_Sample.xlsx    a Smartsheet-style equipment-list export
    Schedule_Sample.xlsx   a Smartsheet-style project-schedule export

EVERYTHING in the output is fabricated: company names, tags, part numbers,
people, dates. The generator is seeded, so the content is the same every run
(the .xlsx zip container is not byte-identical - it carries write timestamps).
Nothing here reads, copies or derives from a real customer workbook. The files
are built from scratch (no copy-and-scrub), so they carry no author, path,
comment or hidden-sheet metadata either.

The EQ List mirrors the *shape* of a real export - 25 columns, a nested tree of
PARENT rows and item rows (Excel row grouping, parents above their children),
row fills by type, centred cells, trailing blank rows - and deliberately
includes the awkward inputs the add-in has to cope with. See tests/README.md
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

# Source column order is deliberately NOT the add-in's default order
# (Status, % Complete, Tasks, Start Date, End Date) - the add-in reorders to the
# saved list (#127), so the fixture should make that observable.
SCH_COLUMNS = [
    ("Tasks", 46.0), ("Duration", 10.0), ("Start Date", 13.0), ("End Date", 13.0),
    ("Predecessors", 14.0), ("% Complete", 12.0), ("Status", 14.0),
    ("Assigned To", 20.0), ("Comments", 40.0),
]
SCOL = {name: i + 1 for i, (name, _) in enumerate(SCH_COLUMNS)}

# (level, name, working-days duration, assignee) - leaf tasks carry dates;
# level 0/1 rows are summary (phase / group) rows and roll their children up.
SCH_PLAN = [
    (0, "Project Kickoff and Planning", None, None),
    (1, "Mobilization", None, None),
    (2, "Kickoff meeting", 1, "Project Manager"),
    (2, "Issue RFQs", 5, "Procurement"),
    (2, "Receive and tabulate bids", 10, "Procurement"),
    (2, "Award purchase orders", 3, "Project Manager"),
    (1, "Engineering", None, None),
    (2, "Piping and instrumentation diagrams", 15, "Engineering"),
    (2, "Equipment datasheets", 10, "Engineering"),
    (2, "Submittal review", 8, "Engineering"),
    (0, "Fabrication and Delivery", None, None),
    (1, "Skid Fabrication", None, None),
    (2, "Fabricate burner skid", 20, "Shop"),
    (2, "Fabricate pump skid", 18, "Shop"),
    (2, "Hydrotest and inspection", 4, "Quality"),
    (1, "Long-Lead Equipment", None, None),
    (2, "Boiler delivery to site", 40, "Procurement"),
    (2, "Control panel delivery to site", 30, "Procurement"),
    (2, "Instrument package delivery", 22, "Procurement"),
    (0, "Site Installation", None, None),
    (1, "Mechanical", None, None),
    (2, "Set boiler and skids", 6, "Field Crew"),
    (2, "Install fuel gas piping", 12, "Field Crew"),
    (2, "Install feedwater and blowdown piping", 14, "Field Crew"),
    (1, "Electrical and Controls", None, None),
    (2, "Pull power and control cable", 9, "Electrical"),
    (2, "Terminate and loop check instruments", 8, "Electrical"),
    (2, "Panel power-up", 2, "Electrical"),
    (0, "Commissioning and Closeout", None, None),
    (1, "Commissioning", None, None),
    (2, "Burner tuning and combustion test", 5, "Commissioning"),
    (2, "Performance test with customer witness", 3, "Commissioning"),
    (1, "Closeout", None, None),
    (2, "Punch list walk-down", 3, "Project Manager"),
    (2, "As-built drawings and O&M manuals", 10, "Engineering"),
    (2, "Final acceptance and handover", 1, "Project Manager"),
]
SCH_COMMENTS = {
    "Issue RFQs": "Send to the approved bidders list only.",
    "Submittal review": "Customer has 10 working days to respond.\nEscalate to the "
                        "project sponsor if the review runs past that.",
    "Boiler delivery to site": "Carrier has a 2-day delivery window; site must have "
                               "crane and laydown area available for the full week "
                               "of arrival, per the lift plan.",
    "Punch list walk-down": "Items carried over from the pre-commissioning walk-down "
                            "must be closed first.",
}


def _workdays_between(a, b):
    """Working days from a to b inclusive."""
    return sum(1 for i in range((b - a).days + 1)
               if (a + dt.timedelta(days=i)).weekday() < 5)


def _add_workdays(d, n):
    """The date n working days after d (d itself counts as day 1 for n == 1)."""
    d = d if d.weekday() < 5 else d + dt.timedelta(days=7 - d.weekday())
    while n > 1:
        d += dt.timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def build_schedule():
    rng = random.Random(SEED + 1)
    wb = Workbook()
    _stamp(wb)
    ws = wb.active
    ws.title = "0000_ Sample Site - Schedule"
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False)
    ws.sheet_format.outlineLevelRow = 2

    for name, width in SCH_COLUMNS:
        ws.column_dimensions[ws.cell(1, SCOL[name]).column_letter].width = width
        ws.cell(1, SCOL[name], name).font = BOLD

    # Pass 1: schedule the leaf tasks back-to-back with a little float.
    cursor = dt.date(2030, 1, 7)             # a Monday, deliberately far-future/fictional
    sched = []                               # (level, name, start, end, dur, who)
    for level, name, dur, who in SCH_PLAN:
        if dur is None:
            sched.append([level, name, None, None, None, who])
            continue
        start = cursor
        end = _add_workdays(start, dur)
        sched.append([level, name, start, end, dur, who])
        cursor = _add_workdays(end, 2) if rng.random() < 0.7 else end + dt.timedelta(days=1)
        while cursor.weekday() >= 5:
            cursor += dt.timedelta(days=1)

    # Pass 2: summary rows take min(start)/max(end) of their descendants.
    for i, rec in enumerate(sched):
        if rec[2] is not None:
            continue
        kids = []
        for later in sched[i + 1:]:
            if later[0] <= rec[0]:
                break
            if later[2] is not None:
                kids.append(later)
        rec[2] = min(k[2] for k in kids)
        rec[3] = max(k[3] for k in kids)
        rec[4] = _workdays_between(rec[2], rec[3])

    today = dt.date(2030, 2, 18)             # fictional "as of" date for status
    for r, (level, name, start, end, dur, who) in enumerate(sched, start=2):
        ws.row_dimensions[r].outlineLevel = level
        is_summary = level <= 1
        if end < today:
            pct, status = 1.0, "Complete"
        elif start > today:
            pct, status = 0.0, "Not Started"
        else:
            span = max((end - start).days, 1)
            pct, status = round((today - start).days / span, 2), "In Progress"
        if name in ("Submittal review", "Boiler delivery to site"):
            status = "At Risk"
        fill = (PARENT_TOP if level == 0 else PARENT_SUB) if is_summary else None
        row_vals = {
            "Tasks": name,
            "Duration": f"{dur}d",
            "Start Date": start, "End Date": end,
            "Predecessors": "" if is_summary or r == 4 else str(r - 1),
            "% Complete": pct, "Status": status,
            "Assigned To": who, "Comments": SCH_COMMENTS.get(name),
        }
        for cname, _ in SCH_COLUMNS:
            v = row_vals[cname]
            c = ws.cell(r, SCOL[cname], None if v in ("", None) else v)  # keep a real 0
            if cname in ("Start Date", "End Date"):
                c.number_format = "mm/dd/yy"
                c.alignment = CENTER
            elif cname == "% Complete":
                c.number_format = "0%"
                c.alignment = CENTER
            elif cname in ("Tasks", "Comments"):
                c.alignment = LEFT_WRAP if cname == "Comments" else Alignment(
                    horizontal="left", indent=level)
            else:
                c.alignment = CENTER
            c.font = BOLD if is_summary else PLAIN
            if fill is not None:
                c.fill = fill

    wb.save(OUT_DIR / "Schedule_Sample.xlsx")
    return len(sched)


if __name__ == "__main__":
    last_row = build_eq_list()
    n_tasks = build_schedule()
    print(f"EQ_List_Sample.xlsx   : {last_row} rows (incl. heading), "
          f"{len(EQ_COLUMNS)} columns")
    print(f"Schedule_Sample.xlsx  : {n_tasks} task rows, {len(SCH_COLUMNS)} columns")
