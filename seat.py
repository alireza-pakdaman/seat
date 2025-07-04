from __future__ import annotations

import datetime, json, pathlib, random, shutil, sys, re
import tkinter as tk
from   tkinter import filedialog, messagebox

import numpy as np, pandas as pd, openpyxl
from   openpyxl.utils import get_column_letter
from   openpyxl.styles import PatternFill                       # NEW

# ─────────────────────────────────────────────────────────────────────────────
#  Seat catalogue – mirrors the visual grid
# ─────────────────────────────────────────────────────────────────────────────
SEATS: dict[str, dict[str, object]] = {
    # Workstations 1-10 (WS 6 is height adjustable)
    **{f"WS {n}":  {"type": "ws",  "adjustable": n == 6, "seat_number": n}  for n in range(1, 11)},
    
    # Regular seats 1-30 (keeping original configuration)
    **{f"Seat {n}": {"type": "reg", "adjustable": n in (13, 15, 30), "seat_number": n}
       for n in range(1, 31)},
    
    # Private rooms 345-354 (345, 347, 348, 349, 351, 354 are height adjustable)
    **{f"Room {n}": {"type": "pr",  "adjustable": n in (345, 347, 348, 349, 351, 354), "seat_number": n}
       for n in range(345, 355)},
    
    # SAS offices 1-12 (SAS 5 is height adjustable)
    **{f"SAS {n}": {"type": "sas", "adjustable": n == 5, "enabled": True, "seat_number": n} for n in range(1, 13)},
    
    # SHA Classrooms
    # SHA 356: 17 seats, seat 1 is height adjustable
    **{f"SHA 356 Seat {n}": {"type": "sha", "classroom": 356, "adjustable": n == 1, "seat_number": n} for n in range(1, 18)},
    
    # SHA 357: 19 seats, seat 1 is height adjustable
    **{f"SHA 357 Seat {n}": {"type": "sha", "classroom": 357, "adjustable": n == 1, "seat_number": n} for n in range(1, 20)},
    
    # SHA 359: 18 seats, seat 1 is height adjustable
    **{f"SHA 359 Seat {n}": {"type": "sha", "classroom": 359, "adjustable": n == 1, "seat_number": n} for n in range(1, 19)},
    
    # SHA 358: 31 seats, seats 1, 2, 3 are height adjustable
    **{f"SHA 358 Seat {n}": {"type": "sha", "classroom": 358, "adjustable": n in (1, 2, 3), "seat_number": n} for n in range(1, 32)},
}

# ───────────────────────────── Constants ──────────────────────────────────
TIME_FORMAT_EXCEL = "h:mm AM/PM"      # makes 22:00 look like 10:00 PM
COLOR_BY_TYPE   = {                   # SHA=yellow, SAS=blue, …
    "sha": "FFF2CC",
    "sas": "DDEBF7",
    "pr":  "E2F0D9",
    "ws":  "FCE4D6",
    "reg": None,                      # no colour for regular seats
}
MIN_COL_WIDTH = 10

# ───────────────────────────── GUI helpers ──────────────────────────────────
def pick_file(title: str, patterns: tuple[tuple[str, str], ...]) -> str:
    root = tk.Tk(); root.withdraw()
    return filedialog.askopenfilename(title=title, filetypes=patterns)

def pick_folder(title: str) -> str:
    root = tk.Tk(); root.withdraw()
    return filedialog.askdirectory(title=title)

# ───────────────────────────── Data ingest ──────────────────────────────────
def read_source(path: str) -> pd.DataFrame:
    """Load roster, tag 'Requires Adjustable', coerce time-like columns."""
    df = (pd.read_csv(path, header=1)
          if path.lower().endswith(".csv")
          else pd.read_excel(path, header=1))

    df["Requires Adjustable"] = (
        df["Test Accommodation"]
        .str.contains("Height Adjustable", case=False, na=False)
    )

    def to_time(col: str) -> pd.Series:
        # Suppress the dateutil warning by being more specific
        with pd.option_context('mode.chained_assignment', None):
            return (pd.to_datetime(df[col].astype(str), errors="coerce", format='mixed')
                    .dt.time.fillna(datetime.time.min))

    for col in ("Begin Time", "End Time", "Class Time"):
        df[col] = to_time(col)

    return df

# ─────────────────────── Seat-assignment engine ─────────────────────────────
def assign_students(df: pd.DataFrame, seat_pool: list[str]
                    ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Greedy, stable assignment of students to seats without overlapping times:
    adjustable-desk requests are placed first; within each sub-group students
    are processed in chronological order.  The pool is shuffled every pass
    for fairness.
    """
    availability = {s: datetime.time.min for s in seat_pool}
    rng          = np.random.default_rng()
    placed, left = [], []

    for needs_adj in (True, False):
        group = df[df["Requires Adjustable"] == needs_adj]\
                  .sort_values("Begin Time")
        for _, stu in group.iterrows():
            valid = [s for s in seat_pool
                     if (not needs_adj or SEATS[s]["adjustable"])]
            rng.shuffle(valid)

            for seat in valid:
                if stu["Begin Time"] >= availability[seat]:
                    availability[seat] = stu["End Time"]
                    placed.append({**stu, "Test Room": seat})
                    break
            else:
                left.append(stu)

    return pd.DataFrame(placed), pd.DataFrame(left)

# ───────────────────────────── Constants and helpers ──────────────────────────────────
def _set_cell(ws, row, col, value, *, fill_hex=None):
    """
    Write *value* into (row, col), prettify if it's a time,
    and optionally apply a background fill colour.
    """
    cell = ws.cell(row=row, column=col, value=value)

    # Make times human-readable in Excel
    if isinstance(value, datetime.time):
        cell.number_format = TIME_FORMAT_EXCEL

    # Shade the cell if a colour was supplied
    if fill_hex:
        cell.fill = PatternFill("solid", fgColor=fill_hex)

    return cell

# ───────────────────────────── Enhanced Excel output helper ──────────────────────────────
def write_excel(df: pd.DataFrame, name: str,
                template: str | None,
                out_dir: str | pathlib.Path,
                assignment_type: str = "ASSIGNED") -> None:

    out_path = pathlib.Path(out_dir) / f"{name}.xlsx"

    # ── 1  Set up the workbook ────────────────────────────────────────────────
    if template:
        shutil.copy(template, out_path)
        wb = openpyxl.load_workbook(out_path)
        # Try to use "Master" sheet, fall back to active sheet if not found
        try:
            ws = wb["Master"]
        except KeyError:
            ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        # Create basic headers if no template
        headers = ['Begin Time', 'End Time', 'Student Number', 'Student Last Name', 'Student First Name',
                  'Check-IN Time', 'Check-OUT Time', 'Course', 'Code', 'Test Room', 'Seat Number', 
                  'Faculty Name', 'Class Time', 'Test Accommodation', 'Invigilator Comment', 'Test Comment']
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=2, column=col_idx, value=header)

    # ── 2  Locate or create "Test Room" & "Seat Number" columns ───────────────
    hdr_row      = 2
    headers      = {ws.cell(hdr_row, col).value: col for col in range(1, ws.max_column+1)}
    test_room_col   = headers.get("Test Room")  or ws.max_column + 1
    seat_number_col = headers.get("Seat Number") or (test_room_col + 1)

    ws.cell(hdr_row, test_room_col,   "Test Room")
    ws.cell(hdr_row, seat_number_col, "Seat Number")

    # ── 3  Clear previous data completely ────────────────────────────────────
    if ws.max_row > hdr_row:
        for row in ws.iter_rows(min_row=hdr_row + 1, max_row=ws.max_row):
            for cell in row:
                cell.value = None   # wipe values & formula results

    # ── 4  Write this cohort's data ───────────────────────────────────────────
    for excel_row, (_, stu) in enumerate(df.iterrows(), start=hdr_row + 1):

        # Core student columns
        _set_cell(ws, excel_row, 1,  stu.get('Begin Time'))          # Begin Time
        _set_cell(ws, excel_row, 2,  stu.get('End Time'))            # End Time
        _set_cell(ws, excel_row, 3,  stu.get('Student Number'))
        _set_cell(ws, excel_row, 4,  stu.get('Student Last Name'))
        _set_cell(ws, excel_row, 5,  stu.get('Student First Name'))
        _set_cell(ws, excel_row, 8,  stu.get('Course'))
        _set_cell(ws, excel_row, 9,  stu.get('Code'))
        _set_cell(ws, excel_row, 12, stu.get('Faculty Name'))
        _set_cell(ws, excel_row, 13, stu.get('Class Time'))
        _set_cell(ws, excel_row, 14, stu.get('Test Accommodation'))

        # ----------  Test Room & Seat Number  ----------
        if assignment_type == "ASSIGNED" and pd.notna(stu.get('Test Room')):
            test_room = stu['Test Room']

            # Pick colour based on room type
            seat_type = SEATS.get(test_room, {}).get("type", "reg")
            fill_hex  = COLOR_BY_TYPE.get(seat_type)

            _set_cell(ws, excel_row, test_room_col, test_room, fill_hex=fill_hex)

            # Seat number logic unchanged
            seat_no = (SEATS.get(test_room, {}).get("seat_number")
                       or re.search(r"\b(?:Seat|WS|Room|SAS) (\d+)", test_room))
            if seat_no:
                _set_cell(ws, excel_row, seat_number_col,
                          seat_no.group(1) if hasattr(seat_no, "group") else seat_no)

    # ── 5  Adjust widths, save, done ─────────────────────────────────────────
    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        if all(cell.value is None for cell in column[2:]):
            ws.column_dimensions[letter].width = MIN_COL_WIDTH
            continue
        max_len = max(len(str(cell.value)) for cell in column if cell.value)
        ws.column_dimensions[letter].width = min(max_len + 2, 50)

    wb.save(out_path)
    print(f"📄  Saved {out_path.name}  ({len(df)} rows)")

# ─────────────────────────────── Main ───────────────────────────────────────
def main() -> None:
    # 1  Pick the roster
    data = pick_file("Select your Excel/CSV roster",
                     (("Excel/CSV", "*.xlsx *.xls *.csv"),))
    if not data:
        print("No data file selected – exiting."); sys.exit()

    # 2  Ask whether Excel outputs are wanted
    want_excel = messagebox.askyesno(
        "Excel outputs?",
        "Generate the usual cohort workbooks as well?\n"
        "(No → skip straight to JSON output only.)"
    )
    
    # 3  Pick template (default to May 5, 2025.xlsx)
    template = None
    if want_excel:
        use_default_template = messagebox.askyesno(
            "Use Default Template?",
            "Use 'May 5, 2025.xlsx' as the template?\n\n"
            "This will maintain the standard format with proper\n"
            "Test Room and Seat Number columns.\n\n"
            "Choose 'No' to select a different template."
        )
        
        if use_default_template:
            template_path = pathlib.Path(__file__).parent / "May 5, 2025.xlsx"
            if template_path.exists():
                template = str(template_path)
                print(f"📋 Using default template: {template}")
            else:
                print("❌ May 5, 2025.xlsx not found in current directory")
                template = pick_file("Select an .xlsx template",
                                   (("Excel", "*.xlsx *.xls"),))
        else:
            template = pick_file("Select an .xlsx template",
                               (("Excel", "*.xlsx *.xls"),))
        
        if not template:
            print("No template chosen; Excel outputs will be basic workbooks.")

    # 4  Ask about each room type individually
    room_preferences = {}
    
    # Private Rooms
    room_preferences['private_rooms'] = messagebox.askyesno(
        "Private Rooms",
        "Include Private Rooms (345-354) in seat assignments?\n\n"
        "Private rooms are typically used for students requiring\n"
        "isolated testing environments or specific accommodations.\n\n"
        "Available: 10 private rooms\n"
        "Height adjustable: Rooms 345, 347, 348, 349, 351, 354"
    )
    
    # Workstations
    room_preferences['workstations'] = messagebox.askyesno(
        "Workstations",
        "Include Workstations (WS 1-10) in seat assignments?\n\n"
        "Workstations are equipped with computers and are suitable\n"
        "for students requiring MS Word, Kurzweil, or similar software.\n\n"
        "Available: 10 workstations\n"
        "Height adjustable: WS 6"
    )
    
    # Regular Seats
    room_preferences['regular_seats'] = messagebox.askyesno(
        "Regular Seats",
        "Include Regular Seats (1-30) in seat assignments?\n\n"
        "Standard testing seats in the main examination area.\n\n"
        "Available: 30 regular seats\n"
        "Height adjustable: Seats 13, 15, 30"
    )
    
    # SAS Offices
    room_preferences['sas_offices'] = messagebox.askyesno(
        "SAS Offices",
        "Include SAS Offices (1-12) in seat assignments?\n\n"
        "SAS offices provide quiet, individual testing environments\n"
        "for students with specific accommodation needs.\n\n"
        "Available: 12 SAS offices\n"
        "Height adjustable: SAS 5"
    )
    
    # SHA Classrooms
    room_preferences['sha_classrooms'] = messagebox.askyesno(
        "SHA Classrooms",
        "Include SHA Classrooms in seat assignments?\n\n"
        "SHA classrooms provide group testing environments:\n"
        "• SHA 356: 17 seats (seat 1 height adjustable)\n"
        "• SHA 357: 19 seats (seat 1 height adjustable)\n"
        "• SHA 358: 31 seats (seats 1,2,3 height adjustable)\n"
        "• SHA 359: 18 seats (seat 1 height adjustable)\n\n"
        "Total: 85 classroom seats"
    )

    # 5  Pick an output folder
    out = pick_folder("Select an output folder")
    if not out:
        print("No output folder – exiting."); sys.exit()

    # 6  Load & preprocess the roster
    df = read_source(data)

    # 7  Cohort splits – EXACTLY as in your original logic
    print("🔍 Debugging cohort splits...")
    print(f"📋 Total students in roster: {len(df)}")
    
    # Debug Test Accommodation column
    print(f"\n🔍 Test Accommodation column analysis:")
    unique_accommodations = df["Test Accommodation"].value_counts()
    print(f"   Unique accommodations found: {len(unique_accommodations)}")
    for accommodation, count in unique_accommodations.head(10).items():
        print(f"   • '{accommodation}': {count} students")
    if len(unique_accommodations) > 10:
        print(f"   ... and {len(unique_accommodations) - 10} more unique accommodations")
    
    # Check for missing accommodation data
    missing_accommodations = df["Test Accommodation"].isna().sum()
    print(f"   Missing accommodations: {missing_accommodations} students")
    
    # DF1_PR - Private Room
    print(f"\n🔍 DF1_PR (Private Room) analysis:")
    pr_mask = df["Test Accommodation"].str.contains("Private Room", case=False, na=False)
    print(f"   Students with 'Private Room' in accommodation: {pr_mask.sum()}")
    DF1_PR = df[pr_mask]
    print(f"   DF1_PR final count: {len(DF1_PR)}")
    
    # DF2_WS - Workstation needs
    print(f"\n🔍 DF2_WS (Workstation needs) analysis:")
    ws_mask = df["Test Accommodation"].str.contains(r"Read and Write|MS Word|Kurzweil", case=False, na=False)
    print(f"   Students with workstation keywords: {ws_mask.sum()}")
    ws_not_pr_mask = ws_mask & ~df.index.isin(DF1_PR.index)
    print(f"   Students with workstation needs (excluding PR): {ws_not_pr_mask.sum()}")
    DF2_WS = df[ws_not_pr_mask]
    print(f"   DF2_WS final count: {len(DF2_WS)}")
    
    # DF3_HAD - Height Adjustable Desks
    print(f"\n🔍 DF3_HAD (Height Adjustable) analysis:")
    print(f"   'Requires Adjustable' column exists: {'Requires Adjustable' in df.columns}")
    if 'Requires Adjustable' in df.columns:
        had_count = df["Requires Adjustable"].sum()
        print(f"   Students requiring adjustable desks: {had_count}")
        # Debug what's in the Test Accommodation that might contain height adjustable
        height_adj_mask = df["Test Accommodation"].str.contains("Height Adjustable", case=False, na=False)
        print(f"   Students with 'Height Adjustable' in accommodation: {height_adj_mask.sum()}")
        DF3_HAD = df[df["Requires Adjustable"]]
        print(f"   DF3_HAD final count: {len(DF3_HAD)}")
    else:
        print("   ❌ 'Requires Adjustable' column not found!")
        DF3_HAD = pd.DataFrame()
    
    # DF4_SCRIBE - Scribe needs
    print(f"\n🔍 DF4_SCRIBE (Scribe needs) analysis:")
    scribe_mask = df["Test Accommodation"].str.contains("Scribe", case=False, na=False)
    print(f"   Students with 'Scribe' in accommodation: {scribe_mask.sum()}")
    DF4_SCRIBE = df[scribe_mask]
    print(f"   DF4_SCRIBE final count: {len(DF4_SCRIBE)}")
    
    # DF6_FINAL - Final exams
    print(f"\n🔍 DF6_FINAL (Final exams) analysis:")
    final_mask = df["Test Accommodation"].str.contains("Final", case=False, na=False)
    print(f"   Students with 'Final' in accommodation: {final_mask.sum()}")
    DF6_FINAL = df[final_mask]
    print(f"   DF6_FINAL final count: {len(DF6_FINAL)}")
    
    # DF7_ES - Evening Students
    print(f"\n🔍 DF7_ES (Evening Students) analysis:")
    print(f"   End Time column type: {df['End Time'].dtype}")
    print(f"   Begin Time column type: {df['Begin Time'].dtype}")
    print(f"   Class Time column type: {df['Class Time'].dtype}")
    
    # Check time analysis
    end_time_22 = df["End Time"].apply(lambda t: t.hour == 22 if hasattr(t, 'hour') else False)
    print(f"   Students with End Time = 22:00: {end_time_22.sum()}")
    
    begin_before_class = df["Begin Time"] < df["Class Time"]
    print(f"   Students with Begin Time < Class Time: {begin_before_class.sum()}")
    
    es_mask = end_time_22 & begin_before_class
    print(f"   Students meeting both ES criteria: {es_mask.sum()}")
    DF7_ES = df[es_mask]
    print(f"   DF7_ES final count: {len(DF7_ES)}")
    
    # DF5_MAIN - Main group
    print(f"\n🔍 DF5_MAIN (Main group) analysis:")
    excluded_indices = pd.concat([DF1_PR, DF2_WS]).index
    print(f"   Students excluded (PR + WS): {len(excluded_indices)}")
    main_mask = ~df.index.isin(excluded_indices)
    print(f"   Students remaining for main group: {main_mask.sum()}")
    DF5_MAIN = df[main_mask]
    print(f"   DF5_MAIN final count: {len(DF5_MAIN)}")
    
    # DF9_CLASSROOMS - Classroom cohort
    print(f"\n🔍 DF9_CLASSROOMS (Classroom cohort) analysis:")
    excluded_indices_all = pd.concat([DF1_PR, DF2_WS, DF5_MAIN]).index
    print(f"   Students excluded (PR + WS + MAIN): {len(excluded_indices_all)}")
    classroom_mask = ~df.index.isin(excluded_indices_all)
    print(f"   Students remaining for classrooms: {classroom_mask.sum()}")
    DF9_CLASSROOMS = df[classroom_mask]
    print(f"   DF9_CLASSROOMS final count: {len(DF9_CLASSROOMS)}")
    
    print(f"\n📊 Summary of all cohorts:")
    print(f"   DF1_PR: {len(DF1_PR)} students")
    print(f"   DF2_WS: {len(DF2_WS)} students")
    print(f"   DF3_HAD: {len(DF3_HAD)} students")
    print(f"   DF4_SCRIBE: {len(DF4_SCRIBE)} students")
    print(f"   DF5_MAIN: {len(DF5_MAIN)} students")
    print(f"   DF6_FINAL: {len(DF6_FINAL)} students")
    print(f"   DF7_ES: {len(DF7_ES)} students")
    print(f"   DF9_CLASSROOMS: {len(DF9_CLASSROOMS)} students")
    
    total_in_cohorts = len(DF1_PR) + len(DF2_WS) + len(DF3_HAD) + len(DF4_SCRIBE) + len(DF5_MAIN) + len(DF6_FINAL) + len(DF7_ES) + len(DF9_CLASSROOMS)
    print(f"   Total students in cohorts: {total_in_cohorts}")
    print(f"   Original roster size: {len(df)}")
    if total_in_cohorts != len(df):
        print(f"   ⚠️  Note: Cohorts may overlap or have gaps")

    splits = {
        "DF1_PR": DF1_PR, "DF2_WS": DF2_WS, "DF3_HAD": DF3_HAD,
        "DF4_SCRIBE": DF4_SCRIBE, "DF5_MAIN": DF5_MAIN,
        "DF6_FINAL": DF6_FINAL, "DF7_ES": DF7_ES, "DF9_CLASSROOMS": DF9_CLASSROOMS
    }
    
    if want_excel:
        print("📊 Generating cohort analysis files...")
        for name, part in splits.items():
            write_excel(part, name, template, out, "PROCESSED")

    # 8  Build seat pools based on user preferences
    seat_pools = {}
    
    if room_preferences['private_rooms']:
        seat_pools['private_rooms'] = [s for s in SEATS if SEATS[s]["type"] == "pr"]
        print(f"✅ Private rooms enabled: {len(seat_pools['private_rooms'])} seats")
    
    if room_preferences['workstations']:
        seat_pools['workstations'] = [s for s in SEATS if SEATS[s]["type"] == "ws"]
        print(f"✅ Workstations enabled: {len(seat_pools['workstations'])} seats")
    
    if room_preferences['regular_seats']:
        seat_pools['regular_seats'] = [s for s in SEATS if SEATS[s]["type"] == "reg"]
        print(f"✅ Regular seats enabled: {len(seat_pools['regular_seats'])} seats")
    
    if room_preferences['sas_offices']:
        seat_pools['sas_offices'] = [s for s in SEATS if SEATS[s]["type"] == "sas"]
        print(f"✅ SAS offices enabled: {len(seat_pools['sas_offices'])} seats")
    
    if room_preferences['sha_classrooms']:
        seat_pools['sha_classrooms'] = [s for s in SEATS if SEATS[s]["type"] == "sha"]
        print(f"✅ SHA classrooms enabled: {len(seat_pools['sha_classrooms'])} seats")

    # 9  Seat-assignment cohorts
    cohorts = {}
    
    # Assign cohorts to appropriate seat pools
    if room_preferences['private_rooms'] and len(DF1_PR) > 0:
        cohorts["DF1_PR"] = (DF1_PR, seat_pools['private_rooms'])
    
    if room_preferences['workstations'] and len(DF2_WS) > 0:
        cohorts["DF2_WS"] = (DF2_WS, seat_pools['workstations'])
    
    if room_preferences['regular_seats'] and len(DF5_MAIN) > 0:
        cohorts["DF5_MAIN"] = (DF5_MAIN, seat_pools['regular_seats'])
    
    if room_preferences['sas_offices']:
        # Find students who might need SAS offices
        remaining_students = df[~df.index.isin(pd.concat([DF1_PR, DF2_WS]).index)]
        special_accommodation_students = remaining_students[
            remaining_students["Test Accommodation"].str.contains(
                r"SAS|Special|Individual|Separate|Extra Time|Alternative|Modified", 
                case=False, na=False
            )
        ]
        
        if len(special_accommodation_students) > 0:
            cohorts["DF8_SAS"] = (special_accommodation_students, seat_pools['sas_offices'])
            print(f"📋 Including {len(special_accommodation_students)} students for SAS office assignment")
    
    if room_preferences['sha_classrooms'] and len(DF9_CLASSROOMS) > 0:
        cohorts["DF9_CLASSROOMS"] = (DF9_CLASSROOMS, seat_pools['sha_classrooms'])

    # 10  Process assignments
    assigns: dict[str, dict[str, object]] = {}
    total_assigned = 0
    total_not_assigned = 0
    
    for name, (part, pool) in cohorts.items():
        print(f"🔄 Processing {name}: {len(part)} students, {len(pool)} seats available")
        assigned, not_assigned = assign_students(part, pool)
        
        total_assigned += len(assigned)
        total_not_assigned += len(not_assigned)
        
        print(f"   ✅ Assigned: {len(assigned)}, ❌ Not assigned: {len(not_assigned)}")

        if want_excel:
            write_excel(assigned,     f"{name}_ASSIGNED",     template, out, "ASSIGNED")
            write_excel(not_assigned, f"{name}_NOT_ASSIGNED", template, out, "NOT_ASSIGNED")

        # build JSON payload
        for _, row in assigned.iterrows():
            assigns[row["Test Room"]] = {
                "student_number": int(row["Student Number"]),
                "last_name":      str(row["Student Last Name"]),
                "first_name":     str(row["Student First Name"]),
                "requiresAdjust": bool(row["Requires Adjustable"]),
            }

    # 11  Save JSON files and provide detailed summary
    pathlib.Path(out, "seats.json").write_text(
        json.dumps(SEATS, indent=2), encoding="utf-8")
    pathlib.Path(out, "assigns.json").write_text(
        json.dumps(assigns, indent=2), encoding="utf-8")

    # Print detailed completion summary
    print(f"\n✅  All done! Excel reports and JSON files generated in:")
    print(f"    📁 {out}")
    print(f"\n📊 Assignment Summary:")
    print(f"    • Total students assigned: {total_assigned}")
    print(f"    • Total students not assigned: {total_not_assigned}")
    if total_assigned + total_not_assigned > 0:
        print(f"    • Assignment success rate: {(total_assigned/(total_assigned+total_not_assigned)*100):.1f}%")
    print(f"\n🏢 Seat Breakdown:")
    room_stats = {}
    for seat_code in assigns.keys():
        if 'WS' in seat_code:
            room_stats['Workstations'] = room_stats.get('Workstations', 0) + 1
        elif 'Room' in seat_code:
            room_stats['Private Rooms'] = room_stats.get('Private Rooms', 0) + 1
        elif 'SAS' in seat_code:
            room_stats['SAS Offices'] = room_stats.get('SAS Offices', 0) + 1
        elif 'SHA' in seat_code:
            room_stats['SHA Classrooms'] = room_stats.get('SHA Classrooms', 0) + 1
        elif 'Seat' in seat_code:
            room_stats['Regular Seats'] = room_stats.get('Regular Seats', 0) + 1
    
    for room_type, count in room_stats.items():
        print(f"    • {room_type}: {count} students")
    
    print(f"\n🎯 Ready for web application launch!")

# ────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()

# ─────────────────────────────── Debug helper ───────────────────────────────────────
def debug_dataframe_splits(df: pd.DataFrame) -> None:
    """Standalone function to debug why certain cohorts are empty."""
    print("🔍 DEBUGGING DATAFRAME SPLITS")
    print("=" * 50)
    
    print(f"📋 Total students in roster: {len(df)}")
    print(f"📋 Column names: {list(df.columns)}")
    
    # Check if required columns exist
    required_columns = ["Test Accommodation", "Begin Time", "End Time", "Class Time"]
    for col in required_columns:
        if col in df.columns:
            print(f"✅ '{col}' column exists")
        else:
            print(f"❌ '{col}' column MISSING!")
    
    if "Test Accommodation" in df.columns:
        print(f"\n🔍 Test Accommodation analysis:")
        print(f"   Non-null accommodations: {df['Test Accommodation'].notna().sum()}")
        print(f"   Null accommodations: {df['Test Accommodation'].isna().sum()}")
        
        # Show sample accommodations
        sample_accommodations = df["Test Accommodation"].dropna().head(10)
        print(f"   Sample accommodations:")
        for i, acc in enumerate(sample_accommodations):
            print(f"     {i+1}. '{acc}'")
    
    # Check Requires Adjustable column
    if "Requires Adjustable" in df.columns:
        print(f"\n🔍 'Requires Adjustable' column:")
        print(f"   True: {df['Requires Adjustable'].sum()}")
        print(f"   False: {(~df['Requires Adjustable']).sum()}")
        print(f"   Null: {df['Requires Adjustable'].isna().sum()}")
    else:
        print(f"\n❌ 'Requires Adjustable' column not found!")
        
    # Test time columns
    time_columns = ["Begin Time", "End Time", "Class Time"]
    for col in time_columns:
        if col in df.columns:
            print(f"\n🔍 '{col}' column:")
            print(f"   Type: {df[col].dtype}")
            print(f"   Non-null: {df[col].notna().sum()}")
            print(f"   Sample values: {df[col].dropna().head(3).tolist()}")
    
    print("\n" + "=" * 50)