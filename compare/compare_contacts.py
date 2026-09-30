"""
compare_contacts.py

Compares the master Contact List against the MYCF and Talent List exports
to find people in MYCF / Talent List that are NOT yet in the contact list.
Designed to run both locally and inside the GitHub Action in this repo.

HOW MATCHING WORKS
-------------------
1. Email match (case-insensitive, exact) -- used whenever both sides have an
   email address. This is the most reliable match, and will become more
   useful as MYCF / Talent List get more emails filled in over time.
2. Name match (exact) -- if no email is available on one/both sides, we
   fall back to matching on normalized full name (lowercased, trimmed,
   extra whitespace collapsed).
3. Fuzzy name match ("Possible Match (Review)") -- catches likely-same-person
   pairs that (1) and (2) miss, e.g. "Sara A." vs "Sara Aldoas", or a name
   with a typo/spelling variant. We only compare people who share the same
   first name (to keep it fast and avoid nonsense matches), then flag a pair
   when either: the last name of one is a single initial that matches the
   start of the other's last name (very high confidence -- the "Sara A."
   case), or the two LAST names are >= 82% similar by character overlap
   (comparing last names, not full names, avoids false positives from a
   shared first name inflating the score). These are NOT auto-merged --
   they land in the Overview tab as "Possible Match (Review)" with a
   confidence score, for manual review.
4. Anything left over is "New".

This is a heuristic, not a guarantee -- always sanity-check the "Possible
Match" rows before treating them as duplicates.

WHAT COUNTS AS "THE CONTACT LIST"
----------------------------------
Only the "new total list" sheet of the contact list workbook is treated as
the current, active set of contacts. The "removed contacts" sheet is
intentionally excluded (those were deliberately removed) -- change
INCLUDE_REMOVED_CONTACTS below to True if you'd rather include them too.

FILE DISCOVERY (for the GitHub Action)
----------------------------------------
File names in this repo have varied over time (spaces vs underscores,
"-2"/"-3" suffixes left over from re-downloads, date stamps, etc). Rather
than hardcode an exact file name, this script searches the repo for the
right file by keyword:
  - Contact list -> file name contains both "contact" and "list"
  - MYCF         -> file name contains "mycf"
  - Talent List  -> file name contains "talent"
(case-insensitive; the generated report itself is excluded from the search
by name so re-running the action never treats last run's report as input).

If more than one file matches a category (e.g. an old copy wasn't cleaned
up), the shortest / most recently modified name is preferred and a warning
is printed -- pass --contacts/--mycf/--talent explicitly to remove the
ambiguity.

USAGE
-----
    # Explicit paths:
    python3 scripts/compare_contacts.py \
        --contacts "Contact list 2026.xlsx" \
        --mycf "MYCF.xlsx" \
        --talent "Talent List.xlsx" \
        --output "Contact_Comparison_Report.xlsx"

    # Auto-discover inputs anywhere under the repo (what the Action uses):
    python3 scripts/compare_contacts.py --repo-root . --output "Contact_Comparison_Report.xlsx"
"""

import argparse
import ast
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

INCLUDE_REMOVED_CONTACTS = False  # set True to also treat "removed contacts" as known
FUZZY_THRESHOLD = 0.82            # character-similarity cutoff for a "Possible Match"

# Directories never worth searching for source data.
IGNORED_DIR_NAMES = {".git", ".github", "node_modules", "__pycache__"}


# --------------------------------------------------------------------------
# File discovery
# --------------------------------------------------------------------------

def discover_file(repo_root, must_include_all, exclude_substrings=()):
    """Find an .xlsx file under repo_root whose (lowercased) name contains
    every string in must_include_all and none of exclude_substrings."""
    candidates = []
    for path in Path(repo_root).rglob("*.xlsx"):
        if any(part in IGNORED_DIR_NAMES or part.startswith(".") for part in path.parts):
            continue
        name = path.name.lower()
        if any(x in name for x in exclude_substrings):
            continue
        if all(k in name for k in must_include_all):
            candidates.append(path)

    if not candidates:
        return None
    if len(candidates) > 1:
        print(
            f"WARNING: multiple files match {must_include_all!r}: "
            f"{[str(c) for c in candidates]} -- using the shortest name. "
            "Pass an explicit --contacts/--mycf/--talent flag to remove ambiguity.",
            file=sys.stderr,
        )
        candidates.sort(key=lambda p: len(p.name))
    return candidates[0]


def resolve_inputs(args):
    repo_root = args.repo_root
    contacts = Path(args.contacts) if args.contacts else discover_file(
        repo_root, must_include_all=["contact", "list"], exclude_substrings=["comparison", "report"]
    )
    mycf = Path(args.mycf) if args.mycf else discover_file(repo_root, must_include_all=["mycf"])
    talent = Path(args.talent) if args.talent else discover_file(repo_root, must_include_all=["talent"])

    missing = [label for label, path in [("contact list", contacts), ("MYCF", mycf), ("Talent List", talent)] if not path]
    if missing:
        raise SystemExit(
            f"Could not find a file for: {', '.join(missing)}. "
            "Pass it explicitly with --contacts/--mycf/--talent, or check the file naming."
        )
    print(f"Using contact list: {contacts}")
    print(f"Using MYCF:         {mycf}")
    print(f"Using Talent List:  {talent}")
    return contacts, mycf, talent


# --------------------------------------------------------------------------
# Normalization helpers
# --------------------------------------------------------------------------

def norm_email(value):
    if value is None or pd.isna(value):
        return None
    value = str(value).strip().lower()
    return value if value and "@" in value else None


def norm_name(value):
    if value is None or pd.isna(value):
        return ""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^a-z0-9\s]", "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def parse_emails_field(value):
    """professionalEmails comes in as a stringified Python list, e.g.
    "['a@b.com', 'c@d.com']", or is empty/NaN. Returns a list of emails."""
    if value is None or pd.isna(value):
        return []
    value = str(value).strip()
    if not value:
        return []
    if value.startswith("["):
        try:
            parsed = ast.literal_eval(value)
            return [norm_email(e) for e in parsed if norm_email(e)]
        except (ValueError, SyntaxError):
            pass
    e = norm_email(value)
    return [e] if e else []


# --------------------------------------------------------------------------
# Loaders — one per source file, each returns a tidy dataframe with the
# common columns: name, email, company, source
# --------------------------------------------------------------------------

def load_contacts(path):
    sheets = ["new total list"] + (["removed contacts"] if INCLUDE_REMOVED_CONTACTS else [])
    frames = []
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame({
            "name": (df.get("Name1", "").fillna("") + " " + df.get("Name 2", "").fillna("")).str.strip(),
            "email": df.get("mail"),
            "company": df.get("company"),
            "source": sheet,
        })
        frames.append(out)
    combined = pd.concat(frames, ignore_index=True)
    combined = combined[combined["name"].str.strip().ne("") | combined["email"].notna()]
    combined["name_key"] = combined["name"].map(norm_name)
    combined["email_key"] = combined["email"].map(norm_email)
    return combined.reset_index(drop=True)


def load_person_export(path, sheet_names, source_label_col="source_list"):
    """Loader shared by MYCF and Talent List. Column layouts have changed
    over time -- older exports only had a 'professionalEmails' field
    (usually empty, formatted as a stringified Python list); newer exports
    have a proper single-value 'Email' column. We prefer 'Email' when
    present and fall back to 'professionalEmails' for older files."""
    frames = []
    for sheet in sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]

        if "Email" in df.columns:
            email_series = df["Email"].map(norm_email)
            all_emails_series = email_series
        elif "professionalEmails" in df.columns:
            parsed = df["professionalEmails"].map(parse_emails_field)
            email_series = parsed.map(lambda lst: lst[0] if lst else None)
            all_emails_series = parsed.map(lambda lst: "; ".join(lst) if lst else None)
        else:
            email_series = pd.Series([None] * len(df))
            all_emails_series = email_series

        out = pd.DataFrame({
            "name": (df.get("firstName", "").fillna("") + " " + df.get("lastName", "").fillna("")).str.strip(),
            "email": email_series,
            "all_emails": all_emails_series,
            "company": df.get("companyName"),
            "headline": df.get("linkedinHeadline"),
            "source": df.get(source_label_col, sheet).fillna(sheet),
        })
        frames.append(out)
    combined = pd.concat(frames, ignore_index=True)
    combined["name_key"] = combined["name"].map(norm_name)
    combined["email_key"] = combined["email"].map(norm_email)
    return combined.reset_index(drop=True)


# --------------------------------------------------------------------------
# Matching
# --------------------------------------------------------------------------

def _valid_str(value):
    """True only for a real, non-empty string key -- guards against NaN
    (which is truthy in Python!) leaking through pandas' object columns
    and causing every NaN-vs-NaN 'email' to look like a match."""
    return isinstance(value, str) and value != ""


def _fuzzy_candidate(name_key, contacts_by_first_token):
    """Look for a likely-same-person match among contacts sharing the same
    first name. Returns (matched_contact_name, confidence) or None."""
    tokens = name_key.split()
    if not tokens:
        return None
    candidates = contacts_by_first_token.get(tokens[0], [])
    best_name, best_score = None, 0.0
    for c_name_key, c_name in candidates:
        if c_name_key == name_key:
            continue
        c_tokens = c_name_key.split()
        score = None
        if len(tokens) >= 2 and len(c_tokens) >= 2:
            last, c_last = tokens[-1], c_tokens[-1]
            if (len(last) == 1 and c_last.startswith(last)) or (len(c_last) == 1 and last.startswith(c_last)):
                score = 0.95
            else:
                # Compare LAST NAMES only -- comparing full names inflates
                # the score with the shared first name.
                score = SequenceMatcher(None, last, c_last).ratio()
        else:
            score = SequenceMatcher(None, name_key, c_name_key).ratio()
        if score > best_score:
            best_score, best_name = score, c_name
    if best_name and best_score >= FUZZY_THRESHOLD:
        return best_name, round(best_score, 2)
    return None


def match_against_contacts(people_df, contacts_df):
    email_to_contact = {}
    for _, row in contacts_df.iterrows():
        if _valid_str(row["email_key"]):
            email_to_contact.setdefault(row["email_key"], row["name"])

    name_to_contact = {}
    contacts_by_first_token = {}
    for _, row in contacts_df.iterrows():
        if _valid_str(row["name_key"]):
            name_to_contact.setdefault(row["name_key"], row["name"] or row["email"])
            first_token = row["name_key"].split()[0]
            contacts_by_first_token.setdefault(first_token, []).append((row["name_key"], row["name"]))

    statuses, matched_on, matched_contact, confidence = [], [], [], []
    for _, row in people_df.iterrows():
        if _valid_str(row["email_key"]) and row["email_key"] in email_to_contact:
            statuses.append("Already in Contact List")
            matched_on.append("Email")
            matched_contact.append(email_to_contact[row["email_key"]])
            confidence.append(1.0)
        elif _valid_str(row["name_key"]) and row["name_key"] in name_to_contact:
            statuses.append("Already in Contact List")
            matched_on.append("Name")
            matched_contact.append(name_to_contact[row["name_key"]])
            confidence.append(1.0)
        else:
            fuzzy = _fuzzy_candidate(row["name_key"], contacts_by_first_token) if _valid_str(row["name_key"]) else None
            if fuzzy:
                statuses.append("Possible Match (Review)")
                matched_on.append("Fuzzy Name")
                matched_contact.append(fuzzy[0])
                confidence.append(fuzzy[1])
            else:
                statuses.append("New")
                matched_on.append("")
                matched_contact.append("")
                confidence.append(None)

    people_df = people_df.copy()
    people_df["Status"] = statuses
    people_df["Matched On"] = matched_on
    people_df["Matched Contact"] = matched_contact
    people_df["Match Confidence"] = confidence
    return people_df


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Arial")
BODY_FONT = Font(name="Arial")


def style_header(ws, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"


def autofit(ws, df, max_width=45):
    for i, col in enumerate(df.columns, start=1):
        if len(df):
            lengths = df[col].map(lambda v: len(str(v)) if v is not None else 0)
            content_max = int(lengths.max())
        else:
            content_max = 12
        width = min(max_width, max(12, content_max, len(str(col)) + 2))
        ws.column_dimensions[get_column_letter(i)].width = width


def write_simple_tab(writer, df, sheet_name, display_cols, rename=None):
    out = df[display_cols].rename(columns=rename or {})
    out.to_excel(writer, sheet_name=sheet_name, index=False)
    return out


def build_workbook(contacts, mycf, talent, output_path):
    overview_cols = {
        "source": "Source List",
        "name": "Name",
        "email": "Email",
        "all_emails": "All Emails Found",
        "company": "Company",
        "headline": "LinkedIn Headline",
        "Status": "Status",
        "Matched On": "Matched On",
        "Matched Contact": "Matched Contact Name",
        "Match Confidence": "Match Confidence",
    }
    overview = pd.concat([mycf, talent], ignore_index=True)
    overview = overview[list(overview_cols.keys())].rename(columns=overview_cols)
    status_order = {"New": 0, "Possible Match (Review)": 1, "Already in Contact List": 2}
    overview["_sort"] = overview["Status"].map(status_order)
    overview = overview.sort_values(["_sort", "Source List", "Name"]).drop(columns="_sort")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        contacts_out = write_simple_tab(
            writer, contacts, "Contacts",
            ["name", "email", "company", "source"],
            {"name": "Name", "email": "Email", "company": "Company", "source": "Source Sheet"},
        )
        mycf_out = write_simple_tab(
            writer, mycf, "MYCF",
            ["name", "email", "all_emails", "company", "headline", "source"],
            {"name": "Name", "email": "Email", "all_emails": "All Emails Found",
             "company": "Company", "headline": "LinkedIn Headline", "source": "Source List"},
        )
        talent_out = write_simple_tab(
            writer, talent, "Talent List",
            ["name", "email", "all_emails", "company", "headline", "source"],
            {"name": "Name", "email": "Email", "all_emails": "All Emails Found",
             "company": "Company", "headline": "LinkedIn Headline", "source": "Source List"},
        )
        overview.to_excel(writer, sheet_name="Overview", index=False, startrow=8)

    wb = load_workbook(output_path)

    for sheet_name, df in [("Contacts", contacts_out), ("MYCF", mycf_out), ("Talent List", talent_out)]:
        ws = wb[sheet_name]
        style_header(ws, len(df.columns))
        autofit(ws, df)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.font = BODY_FONT

    # --- Overview tab: summary block (formulas, not hardcoded numbers) + table ---
    ws = wb["Overview"]
    n_contacts = len(contacts_out)
    n_mycf = len(mycf_out)
    n_talent = len(talent_out)
    header_row = 9
    last_row = header_row + len(overview)

    status_col_idx = list(overview.columns).index("Status") + 1
    source_col_idx = list(overview.columns).index("Source List") + 1
    status_col = get_column_letter(status_col_idx)
    source_col = get_column_letter(source_col_idx)

    ws["A1"] = "Contact List Comparison — Overview"
    ws["A1"].font = Font(bold=True, size=14, name="Arial")

    labels = [
        ("A3", "Total contacts (existing list):", f"={n_contacts}"),
        ("A4", "Total MYCF records:", f"={n_mycf}"),
        ("A5", "Total Talent List records:", f"={n_talent}"),
        ("A6", "MYCF/Talent records already in Contact List:",
         f'=COUNTIF({status_col}{header_row}:{status_col}{last_row},"Already in Contact List")'),
        ("A7", "Possible matches to review (likely same person, different spelling):",
         f'=COUNTIF({status_col}{header_row}:{status_col}{last_row},"Possible Match (Review)")'),
        ("A8", "MYCF/Talent records that are NEW (not in Contact List):",
         f'=COUNTIF({status_col}{header_row}:{status_col}{last_row},"New")'),
    ]
    for cell_ref, label, formula in labels:
        row = int(cell_ref[1:])
        ws[cell_ref] = label
        ws[cell_ref].font = Font(bold=True, name="Arial")
        ws.cell(row=row, column=3).value = formula
        ws.cell(row=row, column=3).font = Font(name="Arial", bold=True)

    style_header(ws, len(overview.columns))
    for col in range(1, len(overview.columns) + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = f"A{header_row + 1}"
    autofit(ws, overview)

    new_fill = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")     # orange
    review_fill = PatternFill(start_color="FFF7B2", end_color="FFF7B2", fill_type="solid")  # yellow
    for r in range(header_row + 1, last_row + 1):
        status_value = ws.cell(row=r, column=status_col_idx).value
        fill = new_fill if status_value == "New" else review_fill if status_value == "Possible Match (Review)" else None
        for c in range(1, len(overview.columns) + 1):
            cell = ws.cell(row=r, column=c)
            if fill:
                cell.fill = fill
            cell.font = Font(name="Arial")

    # --- New sheet: bar chart of New / Possible Match / Already known, per source file ---
    ws_chart = wb.create_sheet("New Contacts Chart")
    ws_chart["A1"] = "New Contacts Found — by Source"
    ws_chart["A1"].font = Font(bold=True, size=14, name="Arial")

    chart_header_row = 3
    chart_headers = ["Source File", "New", "Possible Match (Review)", "Already in Contact List"]
    for i, h in enumerate(chart_headers, start=1):
        cell = ws_chart.cell(row=chart_header_row, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    def countifs_formula(source_values, status_value):
        parts = [f"COUNTIFS(Overview!{source_col}{header_row}:{source_col}{last_row},\"{sv}\","
                 f"Overview!{status_col}{header_row}:{status_col}{last_row},\"{status_value}\")" for sv in source_values]
        return "=" + "+".join(parts)

    chart_rows = [
        ("MYCF", ["MYCF"]),
        ("Talent List", ["Direct Applications", "Open Applications"]),
    ]
    for i, (label, source_values) in enumerate(chart_rows, start=chart_header_row + 1):
        ws_chart.cell(row=i, column=1, value=label).font = Font(name="Arial", bold=True)
        ws_chart.cell(row=i, column=2, value=countifs_formula(source_values, "New")).font = BODY_FONT
        ws_chart.cell(row=i, column=3, value=countifs_formula(source_values, "Possible Match (Review)")).font = BODY_FONT
        ws_chart.cell(row=i, column=4, value=countifs_formula(source_values, "Already in Contact List")).font = BODY_FONT
    chart_data_last_row = chart_header_row + len(chart_rows)

    for col, width in zip("ABCD", [16, 10, 22, 22]):
        ws_chart.column_dimensions[col].width = width

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "New contacts found per source"
    chart.y_axis.title = "Number of people"
    chart.x_axis.title = "Source file"
    chart.style = 10

    data = Reference(ws_chart, min_col=2, max_col=4, min_row=chart_header_row, max_row=chart_data_last_row)
    cats = Reference(ws_chart, min_col=1, max_col=1, min_row=chart_header_row + 1, max_row=chart_data_last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width, chart.height = 22, 12
    ws_chart.add_chart(chart, "A8")

    wb.save(output_path)
    print(f"Saved: {output_path}")
    print(f"  Contacts: {n_contacts}")
    print(f"  MYCF: {n_mycf}  |  Talent List: {n_talent}")
    print(f"  New (not yet in contact list): {(overview['Status'] == 'New').sum()}")
    print(f"  Possible matches to review: {(overview['Status'] == 'Possible Match (Review)').sum()}")
    print(f"  Already known: {(overview['Status'] == 'Already in Contact List').sum()}")


# --------------------------------------------------------------------------
# Formula recalculation (self-contained -- no dependency on external tooling
# beyond the `soffice` binary, which the GitHub Action installs via apt)
# --------------------------------------------------------------------------

RECALC_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def recalculate_with_libreoffice(path, timeout=60):
    """Force Excel formulas to have a cached value by having LibreOffice
    open, recalculate, and re-save the file in place. openpyxl writes
    formulas WITHOUT a cached value, which some viewers (and pandas'
    data_only mode) show as blank until something recalculates them --
    Excel and Google Sheets do this automatically on open, but this makes
    the artifact stored in the repo self-contained and pre-verified.
    Skips gracefully (with a warning) if LibreOffice isn't installed."""
    if not shutil.which("soffice"):
        print("NOTE: 'soffice' (LibreOffice) not found -- skipping formula recalculation. "
              "Excel/Google Sheets will still compute the formulas fine when the file is opened.",
              file=sys.stderr)
        return

    abs_path = str(Path(path).resolve())
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile_dir:
        macro_dir = Path(profile_dir) / "user" / "basic" / "Standard"
        # Bootstrap a throwaway LibreOffice profile, then drop our macro into it.
        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init",
             f"-env:UserInstallation={Path(profile_dir).as_uri()}"],
            capture_output=True, timeout=timeout, check=False,
        )
        macro_dir.mkdir(parents=True, exist_ok=True)
        (macro_dir / "Module1.xba").write_text(RECALC_MACRO)

        cmd = [
            "soffice", "--headless", "--norestore",
            f"-env:UserInstallation={Path(profile_dir).as_uri()}",
            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
            abs_path,
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
            if result.returncode != 0:
                print(f"WARNING: LibreOffice recalculation failed ({result.stderr.strip()}); "
                      "the file was still saved, just without cached formula values.", file=sys.stderr)
            else:
                print("Recalculated formulas with LibreOffice.")
        except subprocess.TimeoutExpired:
            print("WARNING: LibreOffice recalculation timed out; the file was still saved, "
                  "just without cached formula values.", file=sys.stderr)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--contacts", help="Path to the contact list workbook (auto-discovered if omitted)")
    parser.add_argument("--mycf", help="Path to the MYCF workbook (auto-discovered if omitted)")
    parser.add_argument("--talent", help="Path to the Talent List workbook (auto-discovered if omitted)")
    parser.add_argument("--repo-root", default=".", help="Root directory to search when auto-discovering files")
    parser.add_argument("--output", default="Contact_Comparison_Report.xlsx", help="Path to write the report to")
    parser.add_argument("--no-recalc", action="store_true", help="Skip LibreOffice formula recalculation")
    args = parser.parse_args()

    contacts_path, mycf_path, talent_path = resolve_inputs(args)

    contacts = load_contacts(contacts_path)
    mycf = load_person_export(mycf_path, ["MYCF"])
    talent = load_person_export(talent_path, ["Direct Applications", "Open Applications"])

    mycf = match_against_contacts(mycf, contacts)
    talent = match_against_contacts(talent, contacts)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(contacts, mycf, talent, str(output_path))

    if not args.no_recalc:
        recalculate_with_libreoffice(output_path)


if __name__ == "__main__":
    main()
