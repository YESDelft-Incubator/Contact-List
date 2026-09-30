"""
Export PhantomBuster lead data to Excel workbooks for Tableau.

Sources
    Every sheet reads a Phantom's own result file, addressed by agent_id.
    SheetConfig still accepts list_id to read the org-storage leads database
    instead, which no sheet currently uses.

Update model
    --rebuild   write each workbook from scratch, discarding what is on disk.
                Needed once, because the result files carry a different schema
                than the leads database did.
    default     upsert into the existing workbook:
                    a key not seen before is appended,
                    a key seen again has its non-blank fields refreshed,
                    a key that disappears upstream is kept, with last_seen frozen.

Usage
    python export_leads.py --inspect     # fetch, report the schema, write nothing
    python export_leads.py --rebuild     # first run after the source switch
    python export_leads.py               # every run after that
    python export_leads.py --mock        # read fixtures/, no network calls

Note for CI: the upsert reads the workbooks that are already on disk, so the
scheduled job must commit them back to the repository. On a fresh runner with
nothing committed, every run behaves like --rebuild and no history survives.
"""

import argparse
import hashlib
import json
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd
import requests

# ─── CONFIG ──────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent
MAPPINGS_FILE = SCRIPT_DIR / "Mappings Final.xlsx"
FIXTURE_DIR = SCRIPT_DIR / "fixtures"
MAX_CELL_LENGTH = 30000

LEADS_API = "https://api.phantombuster.com/api/v2/org-storage/leads/by-list"
AGENT_API = "https://api.phantombuster.com/api/v2/agents/fetch"
S3_BASE = "https://phantombuster.s3.amazonaws.com"

KEY = "lead_key"
FIRST_SEEN = "first_seen"
LAST_SEEN = "last_seen"
BOOKKEEPING = (KEY, FIRST_SEEN, LAST_SEEN)

# Candidate columns holding the LinkedIn profile URL, tried in order.
# The leads database and the result files do not agree on the name.
DEFAULT_KEY_COLUMNS = (
    "profileUrl",
    "linkedinUrl",
    "linkedinProfileUrl",
    "profileLink",
    "linkedinProfile",
)


@dataclass
class SheetConfig:
    """One PhantomBuster source, written to one Excel sheet.

    Set exactly one of agent_id (read the Phantom's result file) or
    list_id (read the org-storage leads database).
    """

    sheet_name: str  # Excel sheet name, max 31 chars, no [ ] : * ? / \
    agent_id: str = ""
    list_id: str = ""
    # Leave blank to auto-detect. Set only if the Phantom writes a renamed file.
    result_file: str = ""
    # {source: canonical}. Fill from --inspect output.
    rename: dict = field(default_factory=dict)
    key_columns: tuple = DEFAULT_KEY_COLUMNS
    standardise_degree: bool = True
    standardise_location: bool = True

    def __post_init__(self):
        if bool(self.agent_id) == bool(self.list_id):
            raise ValueError(
                f"{self.sheet_name}: set exactly one of agent_id or list_id."
            )

    @property
    def source_id(self):
        return self.agent_id or self.list_id

    @property
    def source_kind(self):
        return "agent" if self.agent_id else "list"


@dataclass
class WorkbookConfig:
    filename: str
    sheets: list


WORKBOOKS = [
    WorkbookConfig(
        filename="YES!Delft Linkedin Followers.xlsx",
        sheets=[
            SheetConfig(
                sheet_name="Blad1",
                agent_id="4167655640536198",
                # rename={"<result file name>": "linkedinSchoolDegree"},
            ),
        ],
    ),
    WorkbookConfig(
        filename="Talent List.xlsx",
        sheets=[
            SheetConfig(
                sheet_name="Direct Applications",
                agent_id="3777858886218670",
                # rename={"<result file name>": "linkedinSchoolDegree"},
            ),
            SheetConfig(
                sheet_name="Open Applications",
                agent_id="1498687002457604",
                # rename={"<result file name>": "linkedinSchoolDegree"},
            ),
        ],
    ),
    WorkbookConfig(
        filename="MYCF.xlsx",
        sheets=[
            SheetConfig(
                sheet_name="MYCF",
                agent_id="3463420412076036",
                # rename={"<result file name>": "linkedinSchoolDegree"},
            ),
        ],
    ),
]


# ─── FETCH ───────────────────────────────────────────────
def get_api_key():
    api_key = os.getenv("PHANTOMBUSTER_API_TOKEN")
    if not api_key:
        raise EnvironmentError(
            "Missing environment variable: PHANTOMBUSTER_API_TOKEN.\n"
            "Set it in your repository secrets or local environment."
        )
    return api_key


def _headers(api_key):
    return {
        "accept": "application/json",
        "content-type": "application/json",
        "X-Phantombuster-Key": api_key,
    }


def _configured_result_names(meta):
    """Pull the result file name out of the agent's saved argument, if it set one."""
    raw = meta.get("argument")
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return []
    if not isinstance(raw, dict):
        return []

    names = []
    for key in ("csvName", "resultFileName", "fileName", "outputFileName", "name"):
        value = raw.get(key)
        if isinstance(value, str) and value.strip():
            stem = value.strip()
            if stem.lower().endswith((".json", ".csv")):
                names.append(stem)
            else:
                names.extend([f"{stem}.json", f"{stem}.csv"])
    return names


def _candidate_result_files(cfg, meta):
    """Ordered, de-duplicated list of file names to try in the agent's S3 folder."""
    candidates = []
    if cfg.result_file:
        candidates.append(cfg.result_file)
    candidates.extend(_configured_result_names(meta))
    candidates.extend(["result.json", "result.csv"])
    return list(dict.fromkeys(candidates))


def _parse_result(response, name, sheet_name):
    if name.lower().endswith(".csv"):
        from io import StringIO

        return pd.read_csv(StringIO(response.text)).to_dict("records")

    data = response.json()
    if not isinstance(data, list):
        raise ValueError(
            f"{sheet_name}: expected a list of records in {name}, "
            f"got {type(data).__name__}"
        )
    return data


def fetch_from_agent(cfg, api_key):
    """Read the Phantom's result file. Carries every column the Phantom emits,
    including any input columns kept via 'Names of the columns to keep'."""
    meta = requests.get(
        AGENT_API, params={"id": cfg.agent_id}, headers=_headers(api_key), timeout=60
    )
    meta.raise_for_status()
    meta = meta.json()

    absent = [k for k in ("orgS3Folder", "s3Folder") if not meta.get(k)]
    if absent:
        raise ValueError(
            f"{cfg.sheet_name}: agent {cfg.agent_id} returned no {absent}. "
            "Has this Phantom ever completed a run?"
        )

    folder = f"{S3_BASE}/{meta['orgS3Folder']}/{meta['s3Folder']}"
    candidates = _candidate_result_files(cfg, meta)
    tried = []

    for name in candidates:
        response = requests.get(f"{folder}/{name}", timeout=300)
        # S3 hides a missing object behind 403 when ListBucket is withheld,
        # so 403 here means "not this name", not "no access".
        if response.status_code in (403, 404):
            tried.append(f"{name} ({response.status_code})")
            continue
        response.raise_for_status()
        print(f"    agent '{meta.get('name', '?')}' -> {name}")
        return _parse_result(response, name, cfg.sheet_name)

    raise FileNotFoundError(
        f"{cfg.sheet_name}: no result file found for agent {cfg.agent_id} "
        f"('{meta.get('name', '?')}').\n"
        f"  folder:  {folder}\n"
        f"  tried:   {', '.join(tried)}\n"
        f"  argument keys: {sorted(_argument_keys(meta))}\n"
        "Open the Phantom's Behavior step, read the result file name, and set "
        "result_file on this SheetConfig."
    )


def _argument_keys(meta):
    raw = meta.get("argument")
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return []
    return list(raw) if isinstance(raw, dict) else []


def fetch_from_list(cfg, api_key):
    """Read the org-storage leads database. Fixed schema, custom columns dropped."""
    response = requests.post(
        f"{LEADS_API}/{cfg.list_id}", headers=_headers(api_key), timeout=300
    )
    response.raise_for_status()
    data = response.json()
    if not isinstance(data, list):
        raise ValueError(f"Unexpected API response format for {cfg.list_id}: {data}")
    return data


def fetch_live(cfg, api_key):
    if cfg.source_kind == "agent":
        return fetch_from_agent(cfg, api_key)
    return fetch_from_list(cfg, api_key)


def fetch_mock(cfg, api_key=None):
    path = FIXTURE_DIR / f"{cfg.source_id}.json"
    if not path.exists():
        raise FileNotFoundError(
            f"No fixture for {cfg.source_kind} {cfg.source_id} at {path}"
        )
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ─── MAPPINGS ────────────────────────────────────────────
def load_mapping(sheet, key_col, value_col):
    mappings_df = pd.read_excel(MAPPINGS_FILE, sheet_name=sheet)
    mappings_df.columns = mappings_df.columns.str.strip()
    missing = {key_col, value_col} - set(mappings_df.columns)
    if missing:
        raise KeyError(
            f"Columns {sorted(missing)} not found in sheet '{sheet}' of "
            f"{MAPPINGS_FILE.name}. Found: {list(mappings_df.columns)}"
        )
    pairs = mappings_df[[key_col, value_col]].dropna()
    return dict(zip(pairs[key_col], pairs[value_col]))


def apply_mapping(df, column, mapping, label, unmapped_log):
    """Map a column in place, keep the raw value, record misses."""
    if column not in df.columns:
        print(f"    '{column}' not present, skipping {label} standardisation.")
        return df

    raw_col = f"{column}_raw"
    if raw_col not in df.columns:
        df[raw_col] = df[column]

    mask = df[column].notna() & (df[column] != "") & ~df[column].isin(mapping)
    misses = Counter(df.loc[mask, column])
    df[column] = df[column].apply(lambda x: mapping.get(x, x))

    if misses:
        total = sum(misses.values())
        print(f"    {label}: {total} rows across {len(misses)} unmapped values")
        for value, n in misses.most_common(5):
            print(f"        {n:>5}  {value}")
        if len(misses) > 5:
            print(f"        ... and {len(misses) - 5} more, see unmapped_report.csv")
        unmapped_log.extend(
            {"field": label, "value": v, "rows": n} for v, n in misses.items()
        )
    return df


def split_location(df):
    """Derive City and Country from the standardised location string."""
    if "location" not in df.columns:
        print("    'location' not present, skipping city/country split.")
        return df

    def parse(value):
        if not isinstance(value, str) or not value.strip():
            return "", ""
        parts = [p.strip() for p in value.split(",") if p.strip()]
        if not parts:
            return "", ""
        if len(parts) == 1:
            return "", parts[0]  # unchanged from the original script
        return parts[0], parts[-1]  # 2 or more: first is city, last is country

    parsed = df["location"].apply(parse)
    df["City"] = [p[0] for p in parsed]
    df["Country"] = [p[1] for p in parsed]

    blank = int((df["Country"] == "").sum())
    if blank:
        print(f"    {blank} of {len(df)} rows have no Country after the split")
    return df


# ─── IDENTITY ────────────────────────────────────────────
_LI_HOST = re.compile(r"^(?:https?://)?(?:[a-z]{2,3}\.)?(?:www\.)?linkedin\.com", re.I)


def normalise_linkedin_url(value):
    """Reduce the many shapes of a LinkedIn URL to one comparable string."""
    if not isinstance(value, str) or not value.strip():
        return ""
    v = value.strip().split("?")[0].rstrip("/")
    return _LI_HOST.sub("linkedin.com", v).lower()


def row_fingerprint(row):
    """Fallback identity for a record with no usable profile URL."""
    payload = "|".join(f"{k}={row[k]!r}" for k in sorted(row.index) if k not in BOOKKEEPING)
    return "fp:" + hashlib.sha1(payload.encode("utf-8")).hexdigest()[:16]


def derive_key(df, cfg):
    """Attach a stable per-lead key and collapse duplicates within this batch."""
    col = next((c for c in cfg.key_columns if c in df.columns), None)
    if col is None:
        raise KeyError(
            f"{cfg.sheet_name}: no profile URL column found. Tried "
            f"{list(cfg.key_columns)}. Present: {sorted(df.columns)}\n"
            "Add the right name to key_columns on this SheetConfig."
        )

    keys = df[col].apply(normalise_linkedin_url)
    blank = keys == ""
    if blank.any():
        print(f"    {int(blank.sum())} rows have no usable '{col}', keyed by content hash")
        keys.loc[blank] = df.loc[blank].apply(row_fingerprint, axis=1)

    df = df.copy()
    df[KEY] = keys

    dupes = int(df[KEY].duplicated().sum())
    if dupes:
        print(f"    collapsed {dupes} duplicate keys within this batch, kept the last")
        df = df.drop_duplicates(KEY, keep="last")
    return df


# ─── ASSEMBLY ────────────────────────────────────────────
def truncate(df):
    """Excel refuses a cell over 32,767 characters. Applied before comparison so
    stored and incoming values are on equal footing."""
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: str(x)[:MAX_CELL_LENGTH] if isinstance(x, str) else x
        )
    return df


def prepare(leads, cfg, degree_map, location_map, unmapped_log):
    df = pd.DataFrame(leads)
    if df.empty:
        print("    no records")
        return df

    df.drop(
        columns=[c for c in ("editionsHistory", "metadata") if c in df.columns],
        inplace=True,
    )

    collisions = [
        src for src, dst in cfg.rename.items() if src in df.columns and dst in df.columns
    ]
    if collisions:
        raise ValueError(
            f"{cfg.sheet_name}: rename would overwrite existing columns {collisions}. "
            "Both the source and target names are present in the data."
        )
    df = df.rename(columns=cfg.rename)

    if cfg.standardise_degree:
        df = apply_mapping(df, "linkedinSchoolDegree", degree_map, "degree", unmapped_log)
    if cfg.standardise_location:
        df = apply_mapping(df, "location", location_map, "location", unmapped_log)
    df = split_location(df)

    df["source_list"] = cfg.sheet_name
    df = derive_key(df, cfg)
    df = truncate(df)

    # Keep the original leading column order so existing Tableau fields resolve.
    lead = [c for c in ("location", "City", "Country") if c in df.columns]
    rest = [c for c in df.columns if c not in lead]
    return df[lead + rest]


# ─── UPSERT ──────────────────────────────────────────────
def is_blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and not value.strip()


def read_existing(path, sheet_name):
    """Return the sheet as it stands on disk, or None if there is nothing usable."""
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=object)
    except ValueError:  # sheet absent from an existing workbook
        return None
    if df.empty or KEY not in df.columns:
        return None
    return df.drop_duplicates(KEY, keep="last")


def upsert(existing, incoming, today):
    """Merge a fresh batch into the stored sheet.

    New keys are appended. Keys present in both have their non-blank incoming
    fields written over the stored ones, so an upstream blank never erases data
    already collected. Keys absent from the batch survive untouched, with their
    last_seen left at whatever it was.
    """
    if existing is None:
        out = incoming.copy()
        out[FIRST_SEEN] = today
        out[LAST_SEEN] = today
        return out, {"added": len(out), "refreshed": 0, "retained": 0}

    exi = existing.set_index(KEY)
    inc = incoming.set_index(KEY)

    # Stored column order first, then anything the batch introduces.
    cols = [c for c in exi.columns if c not in BOOKKEEPING]
    cols += [c for c in inc.columns if c not in cols and c not in BOOKKEEPING]

    exi_data = exi.reindex(columns=cols)
    inc_data = inc.reindex(columns=cols)

    # Mask blanks so combine_first falls through to the stored value.
    blanks = inc_data.apply(lambda s: s.map(is_blank))
    merged = inc_data.mask(blanks).combine_first(exi_data)

    seen_now = set(inc.index)
    prev_first = exi[FIRST_SEEN] if FIRST_SEEN in exi.columns else pd.Series(dtype=object)
    prev_last = exi[LAST_SEEN] if LAST_SEEN in exi.columns else pd.Series(dtype=object)

    merged[FIRST_SEEN] = [
        today if k not in prev_first.index else prev_first[k] for k in merged.index
    ]
    merged[LAST_SEEN] = [
        today if k in seen_now else prev_last.get(k, "") for k in merged.index
    ]

    order = [k for k in exi.index if k in merged.index]
    order += [k for k in inc.index if k not in exi.index]
    merged = merged.loc[order].reset_index()

    added = len(seen_now - set(exi.index))
    return merged, {
        "added": added,
        "refreshed": len(seen_now) - added,
        "retained": len(set(exi.index) - seen_now),
    }


# ─── INSPECT ─────────────────────────────────────────────
def inspect(cfg, df, workbook_path):
    """Report what the new source gives against what the workbook already holds."""
    new_cols = set(df.columns)
    print(f"    {len(new_cols)} columns from the source")

    for required in ("linkedinSchoolDegree", "location"):
        if required not in new_cols:
            near = [c for c in sorted(new_cols) if required[:6].lower() in c.lower()]
            print(f"    MISSING '{required}'. Closest names: {near or 'none'}")
            print(f"      -> add a rename entry on {cfg.sheet_name}")

    old = read_existing(workbook_path, cfg.sheet_name)
    if old is None:
        print("    no comparable sheet on disk yet")
    else:
        old_cols = set(old.columns) - set(BOOKKEEPING)
        gone = sorted(old_cols - new_cols)
        fresh = sorted(new_cols - old_cols)
        print(f"    dropped vs current workbook ({len(gone)}): {gone}")
        print(f"    new vs current workbook ({len(fresh)}): {fresh}")

    print(f"    all columns: {sorted(new_cols)}")


# ─── MAIN ────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mock", action="store_true", help="read fixtures, no API calls")
    parser.add_argument(
        "--rebuild", action="store_true", help="discard existing workbooks and start over"
    )
    parser.add_argument(
        "--inspect", action="store_true", help="fetch and report the schema, write nothing"
    )
    args = parser.parse_args()

    if args.mock:
        fetch, api_key = fetch_mock, None
        print("MOCK MODE: reading fixtures, no API calls.\n")
    else:
        fetch, api_key = fetch_live, get_api_key()

    if args.rebuild and not args.inspect:
        print("REBUILD MODE: existing workbook contents will be discarded.\n")

    degree_map = load_mapping("Degrees", "Unique Degrees", "Standard Degree Level")
    location_map = load_mapping("Locations", "Variation", "Canonical")
    print(f"Loaded {len(degree_map)} degree and {len(location_map)} location mappings.\n")

    today = date.today().isoformat()
    unmapped_log = []
    summary = []
    pending = []  # (out_path, {sheet_name: df}) built fully before anything is written

    for wb in WORKBOOKS:
        out_path = SCRIPT_DIR / wb.filename
        print(wb.filename)
        prepared = {}

        # Read every stored sheet before opening the writer, which truncates the file.
        stored = {
            cfg.sheet_name: (None if args.rebuild else read_existing(out_path, cfg.sheet_name))
            for cfg in wb.sheets
        }

        for cfg in wb.sheets:
            print(f"  {cfg.sheet_name} ({cfg.source_kind} {cfg.source_id})")
            leads = fetch(cfg, api_key)
            print(f"    retrieved {len(leads)} records")

            df = prepare(leads, cfg, degree_map, location_map, unmapped_log)
            if df.empty:
                # Never let an empty upstream wipe a stored sheet.
                if stored[cfg.sheet_name] is not None:
                    print("    empty batch, keeping the stored sheet as is")
                    prepared[cfg.sheet_name] = stored[cfg.sheet_name]
                continue

            if args.inspect:
                inspect(cfg, df, out_path)
                continue

            merged, stats = upsert(stored[cfg.sheet_name], df, today)
            print(
                f"    {stats['added']} added, {stats['refreshed']} refreshed, "
                f"{stats['retained']} retained from earlier runs"
            )
            prepared[cfg.sheet_name] = merged
            summary.append((wb.filename, cfg.sheet_name, len(merged), len(merged.columns)))

        if args.inspect:
            print()
            continue

        if not prepared:
            print("  nothing to write, skipping file\n")
            continue

        pending.append((out_path, prepared))
        print("  prepared, queued for writing\n")

    if args.inspect:
        print("Inspect only, nothing written.")
        return

    # Every source fetched cleanly, so it is safe to touch the workbooks.
    for out_path, prepared in pending:
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            writer.book.strings_to_urls = False
            for sheet_name, df in prepared.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        print(f"wrote {out_path.name}")
    print()

    if unmapped_log:
        report = (
            pd.DataFrame(unmapped_log)
            .groupby(["field", "value"], as_index=False)["rows"]
            .sum()
            .sort_values(["field", "rows"], ascending=[True, False])
        )
        report.to_csv(SCRIPT_DIR / "unmapped_report.csv", index=False)
        print(f"Wrote unmapped_report.csv with {len(report)} values needing a mapping.\n")

    print("Summary")
    for filename, sheet, rows, cols in summary:
        print(f"  {filename:36} {sheet:22} {rows:>6} rows  {cols:>3} cols")


if __name__ == "__main__":
    main()
