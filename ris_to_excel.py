# ris_to_excel.py — Convert RIS exports into grouped Excel files

from __future__ import annotations

from pathlib import Path
import re
import pandas as pd
from urllib.parse import urlparse

# Final Excel column order
OUT_COLS = ["Title", "Year", "Index", "DOI", "Author Surname", "Author Name"]

# Domain → nice provider name mapping (used for Index column)
PROVIDER_MAP = {
    "ieeexplore.ieee.org": "IEEE Xplore",
    "sciencedirect.com": "ScienceDirect",
    "webofscience.com": "Web of Science",
    "link.springer.com": "SpringerLink",
    "springer.com": "SpringerLink",
    "onlinelibrary.wiley.com": "Wiley Online Library",
    "wiley.com": "Wiley Online Library",
    "pubmed.ncbi.nlm.nih.gov": "PubMed",
    "dl.acm.org": "ACM Digital Library",
    "acm.org": "ACM Digital Library",
    "tandfonline.com": "Taylor & Francis",
    "nature.com": "Nature",
}

# Optional grouping by filename:
# 1.(Query Name)_0-100.ris  OR  1.(Query Name)_all.ris
FNAME_RE = re.compile(r"^(?P<num>\d+)\.\((?P<query>.+)\)_.+\.ris$", re.IGNORECASE)

# Flexible RIS tag format (handles different spacing styles)
# Example: "TI  - Some Title" / "DO - 10.xxxx/xxxx"
TAG_LINE_RE = re.compile(r"^(?P<tag>[A-Z0-9]{2})\s*-\s*(?P<val>.*)$")


def clean_doi(x: str) -> str:
    # RIS exports sometimes store DOI as a doi.org URL
    x = (x or "").strip()
    return re.sub(r"^https?://doi\.org/", "", x, flags=re.IGNORECASE).strip()


def surname_from_author(author: str) -> str:
    """Extract surname from RIS author string. Supports 'Surname, Name' and 'Name Surname'."""
    author = (author or "").strip()
    if not author:
        return ""
    if "," in author:
        return author.split(",", 1)[0].strip()
    parts = author.split()
    return parts[-1].strip() if parts else ""


def year_from_any(r: dict) -> str:
    """Extract a 4-digit year from common RIS date fields (PY/Y1/DA/DP)."""
    for k in ["PY", "Y1", "DA", "DP"]:
        v = r.get(k)
        if isinstance(v, list):
            v = v[0] if v else ""
        v = (v or "").strip()
        if v:
            m = re.search(r"\b(19|20)\d{2}\b", v)
            if m:
                return m.group(0)
    return ""


def get_first_from_tags(r: dict, tags: list[str]) -> str:
    """Return the first non-empty value found among the given RIS tags."""
    for tag in tags:
        val = r.get(tag)
        if not val:
            continue
        if isinstance(val, list):
            for x in val:
                x = (x or "").strip()
                if x:
                    return x
        else:
            x = str(val).strip()
            if x:
                return x
    return ""


def parse_ris_file(path: Path) -> list[dict]:
    """
    Universal RIS parser:
    - supports repeated tags (stored as lists)
    - supports wrapped/continuation lines
    - splits entries by TY ... ER
    """
    records: list[dict] = []
    current: dict = {}
    last_tag: str | None = None

    def start_new():
        nonlocal current, last_tag
        current = {}
        last_tag = None

    def finish_one():
        nonlocal current
        if current:
            records.append(current)
        current = {}

    def add_value(tag: str, val: str):
        # RIS tags like AU can repeat; store repeats as list
        if tag in current:
            if isinstance(current[tag], list):
                current[tag].append(val)
            else:
                current[tag] = [current[tag], val]
        else:
            current[tag] = val

    start_new()

    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for raw in f:
            line = raw.rstrip("\r\n")

            m = TAG_LINE_RE.match(line)
            if not m:
                # Continuation line: append to the previous tag value
                if last_tag and line.strip():
                    if isinstance(current.get(last_tag), list):
                        current[last_tag][-1] = (str(current[last_tag][-1]) + " " + line.strip()).strip()
                    else:
                        current[last_tag] = (str(current.get(last_tag, "")) + " " + line.strip()).strip()
                continue

            tag = m.group("tag").strip()
            val = m.group("val").strip()
            last_tag = tag

            if tag == "ER":
                finish_one()
                start_new()
                continue

            add_value(tag, val)

    return records


def extract_group_header(path: Path) -> tuple[int | None, str]:
    """
    If filename matches N.(query)_something.ris, group by N.(query).
    Otherwise group by filename (FILE: ...).
    """
    m = FNAME_RE.match(path.name)
    if m:
        num = int(m.group("num"))
        header = f"{num}.({m.group('query')})"
        return num, header
    return None, f"FILE: {path.name}"


def index_from_record(r: dict, fallback: str) -> str:
    """
    Determine Index (source/provider) using:
    1) UR domain → PROVIDER_MAP
    2) DOI prefix heuristics (useful when UR is a doi.org link)
    3) cleaned folder name fallback
    """
    # 1) Try UR
    ur = r.get("UR")
    if isinstance(ur, list):
        ur = ur[0] if ur else ""
    ur = (ur or "").strip()

    if ur:
        try:
            host = urlparse(ur).netloc.lower()
            for k, v in PROVIDER_MAP.items():
                if k in host:
                    return v
        except Exception:
            pass

    # 2) Try DOI prefix heuristics
    doi = r.get("DO") or r.get("DOI") or ""
    doi = str(doi).strip().lower()
    if doi:
        if doi.startswith("10.1109"):
            return "IEEE Xplore"
        if doi.startswith("10.1016"):
            return "ScienceDirect"
        if doi.startswith("10.1007"):
            return "SpringerLink"
        if doi.startswith("10.1002"):
            return "Wiley Online Library"
        if doi.startswith("10.1145"):
            return "ACM Digital Library"

    # 3) Clean fallback (avoid ugly folder names like ieee_exports)
    fb = fallback.replace("_", " ").strip()
    fb = re.sub(r"\bexports\b", "", fb, flags=re.IGNORECASE).strip()
    return fb.title() if fb else "Unknown Source"


def ris_records_to_rows(records: list[dict], source_name: str, source_file: str) -> list[dict]:
    """Convert parsed RIS dictionaries into rows matching OUT_COLS + debug fields."""
    rows: list[dict] = []

    title_tags = ["T1", "TI", "CT", "BT"]
    doi_tags = ["DO", "DOI"]
    author_tags = ["AU", "A1", "AF", "A2", "A3", "A4", "ED"]

    for r in records:
        title = get_first_from_tags(r, title_tags)
        year = year_from_any(r)
        doi = clean_doi(get_first_from_tags(r, doi_tags))
        first_author = get_first_from_tags(r, author_tags)
        first_surname = surname_from_author(first_author)

        rows.append({
            "Title": title,
            "Year": year,
            "Index": index_from_record(r, fallback=source_name),
            "DOI": doi,
            "Author Surname": first_surname,
            "Author Name": first_author,
            "_source_file": source_file,  # debug: which RIS file this row came from
        })

    return rows


def parse_year_range(s: str) -> tuple[int | None, int | None]:
    """Parse year filter string like '2007-2017'. Returns (min_year, max_year)."""
    s = (s or "").strip()
    if not s:
        return None, None
    m = re.match(r"^\s*(\d{4})\s*-\s*(\d{4})\s*$", s)
    if not m:
        raise ValueError("Year filter must look like 2007-2017 (or empty).")
    y1, y2 = int(m.group(1)), int(m.group(2))
    if y1 > y2:
        y1, y2 = y2, y1
    return y1, y2


def matches_only_filter(row: pd.Series, terms: list[str]) -> bool:
    """Return True if all 'terms' appear (case-insensitive) in the row's key fields."""
    if not terms:
        return True
    hay = " ".join([
        str(row.get("Title", "") or ""),
        str(row.get("DOI", "") or ""),
        str(row.get("Author Name", "") or ""),
        str(row.get("Index", "") or ""),
    ]).lower()
    return all(t.lower() in hay for t in terms)


def apply_filters(df: pd.DataFrame, min_year: int | None, max_year: int | None, only_terms: list[str]) -> pd.DataFrame:
    """Apply year range filter and ONLY keyword filter to a DataFrame."""
    out = df.copy()

    # Year filter
    if min_year is not None or max_year is not None:
        def year_ok(y):
            y = str(y or "").strip()
            if not y.isdigit():
                return False
            yi = int(y)
            if min_year is not None and yi < min_year:
                return False
            if max_year is not None and yi > max_year:
                return False
            return True

        out = out[out["Year"].apply(year_ok)]

    # ONLY filter (AND logic across terms)
    if only_terms:
        out = out[out.apply(lambda r: matches_only_filter(r, only_terms), axis=1)]

    return out


def deduplicate(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicates:
    - prefer DOI-based dedup when DOI exists
    - fallback to Title+Year+Author Name when DOI is missing
    """
    if df.empty:
        return df

    doi_nonempty = df["DOI"].astype(str).str.strip().ne("")
    df_with_doi = df[doi_nonempty].drop_duplicates(subset=["DOI"], keep="first")
    df_no_doi = df[~doi_nonempty].drop_duplicates(subset=["Title", "Year", "Author Name"], keep="first")

    return pd.concat([df_with_doi, df_no_doi], ignore_index=True)


def build_grouped_excel_for_folder(
        ris_dir: Path,
        out_dir: Path,
        index_name: str | None,
        min_year: int | None,
        max_year: int | None,
        only_terms: list[str],
        do_dedup: bool
) -> Path | None:
    """
    Process one folder of RIS files and create a grouped Excel output.

    - Groups results by query header if filenames follow N.(query)_*.ris
      otherwise groups by file name.
    - Applies optional year/keyword filters per group.
    - Optionally de-duplicates (DOI first, then Title+Year+Author).
    - Writes two sheets: grouped_results + raw_results.

    Returns output path if created; otherwise None (if folder has no RIS files).
    """
    # Collect RIS files (case-insensitive extension)
    files = sorted(list(ris_dir.glob("*.ris")) + list(ris_dir.glob("*.RIS")))
    if not files:
        return None

    # Folder name is used as the "source name" fallback
    index_name = index_name or ris_dir.name

    # Output path (one Excel per folder)
    out_xlsx = out_dir / f"results_{ris_dir.name}_grouped.xlsx"

    # --- Build file groups ---
    groups: dict[str, dict] = {}
    order_keys: list[tuple[int | None, str]] = []

    for f in files:
        num, header = extract_group_header(f)  # header = "N.(query)" or "FILE: name"
        if header not in groups:
            groups[header] = {"num": num, "files": []}
            order_keys.append((num, header))
        groups[header]["files"].append(f)

    # Sort: numeric headers first, then FILE:... groups
    order_keys.sort(key=lambda x: (x[0] is None, x[0] if x[0] is not None else 10**9, x[1]))

    # --- Stats for CLI summary ---
    total_parsed = 0
    total_kept_after_filters = 0
    total_dedup_removed = 0

    # --- Output accumulators ---
    output_rows: list[dict] = []
    raw_rows: list[pd.DataFrame] = []

    for _, header in order_keys:
        # Header row for this group in grouped_results
        output_rows.append({c: "" for c in OUT_COLS})
        output_rows[-1]["Title"] = header

        # Parse all RIS files in this group
        frames: list[pd.DataFrame] = []
        for f in groups[header]["files"]:
            records = parse_ris_file(f)
            rows = ris_records_to_rows(records, source_name=index_name, source_file=f.name)
            df = pd.DataFrame(rows)
            df["_group"] = header
            frames.append(df)

        df_group = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=OUT_COLS)

        # Count how many were parsed before filters
        total_parsed += len(df_group)

        # Filters
        df_group = apply_filters(df_group, min_year, max_year, only_terms)
        total_kept_after_filters += len(df_group)

        # Dedup (count removals)
        if do_dedup and not df_group.empty:
            before = len(df_group)
            df_group = deduplicate(df_group)
            total_dedup_removed += (before - len(df_group))

        # Append rows to grouped_results sheet
        for _, r in df_group.iterrows():
            output_rows.append({c: r.get(c, "") for c in OUT_COLS})

        # Blank line between groups
        output_rows.append({c: "" for c in OUT_COLS})

        # raw_results is the concatenation of final processed groups
        raw_rows.append(df_group)

    # Build final DataFrames
    df_out = pd.DataFrame(output_rows, columns=OUT_COLS)
    df_raw = pd.concat(raw_rows, ignore_index=True) if raw_rows else pd.DataFrame()

    # Ensure output directory exists
    out_dir.mkdir(parents=True, exist_ok=True)

    # Write Excel
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="grouped_results")
        df_raw.to_excel(w, index=False, sheet_name="raw_results")

    # One-line summary (nice CLI UX)
    print(
        f"[{ris_dir.name}] files={len(files)} parsed={total_parsed} "
        f"kept={total_kept_after_filters} dedup_removed={total_dedup_removed} "
        f"output={out_xlsx}"
    )

    return out_xlsx


def prompt_user():
    """
    Interactive CLI prompt with validation.
    Re-asks until valid input is provided.
    """
    print("\n=== RIS → Excel ===")

    # --- Year filter ---
    while True:
        yr = input("Year filter (example 2007-2017, leave empty for no filter): ").strip()
        if not yr:
            min_year, max_year = None, None
            break
        try:
            min_year, max_year = parse_year_range(yr)
            break
        except ValueError as e:
            print(f"{e}")

    # --- ONLY filter ---
    only_raw = input(
        "ONLY filter keywords (comma-separated; all must appear; leave empty for no filter): "
    ).strip()
    only_terms = [t.strip() for t in only_raw.split(",") if t.strip()] if only_raw else []

    # --- Deduplication ---
    while True:
        dedup_raw = input("Write only ONE from duplicate articles? (y/n): ").strip().lower()
        if dedup_raw in {"y", "yes"}:
            do_dedup = True
            break
        if dedup_raw in {"n", "no"}:
            do_dedup = False
            break
        print("Please enter 'y' or 'n'.")

    # --- Source folders ---
    while True:
        src_raw = input("Source folders: type 'all' or comma-separated folder names: ").strip()
        if not src_raw:
            print("Please type 'all' or provide at least one folder name.")
            continue
        # accept anything here; actual existence is handled later
        break

    return min_year, max_year, only_terms, do_dedup, src_raw


def find_ris_folders(root: Path) -> list[Path]:
    """
    Return folders under `root` that contain at least one .ris/.RIS file.
    Prevents scanning unrelated dirs like .venv, outputs, etc.
    """
    folders = []
    for p in root.iterdir():
        if not p.is_dir():
            continue
        if list(p.glob("*.ris")) or list(p.glob("*.RIS")):
            folders.append(p)
    return sorted(folders)


def main():
    """Entry point: gather options, process selected folders, print output paths."""
    min_year, max_year, only_terms, do_dedup, src_raw = prompt_user()
    out_dir = Path("outputs")

    # Choose folders: either auto-detect RIS folders or use the user list
    if src_raw.lower() == "all" or not src_raw:
        source_dirs = find_ris_folders(Path("."))
        print(f"\nDetected {len(source_dirs)} RIS folders: {[d.name for d in source_dirs]}")
    else:
        source_dirs = [Path(s.strip()) for s in src_raw.split(",") if s.strip()]

    created = []
    for d in source_dirs:
        out = build_grouped_excel_for_folder(
            ris_dir=d,
            out_dir=out_dir,
            index_name=None,
            min_year=min_year,
            max_year=max_year,
            only_terms=only_terms,
            do_dedup=do_dedup,
        )
        if out:
            created.append(out)

    print("\nCreated outputs:")
    if created:
        for p in created:
            print(f"  - {p}")
    else:
        print("  (No RIS files found in the selected folders.)")


if __name__ == "__main__":
    main()
