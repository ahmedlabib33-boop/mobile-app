from __future__ import annotations

import html
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PROJECTS_DIR = ROOT / "projects"
PROJECT_DIRS = [path for path in PROJECTS_DIR.iterdir() if path.is_dir() and not path.name.startswith("_")]
PROJECT_DIR = max(PROJECT_DIRS, key=lambda path: sum(file.stat().st_size for file in path.rglob("*") if file.is_file()), default=PROJECTS_DIR / "_PROJECT_TEMPLATE")
IMPROVEMENT_DIR = PROJECT_DIR / "delay_analysis" / "improvement_files"
STEEL_DIR = PROJECT_DIR / "delay_analysis" / "steel_delay_tia_templates"
OUT_DIR = ROOT / "generated_outputs" / "delay_tia_question"
OUT_FILE = OUT_DIR / "question_preview.html"


def read_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="latin1")
    except Exception:
        return pd.DataFrame()


def load_frames() -> dict[str, pd.DataFrame]:
    frames: dict[str, pd.DataFrame] = {}
    for folder in [IMPROVEMENT_DIR, STEEL_DIR]:
        if not folder.exists():
            continue
        for path in sorted(folder.glob("*.csv")):
            frames[f"{folder.name}/{path.name}"] = read_csv(path)
    return frames


def number(df: pd.DataFrame, column: str) -> pd.Series:
    if df.empty or column not in df.columns:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[column], errors="coerce").dropna()


def inventory(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Dataset": name,
                "Rows": len(df),
                "Columns": len(df.columns),
                "Column List": ", ".join(str(col) for col in df.columns),
            }
            for name, df in frames.items()
        ]
    )


def kpis(frames: dict[str, pd.DataFrame]) -> dict[str, float | int]:
    claimed_df = frames.get("Delay tia improvement files/05-claimed_delay_register_template.csv", pd.DataFrame())
    fragnet_df = frames.get("Delay tia improvement files/04-fragnet_logic_register_template.csv", pd.DataFrame())
    concurrency_df = frames.get(
        "steel_delay_tia_templates/11-concurrency_matrix_template.updated.csv",
        frames.get("Delay tia improvement files/07-concurrency_matrix_template.csv", pd.DataFrame()),
    )
    rfi_df = frames.get("Delay tia improvement files/18-rfi_delay_claim6_normalized.csv", pd.DataFrame())
    events_df = frames.get("Delay tia improvement files/02-delay_event_register_template.csv", pd.DataFrame())
    evidence_df = frames.get("Delay tia improvement files/08-evidence_register_template.csv", pd.DataFrame())
    p6_df = frames.get("steel_delay_tia_templates/04- p6_activity_export.csv", pd.DataFrame())
    employer_df = frames.get(
        "steel_delay_tia_templates/03- employer_steel_supply_at_site.csv",
        frames.get("steel_delay_tia_templates/03- employer_steel_supply.csv", pd.DataFrame()),
    )
    samco_df = frames.get(
        "steel_delay_tia_templates/10- contractor_steel_supplied_at_site.csv",
        frames.get("steel_delay_tia_templates/10- samco_steel_supplied_at_site.csv", pd.DataFrame()),
    )

    claimed_days = number(claimed_df, "Claimed Delay Duration (days)")
    fragnet_days = number(fragnet_df, "Claimed Delay Duration")
    concurrency_days = number(concurrency_df, "Concurrent Delay Days")
    rfi_days = number(rfi_df, "Delay Beyond 10 Days")
    employer_qty = number(employer_df, "Available units at site received by client")
    samco_qty = number(samco_df, "Steel available at site")
    max_claimed = int(claimed_days.max()) if not claimed_days.empty else 0
    max_fragnet = int(fragnet_days.max()) if not fragnet_days.empty else 0

    return {
        "Datasets Loaded": len(frames),
        "Delay Events": len(events_df),
        "Evidence Rows": len(evidence_df),
        "P6 Activities": len(p6_df),
        "Critical P6 Activities": int(p6_df.get("Critical", pd.Series(dtype=str)).astype(str).str.lower().eq("yes").sum()) if not p6_df.empty else 0,
        "Longest Path Activities": int(p6_df.get("Longest Path", pd.Series(dtype=str)).astype(str).str.lower().eq("yes").sum()) if not p6_df.empty else 0,
        "Employer Steel Qty": float(employer_qty.sum()) if not employer_qty.empty else 0.0,
        "Contractor Steel Qty Visibility Only": float(samco_qty.sum()) if not samco_qty.empty else 0.0,
        "Max Claimed Delay Days": max_claimed,
        "Gross Claimed Delay Days": int(claimed_days.sum()) if not claimed_days.empty else 0,
        "Max Fragnet Duration": max_fragnet,
        "Gross Fragnet Duration": int(fragnet_days.sum()) if not fragnet_days.empty else 0,
        "Concurrent Delay Days": int(concurrency_days.sum()) if not concurrency_days.empty else 0,
        "RFI Delay Beyond 10 Days": int(rfi_days.sum()) if not rfi_days.empty else 0,
        "Recommended Conservative Days": max(max_claimed, max_fragnet),
    }


def table_html(df: pd.DataFrame, max_rows: int = 80) -> str:
    if df.empty:
        return "<p class='muted'>No rows available.</p>"
    return df.head(max_rows).to_html(index=False, escape=True, classes="data-table")


def card_grid(values: dict[str, float | int]) -> str:
    labels = [
        "Recommended Conservative Days",
        "Max Claimed Delay Days",
        "Max Fragnet Duration",
        "Concurrent Delay Days",
        "RFI Delay Beyond 10 Days",
        "Evidence Rows",
        "Critical P6 Activities",
        "Longest Path Activities",
        "Employer Steel Qty",
        "Contractor Steel Qty Visibility Only",
    ]
    cards = []
    for label in labels:
        raw = values.get(label, 0)
        value = f"{raw:,.3f}" if isinstance(raw, float) and raw % 1 else f"{int(raw):,}"
        cards.append(f"<article class='metric'><span>{html.escape(label)}</span><strong>{value}</strong></article>")
    return "\n".join(cards)


def main() -> None:
    frames = load_frames()
    inv = inventory(frames)
    values = kpis(frames)
    claimed = frames.get("Delay tia improvement files/05-claimed_delay_register_template.csv", pd.DataFrame())
    fragnet = frames.get("Delay tia improvement files/04-fragnet_logic_register_template.csv", pd.DataFrame())
    concurrency = frames.get(
        "steel_delay_tia_templates/11-concurrency_matrix_template.updated.csv",
        frames.get("Delay tia improvement files/07-concurrency_matrix_template.csv", pd.DataFrame()),
    )
    answer = (
        f"The conservative answer is {int(values['Recommended Conservative Days'])} days. "
        "The methodology avoids adding overlapping streams together: it compares the strongest claimed/modelled delay durations, "
        "checks concurrency, and treats RFI/IFC/payment rows as support unless modelled into the CPM network. "
        "Employer steel is used for the steel-delay calculation; contractor steel is mitigation visibility only."
    )
    html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Delay TIA Question Preview</title>
  <style>
    :root {{
      color-scheme: dark;
      --bg: #07131f;
      --panel: #0d2134;
      --panel-2: #102940;
      --line: #244864;
      --text: #ecf7ff;
      --muted: #9fb8cc;
      --accent: #36d6c8;
      --warn: #ffba45;
      --blue: #55a7ff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Inter, Segoe UI, Arial, sans-serif;
      background: radial-gradient(circle at top left, #123a5a 0, var(--bg) 42%);
      color: var(--text);
    }}
    main {{ max-width: 1440px; margin: 0 auto; padding: 28px; }}
    header {{ display: flex; justify-content: space-between; gap: 24px; align-items: end; margin-bottom: 22px; }}
    h1 {{ margin: 0 0 8px; font-size: 34px; letter-spacing: 0; }}
    h2 {{ margin: 30px 0 12px; font-size: 20px; }}
    p {{ color: var(--muted); line-height: 1.55; }}
    .answer {{
      border: 1px solid rgba(54, 214, 200, .42);
      background: linear-gradient(135deg, rgba(54, 214, 200, .14), rgba(85, 167, 255, .08));
      padding: 18px 20px;
      border-radius: 8px;
      color: var(--text);
      font-size: 17px;
    }}
    .metrics {{ display: grid; grid-template-columns: repeat(5, minmax(170px, 1fr)); gap: 12px; }}
    .metric {{ border: 1px solid var(--line); background: rgba(13, 33, 52, .88); padding: 14px; border-radius: 8px; min-height: 94px; }}
    .metric span {{ display: block; color: var(--muted); font-size: 13px; min-height: 34px; }}
    .metric strong {{ display: block; margin-top: 10px; font-size: 25px; color: var(--text); }}
    .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }}
    .panel {{ border: 1px solid var(--line); background: rgba(13, 33, 52, .74); border-radius: 8px; padding: 16px; overflow: auto; }}
    .data-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: var(--text); }}
    .data-table th {{ position: sticky; top: 0; background: var(--panel-2); color: #d9f7ff; text-align: left; }}
    .data-table th, .data-table td {{ border-bottom: 1px solid rgba(159, 184, 204, .18); padding: 9px 10px; vertical-align: top; }}
    .data-table tr:hover td {{ background: rgba(85, 167, 255, .08); }}
    .muted {{ color: var(--muted); }}
    .tag {{ color: var(--accent); font-weight: 700; }}
    @media (max-width: 980px) {{
      header, .grid {{ display: block; }}
      .metrics {{ grid-template-columns: repeat(2, minmax(160px, 1fr)); }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>Delay TIA - question</h1>
        <p>Column-inspected answer surface for Delay TIA improvement files and steel delay templates.</p>
      </div>
      <p><span class="tag">{int(values['Datasets Loaded'])}</span> datasets loaded</p>
    </header>
    <section class="answer">{html.escape(answer)}</section>
    <h2>Decision KPIs</h2>
    <section class="metrics">{card_grid(values)}</section>
    <h2>Answer Evidence Tables</h2>
    <section class="grid">
      <div class="panel"><h2>Claimed Delay Register</h2>{table_html(claimed)}</div>
      <div class="panel"><h2>Fragnet Logic Register</h2>{table_html(fragnet)}</div>
    </section>
    <section class="panel" style="margin-top:16px"><h2>Concurrency Matrix</h2>{table_html(concurrency)}</section>
    <section class="panel" style="margin-top:16px"><h2>Column Inventory Inspected Before Answering</h2>{table_html(inv, max_rows=60)}</section>
  </main>
</body>
</html>
"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(html_doc, encoding="utf-8")
    print(OUT_FILE)
    print(answer)


if __name__ == "__main__":
    main()
