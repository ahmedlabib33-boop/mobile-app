from pathlib import Path

import pandas as pd

from src.construction_system.letters_auto_ingest import (
    ACE_LINKS_SHEET,
    ACE_SHEET,
    AUTO_REGISTER_SHEET,
    SAMCO_LINKS_SHEET,
    SAMCO_SHEET,
    THREADS_SHEET,
    folder_fingerprint,
    merge_inbox_letters,
)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_auto_ingest_adds_direction_links_and_threads(tmp_path):
    samco_dir = tmp_path / "From SAMCO to ACEPM"
    ace_dir = tmp_path / "From ACEPM to SAMCO"
    samco_dir.mkdir(parents=True)
    ace_dir.mkdir(parents=True)

    (samco_dir / "083-steel-delay.txt").write_text(
        "Date: 18-Jun-2026\nSubject: Delayed free issue reinforcement steel\n"
        "This notice records steel supply delay and requests extension of time.\n"
        "Related response BD-ACEPM-SAMCO-LET-038.",
        encoding="utf-8",
    )
    (ace_dir / "038-response.txt").write_text(
        "Date: 19-Jun-2026\nSubject: Response to steel delay notice\n"
        "Reference BD-CW-SAMCO-ACE-LET-STR-083. Engineer response regarding reinforcement steel.",
        encoding="utf-8",
    )

    result = merge_inbox_letters({}, tmp_path, read_text)

    assert result[SAMCO_SHEET].iloc[0]["Ref No"] == "BD-CW-SAMCO-ACE-LET-STR-083"
    assert result[ACE_SHEET].iloc[0]["Ref No"] == "BD-ACEPM-SAMCO-LET-038"
    assert result[SAMCO_LINKS_SHEET].iloc[0]["Related ACE Ref No(s)"] == "BD-ACEPM-SAMCO-LET-038"
    assert result[ACE_LINKS_SHEET].iloc[0]["Related SAMCO Ref No(s)"] == "BD-CW-SAMCO-ACE-LET-STR-083"
    assert not result[THREADS_SHEET].empty
    assert result[AUTO_REGISTER_SHEET]["Status"].eq("Added Automatically").all()


def test_auto_ingest_does_not_duplicate_existing_reference(tmp_path):
    samco_dir = tmp_path / "From SAMCO to ACEPM"
    samco_dir.mkdir(parents=True)
    (samco_dir / "083-steel-delay.txt").write_text(
        "Date: 18-Jun-2026\nSubject: Steel supply delay",
        encoding="utf-8",
    )
    source = {SAMCO_SHEET: pd.DataFrame([{"Ref No": "BD-CW-SAMCO-ACE-LET-STR-083"}])}

    result = merge_inbox_letters(source, tmp_path, read_text)

    assert len(result[SAMCO_SHEET]) == 1
    assert result[AUTO_REGISTER_SHEET].iloc[0]["Status"] == "Already Registered"


def test_folder_fingerprint_changes_when_letter_changes(tmp_path):
    letter = tmp_path / "From SAMCO to ACEPM" / "084.txt"
    letter.parent.mkdir(parents=True)
    letter.write_text("first", encoding="utf-8")
    first = folder_fingerprint(tmp_path)
    letter.write_text("second version", encoding="utf-8")
    second = folder_fingerprint(tmp_path)

    assert first != second
