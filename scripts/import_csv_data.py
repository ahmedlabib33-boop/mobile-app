from pathlib import Path
import argparse
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from construction_system.database import DEFAULT_DB_PATH
from construction_system.importers import import_csv_folder


def main():
    project_dirs = [path for path in (ROOT / "projects").iterdir() if path.is_dir() and not path.name.startswith("_")]
    default_project = max(project_dirs, key=lambda path: sum(file.stat().st_size for file in path.rglob("*") if file.is_file()), default=ROOT / "projects" / "_PROJECT_TEMPLATE")
    parser = argparse.ArgumentParser(description="Import CSV files into the construction project database.")
    parser.add_argument(
        "folder",
        nargs="?",
        default=str(default_project / "data" / "import_templates"),
        help="Project-owned folder containing CSV files.",
    )
    parser.add_argument("--reset", action="store_true", help="Recreate the database before import")
    args = parser.parse_args()

    results = import_csv_folder(args.folder, DEFAULT_DB_PATH, reset=args.reset)

    if not results:
        print("No CSV files were imported.")
        return

    for table_name, count in results.items():
        print(f"{table_name}: {count} rows imported")

    print(f"Database updated at: {DEFAULT_DB_PATH}")


if __name__ == "__main__":
    main()
