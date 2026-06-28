from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from construction_system.database import DEFAULT_DB_PATH
from construction_system.reporting import generate_html_report


def main():
    output = generate_html_report(DEFAULT_DB_PATH)
    print(f"HTML report generated at: {output}")


if __name__ == "__main__":
    main()