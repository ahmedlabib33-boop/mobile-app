from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from construction_system.database import DEFAULT_DB_PATH, init_db
from construction_system.seed import seed_demo_data


def main():
    init_db(DEFAULT_DB_PATH, reset=True)
    seed_demo_data(DEFAULT_DB_PATH)
    print(f"Database created and seeded at: {DEFAULT_DB_PATH}")


if __name__ == "__main__":
    main()