from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from reference_db import bootstrap_reference_data

APP_DB_FILE = BASE_DIR / "db" / "App_Url Data base.xlsx"
CITY_DB_FILE = BASE_DIR / "db" / "City for Aoutomation.xlsx"


if __name__ == "__main__":
    bootstrap_reference_data(APP_DB_FILE, CITY_DB_FILE)
    print("PostgreSQL reference data bootstrap complete.")
