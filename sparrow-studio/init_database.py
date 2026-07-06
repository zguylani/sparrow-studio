from pathlib import Path
import sqlite3

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "sparrow_studio.db"
SCHEMA_PATH = BASE_DIR / "database" / "schema.sql"


def init_database():
    DATA_DIR.mkdir(exist_ok=True)

    if not SCHEMA_PATH.exists():
        raise FileNotFoundError(f"Schema file not found: {SCHEMA_PATH}")

    schema = SCHEMA_PATH.read_text(encoding="utf-8")

    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.executescript(schema)

    print(f"Sparrow Studio database created at: {DB_PATH}")


if __name__ == "__main__":
    init_database()