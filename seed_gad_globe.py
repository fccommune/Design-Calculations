"""
Bulk-load one of the GAD Automation Globe valve lookup tables from a CSV
file - use this to migrate data out of the old SQL Server tables
(GAGlobeTable, GAGlobeHookUp, GACrossSecGlobe, GADimvalveGlobe,
GADimActGlobe) into this app's Postgres tables, which start empty.

The CSV's header row must use this app's column names (snake_case, no
spaces) - see the --columns output below for each table, or app/models.py.
Export from SQL Server Management Studio with "Results To > Save Results
As..." (or a SELECT ... query) and rename the header row to match.

Usage:
    python seed_gad_globe.py <table> <csv_path> [--replace]

    <table> is one of: ga_globe_table, ga_globe_hookup, ga_crosssec_globe,
                        ga_dim_valve_globe, ga_dim_act_globe

    --replace   delete all existing rows in that table first (default is to
                append)

Examples:
    python seed_gad_globe.py ga_globe_table data/GAGlobeTable.csv
    python seed_gad_globe.py ga_dim_act_globe data/GADimActGlobe.csv --replace
"""
import csv
import sys

from dotenv import load_dotenv

load_dotenv()

from app import create_app
from app.extensions import db
from app.models import (
    GAGlobeTable,
    GAGlobeHookUp,
    GACrossSecGlobe,
    GADimValveGlobe,
    GADimActGlobe,
)

TABLES = {
    "ga_globe_table": GAGlobeTable,
    "ga_globe_hookup": GAGlobeHookUp,
    "ga_crosssec_globe": GACrossSecGlobe,
    "ga_dim_valve_globe": GADimValveGlobe,
    "ga_dim_act_globe": GADimActGlobe,
}


def _columns(model):
    return [c.name for c in model.__table__.columns if c.name != "id"]


def main():
    args = sys.argv[1:]
    if len(args) not in (2, 3) or args[0] not in TABLES:
        print(__doc__)
        print("Known tables and expected CSV columns:\n")
        for name, model in TABLES.items():
            print(f"  {name}: {', '.join(_columns(model))}")
        sys.exit(1)

    table_name, csv_path = args[0], args[1]
    replace = "--replace" in args
    model = TABLES[table_name]
    expected_columns = set(_columns(model))

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = set(reader.fieldnames or [])
        unknown = fieldnames - expected_columns
        if unknown:
            print(f"Warning: ignoring unknown CSV columns: {', '.join(sorted(unknown))}")

        rows = []
        for raw in reader:
            data = {k: (v if v not in (None, "") else None) for k, v in raw.items() if k in expected_columns}
            rows.append(data)

    if not rows:
        print("No data rows found in the CSV.")
        sys.exit(1)

    app = create_app("development")
    with app.app_context():
        if replace:
            deleted = model.query.delete()
            print(f"Deleted {deleted} existing row(s) from {table_name}.")

        db.session.bulk_insert_mappings(model, rows)
        db.session.commit()
        print(f"Inserted {len(rows)} row(s) into {table_name}.")


if __name__ == "__main__":
    main()
