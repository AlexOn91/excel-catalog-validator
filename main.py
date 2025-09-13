# main.py
import sys
import argparse
from offline_app import main as launch_gui
from downloadfailreport import export_data_format_fails

def cli_report(args):
    # aici poți prelua argumente din linia de comandă, dacă ai nevoie
    export_data_format_fails()

def main():
    parser = argparse.ArgumentParser(
        description="OfflineCatalogValidator: GUI & raportare"
    )
    parser.add_argument(
        "--report",
        action="store_true",
        help="Rulează doar raportarea de eşecuri (CLI)."
    )
    args = parser.parse_args()

    if args.report:
        cli_report(args)
    else:
        launch_gui()

if __name__ == "__main__":
    main()
