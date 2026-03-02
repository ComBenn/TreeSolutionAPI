# menu.py

from pathlib import Path
import pandas as pd

from config import (
    DEFAULT_USERS_FILE,
    DEFAULT_USERS_SHEET,
    DEFAULT_KEYWORDS_FILE,
    DEFAULT_OUTPUT_FILE,
)
from io_utils import load_table, load_keywords_txt, append_keywords_txt
from filters_technical import mark_technical_accounts
from filters_employee_list import mark_by_employee_list
from exporter import build_upload_export, export_utf8_csv


class AppState:
    def __init__(self):
        self.users_file = DEFAULT_USERS_FILE
        self.users_sheet = DEFAULT_USERS_SHEET
        self.keywords_file = DEFAULT_KEYWORDS_FILE
        self.output_file = DEFAULT_OUTPUT_FILE

        self.original_df: pd.DataFrame | None = None
        self.current_df: pd.DataFrame | None = None

    def load_users(self):
        self.original_df = load_table(self.users_file, self.users_sheet)
        self.current_df = self.original_df.copy()

    def reset(self):
        if self.original_df is None:
            raise RuntimeError("Noch keine Benutzerdatei geladen.")
        self.current_df = self.original_df.copy()


def _print_head(df: pd.DataFrame, n: int = 10):
    print(f"Zeilen: {len(df)}")
    if len(df) == 0:
        print("(leer)")
        return
    print(df.head(n).to_string(index=False))


def _choose_file(prompt: str, default: str | None = None) -> str:
    text = f"{prompt}"
    if default:
        text += f" [{default}]"
    text += ": "
    value = input(text).strip()
    return value if value else (default or "")


def _choose_sheet(prompt: str, default: str | None = None) -> str | None:
    text = f"{prompt}"
    if default:
        text += f" [{default}]"
    text += ": "
    value = input(text).strip()
    return value if value else default


def run_menu():
    state = AppState()

    while True:
        print("\n=== Menü ===")
        print(f"Benutzerdatei: {state.users_file} | Sheet: {state.users_sheet}")
        print(f"Keyword-Datei: {state.keywords_file}")
        print(f"Output-Datei: {state.output_file}")
        if state.current_df is not None:
            print(f"Aktuelle Zeilen im Arbeitssatz: {len(state.current_df)}")
        else:
            print("Aktuelle Zeilen im Arbeitssatz: (nicht geladen)")

        print("\n1) Benutzerdatei laden")
        print("2) Arbeitssatz zurücksetzen (Original wiederherstellen)")
        print("3) Technische Accounts markieren")
        print("4) Nur technische Accounts auswählen")
        print("5) Technische Accounts ausschliessen")
        print("6) Mitarbeiterliste markieren (Datei frei wählen)")
        print("7) Nur markierte Mitarbeiterliste auswählen")
        print("8) Markierte Mitarbeiterliste ausschliessen")
        print("9) Vorschau aktueller Arbeitssatz")
        print("10) Keywords anzeigen")
        print("11) Keywords ergänzen")
        print("12) Upload.csv exportieren (UTF-8)")
        print("13) Einstellungen ändern (Dateien/Sheets)")
        print("0) Beenden")

        choice = input("Auswahl: ").strip()

        try:
            if choice == "1":
                state.load_users()
                print(f"Benutzer geladen: {len(state.current_df)}")

            elif choice == "2":
                state.reset()
                print("Arbeitssatz auf Original zurückgesetzt.")

            elif choice == "3":
                if state.current_df is None:
                    print("Zuerst Benutzerdatei laden.")
                    continue
                keywords = load_keywords_txt(state.keywords_file)
                print(f"Keywords geladen: {len(keywords)}")
                state.current_df = mark_technical_accounts(state.current_df, keywords)
                print("Markierung durchgeführt: flag_technical_account / flag_technical_reason")
                print(f"Treffer: {int(state.current_df['flag_technical_account'].sum())}")

            elif choice == "4":
                if state.current_df is None or "flag_technical_account" not in state.current_df.columns:
                    print("Zuerst technische Accounts markieren.")
                    continue
                state.current_df = state.current_df[state.current_df["flag_technical_account"] == True].copy()
                print(f"Nur technische Accounts ausgewählt. Verbleibend: {len(state.current_df)}")

            elif choice == "5":
                if state.current_df is None or "flag_technical_account" not in state.current_df.columns:
                    print("Zuerst technische Accounts markieren.")
                    continue
                state.current_df = state.current_df[state.current_df["flag_technical_account"] != True].copy()
                print(f"Technische Accounts ausgeschlossen. Verbleibend: {len(state.current_df)}")

            elif choice == "6":
                if state.current_df is None:
                    print("Zuerst Benutzerdatei laden.")
                    continue

                list_file = _choose_file("Pfad Mitarbeiterliste", "")
                if not list_file:
                    print("Keine Datei angegeben.")
                    continue

                list_sheet = None
                if Path(list_file).suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                    list_sheet = _choose_sheet("Sheet der Mitarbeiterliste", None)

                df_list = load_table(list_file, list_sheet)
                flag_name = input("Flag-Name für diese Liste [flag_employee_list]: ").strip() or "flag_employee_list"

                state.current_df = mark_by_employee_list(state.current_df, df_list, flag_name=flag_name)
                print(f"Markierung durchgeführt: {flag_name} / {flag_name}_reason")
                print(f"Treffer: {int(state.current_df[flag_name].sum())}")

            elif choice == "7":
                if state.current_df is None:
                    print("Zuerst Daten laden und eine Mitarbeiterliste markieren.")
                    continue
                flag_name = input("Welcher Flag-Name soll gefiltert werden? [flag_employee_list]: ").strip() or "flag_employee_list"
                if flag_name not in state.current_df.columns:
                    print(f"Flag nicht gefunden: {flag_name}")
                    continue
                state.current_df = state.current_df[state.current_df[flag_name] == True].copy()
                print(f"Nur markierte Zeilen ausgewählt. Verbleibend: {len(state.current_df)}")

            elif choice == "8":
                if state.current_df is None:
                    print("Zuerst Daten laden und eine Mitarbeiterliste markieren.")
                    continue
                flag_name = input("Welcher Flag-Name soll ausgeschlossen werden? [flag_employee_list]: ").strip() or "flag_employee_list"
                if flag_name not in state.current_df.columns:
                    print(f"Flag nicht gefunden: {flag_name}")
                    continue
                state.current_df = state.current_df[state.current_df[flag_name] != True].copy()
                print(f"Markierte Zeilen ausgeschlossen. Verbleibend: {len(state.current_df)}")

            elif choice == "9":
                if state.current_df is None:
                    print("Noch keine Benutzerdatei geladen.")
                    continue
                _print_head(state.current_df, n=10)

            elif choice == "10":
                kws = sorted(load_keywords_txt(state.keywords_file))
                print(f"Keywords ({len(kws)}):")
                for k in kws:
                    print(k)

            elif choice == "11":
                raw = input("Neue Keywords (kommagetrennt): ").strip()
                if not raw:
                    print("Keine Eingabe.")
                    continue
                n_added = append_keywords_txt(state.keywords_file, [x.strip() for x in raw.split(",")])
                print(f"Keywords ergänzt: {n_added}")

            elif choice == "12":
                if state.current_df is None:
                    print("Zuerst Benutzerdatei laden.")
                    continue
                export_df = build_upload_export(state.current_df)
                export_utf8_csv(export_df, state.output_file)
                print(f"Export geschrieben: {state.output_file} (UTF-8 CSV)")
                print(f"Zeilen exportiert: {len(export_df)}")

            elif choice == "13":
                state.users_file = _choose_file("Benutzerdatei", state.users_file)
                if Path(state.users_file).suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                    state.users_sheet = _choose_sheet("Sheet in Benutzerdatei", state.users_sheet)
                else:
                    state.users_sheet = None

                state.keywords_file = _choose_file("Keyword-Datei (.txt)", state.keywords_file)
                state.output_file = _choose_file("Output-Datei (.csv)", state.output_file)

                print("Einstellungen aktualisiert.")

            elif choice == "0":
                print("Beendet.")
                break

            else:
                print("Ungültige Auswahl.")

        except Exception as e:
            print(f"Fehler: {e}")
