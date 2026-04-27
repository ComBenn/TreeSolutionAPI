from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

import pandas as pd

from config import (
    DEFAULT_KEYWORDS_FILE,
    DEFAULT_OUTPUT_FILE,
    DEFAULT_USERS_FILE,
    DEFAULT_USERS_SHEET,
)
from io_utils import load_table


def _bundle_base_dir() -> Path:
    """Liefert das Verzeichnis der gebuendelten Anwendungsdateien."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _app_runtime_dir() -> Path:
    """Bestimmt das beschreibbare Laufzeitverzeichnis fuer Statusdateien."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class AppState:
    """Haelt Benutzerpfade, geladene DataFrames und persistente Runtime-Dateien."""
    def __init__(self) -> None:
        self.bundle_dir = _bundle_base_dir()
        self.runtime_dir = _app_runtime_dir()
        self.runtime_dir.mkdir(parents=True, exist_ok=True)

        self.users_file = str(self.runtime_dir / DEFAULT_USERS_FILE)
        self.users_sheet = DEFAULT_USERS_SHEET
        self.keywords_file = str(self.runtime_dir / DEFAULT_KEYWORDS_FILE)
        self.output_file = str(self.runtime_dir / DEFAULT_OUTPUT_FILE)
        self.original_df: pd.DataFrame | None = None
        self.current_df: pd.DataFrame | None = None
        self.batch_export_tracker_file = self.runtime_dir / "batch_export_tracker.json"
        self.ui_state_file = self.runtime_dir / "ui_state.json"
        self.batch_exported_ids: set[str] = set()
        self.runtime_warnings: list[str] = []
        self.batch_export_tracker_error: str | None = None
        self._seed_runtime_file("README.md")
        self._seed_runtime_file("keywords_technische_accounts.txt")
        self._seed_runtime_file("batch_export_tracker.json", default_text='{"exported_ids": []}\n')
        self._seed_runtime_file("ui_state.json", default_text="{}\n")
        self._load_batch_export_tracker()

    def _add_runtime_warning(self, message: str) -> None:
        """Puffert nicht-fatale Laufzeitwarnungen fuer die spaetere UI-Ausgabe."""
        if message:
            self.runtime_warnings.append(message)

    def consume_runtime_warnings(self) -> list[str]:
        """Liefert gepufferte Warnungen und leert den internen Speicher."""
        warnings = list(self.runtime_warnings)
        self.runtime_warnings.clear()
        return warnings

    def _seed_runtime_file(self, filename: str, default_text: str | None = None) -> None:
        """Erzeugt oder kopiert benoetigte Runtime-Dateien beim Programmstart."""
        dst = self.runtime_dir / filename
        if dst.exists():
            return
        src = self.bundle_dir / filename
        try:
            if src.exists():
                shutil.copy2(src, dst)
            elif default_text is not None:
                dst.write_text(default_text, encoding="utf-8")
        except Exception as exc:
            self._add_runtime_warning(
                f"Runtime-Datei konnte nicht vorbereitet werden: {filename} ({exc})"
            )

    def load_users(self) -> None:
        """Laedt die Benutzerquelle und setzt Original- sowie Arbeitskopie."""
        self.original_df = load_table(self.users_file, self.users_sheet)
        self.current_df = self.original_df.copy()

    def reset(self) -> None:
        """Setzt die aktuelle Auswahl auf die zuletzt geladene Originaldatei zurueck."""
        if self.original_df is None:
            raise RuntimeError("Noch keine Benutzerdatei geladen.")
        self.current_df = self.original_df.copy()

    def _load_batch_export_tracker(self) -> None:
        """Laedt die persistente Batch-Merkliste oder markiert sie als defekt."""
        p = self.batch_export_tracker_file
        self.batch_export_tracker_error = None
        if not p.exists():
            self.batch_exported_ids = set()
            return
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
            ids = payload.get("exported_ids", [])
            if not isinstance(ids, list):
                raise ValueError("Feld 'exported_ids' muss eine Liste sein.")
            self.batch_exported_ids = {str(x) for x in ids if str(x).strip()}
        except Exception as exc:
            self.batch_exported_ids = set()
            backup_path = self._backup_invalid_tracker_file(p)
            backup_hint = f" Sicherung: {backup_path.name}." if backup_path is not None else ""
            self.batch_export_tracker_error = (
                "Batch-Merkliste ist beschädigt und wurde nicht übernommen."
                f"{backup_hint} Bitte Merkliste prüfen oder zurücksetzen. ({exc})"
            )
            self._add_runtime_warning(self.batch_export_tracker_error)

    def _backup_invalid_tracker_file(self, path: Path) -> Path | None:
        """Sichert eine beschaedigte Tracker-Datei unter einem freien Namen."""
        suffix = path.suffix or ""
        for counter in range(100):
            if counter == 0:
                candidate = path.with_name(f"{path.name}.invalid")
            else:
                candidate = path.with_name(f"{path.name}.invalid.{counter}")
            if candidate.exists():
                continue
            try:
                shutil.copy2(path, candidate)
                return candidate
            except Exception as exc:
                self._add_runtime_warning(
                    f"Beschädigte Batch-Merkliste konnte nicht gesichert werden: {path.name} ({exc})"
                )
                return None
        self._add_runtime_warning(
            f"Beschädigte Batch-Merkliste konnte nicht gesichert werden: {path.name} (kein freier Sicherungsname)"
        )
        return None

    def save_batch_export_tracker(self) -> None:
        """Schreibt die aktuelle Batch-Merkliste als JSON auf die Platte."""
        p = self.batch_export_tracker_file
        payload = {"exported_ids": sorted(self.batch_exported_ids)}
        p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.batch_export_tracker_error = None

    def reset_batch_export_tracker(self) -> None:
        """Leert die Batch-Merkliste und speichert den Reset sofort persistent."""
        self.batch_exported_ids.clear()
        self.batch_export_tracker_error = None
        self.save_batch_export_tracker()
