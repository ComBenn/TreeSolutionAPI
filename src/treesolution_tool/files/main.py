# main.py

try:
    from ui_app import run_ui
except Exception:
    run_ui = None

from menu import run_menu


if __name__ == "__main__":
    if run_ui is not None:
        run_ui()
    else:
        run_menu()
