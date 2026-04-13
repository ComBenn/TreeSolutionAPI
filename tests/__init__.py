from pathlib import Path
import sys


SRC_DIR = Path(__file__).resolve().parents[1] / "src" / "treesolution_helper" / "files"
src_dir_str = str(SRC_DIR)
if src_dir_str not in sys.path:
    sys.path.insert(0, src_dir_str)
