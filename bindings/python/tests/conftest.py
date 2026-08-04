import sys
from pathlib import Path

# Make the compiled xlpp extension (built by the repo-root setup.py into the
# repository root) importable when running pytest. conftest lives at
# <root>/bindings/python/tests/, so parents[3] is the repository root.
REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))
