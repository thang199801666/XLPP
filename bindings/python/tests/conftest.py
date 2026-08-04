import sys
from pathlib import Path

# Make the compiled xlpp extension (in bindings/python) importable when
# running pytest from the tests/ directory.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
