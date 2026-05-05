"""
Local deployment configuration template.

Copy this file to `local_config.py` and edit the values to match your
organisation's lab names and shared folder layout.  The real local_config.py
is gitignored so your deployment-specific values never get committed.

If `local_config.py` is missing at runtime the app falls back to a single
generic "Lab 1" entry pointing at ~/Documents/Reports — usable for testing
but probably not what you want in production.
"""

from pathlib import Path

# Labs that appear in the dropdown
LABS = [
    "Lab 1",
    "Lab 2",
    # add more...
]

# Base folder under each user's home directory.  Each lab's destination is
# ONEDRIVE_BASE / <lab name> / "Plant Testing".
ONEDRIVE_BASE = (
    Path.home()
    / "OneDrive - Your Company, Inc"
    / "Your Department"
    / "Your Subfolder"
)
