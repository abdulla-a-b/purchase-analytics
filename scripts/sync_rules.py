#!/usr/bin/env python3
"""Copy the shared data-entry rules (and the calculation engine) from assets/js/entry-rules.js into apps-script/Code.gs
(between the <entry-rules> markers) so the website and the Google Sheet use identical rules."""
from pathlib import Path
import re, sys
ROOT = Path(__file__).resolve().parents[1]
src = (ROOT / "assets/js/entry-rules.js").read_text(encoding="utf-8")
block = src[src.index("// <entry-rules>"):src.index("// </entry-rules>") + len("// </entry-rules>")]
gs_path = ROOT / "apps-script/Code.gs"
gs = gs_path.read_text(encoding="utf-8")
new = re.sub(r"// <entry-rules>.*?// </entry-rules>", lambda m: block, gs, flags=re.S)
gs_path.write_text(new, encoding="utf-8")
print("✔ rules synced into apps-script/Code.gs" if new != gs else "✔ Code.gs already up to date")

# The email alerts run the dashboard's own calculation engine inside Apps Script.
eng = (ROOT / "assets/js/engine.js").read_text(encoding="utf-8")
out = ("/* Engine.gs — exact copy of assets/js/engine.js (written by scripts/sync_rules.py; do not edit here).\n"
       " * Lets the email alerts calculate exactly like the web dashboard. */\n" + eng)
eg = ROOT / "apps-script/Engine.gs"
if not eg.exists() or eg.read_text(encoding="utf-8") != out:
    eg.write_text(out, encoding="utf-8"); print("✔ engine copied into apps-script/Engine.gs")
else:
    print("✔ Engine.gs already up to date")
