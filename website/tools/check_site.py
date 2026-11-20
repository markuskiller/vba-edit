"""Consistency checks for the static site: internal links, anchors, recordings, images, inline styles.

python check_site.py        (exit code 1 when something is wrong)
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PUBLIC = Path(__file__).resolve().parent.parent / "public"
problems: list[str] = []

pages = {page.name: page.read_text(encoding="utf-8") for page in PUBLIC.glob("*.html")}
ids = {name: set(re.findall(r'\bid="([^"]+)"', text)) for name, text in pages.items()}
casts_text = (PUBLIC / "assets" / "casts.js").read_text(encoding="utf-8")
casts = json.loads(casts_text[casts_text.index("{") : casts_text.rindex("}") + 1])

for name, text in pages.items():
    for duplicate in {i for i in re.findall(r'\bid="([^"]+)"', text) if text.count(f'id="{i}"') > 1}:
        problems.append(f"{name}: id used twice: {duplicate}")
    for href in re.findall(r'href="([^"]+)"', text):
        if href.startswith(("http://", "https://", "mailto:")):
            continue
        target, _, anchor = href.partition("#")
        if target.startswith("/"):
            # 404.html uses root-relative links, in the extensionless form Cloudflare Pages serves.
            target = target.lstrip("/") or "index.html"
            if not (PUBLIC / target).exists() and (PUBLIC / (target + ".html")).exists():
                target += ".html"
        target = target or name
        if target not in pages and not (PUBLIC / target).exists():
            problems.append(f"{name}: link to missing file: {href}")
        elif anchor and target in ids and anchor not in ids[target]:
            problems.append(f"{name}: link to missing anchor: {href}")
    for source in re.findall(r'(?:src|href)="(assets/[^"]+|favicon[^"]+)"', text):
        if not (PUBLIC / source).exists():
            problems.append(f"{name}: missing asset: {source}")
    for cast, steps in re.findall(r'data-cast="([^"]+)"(?:\s+data-steps="([^"]+)")?', text):
        if cast not in casts:
            problems.append(f"{name}: unknown recording: {cast}")
        elif steps:
            bounds = [int(part) for part in steps.split("-")]
            if bounds[0] >= len(casts[cast]) or (len(bounds) > 1 and bounds[1] > len(casts[cast])):
                problems.append(f"{name}: recording {cast} has no steps {steps} (it has {len(casts[cast])})")
    if re.search(r'\sstyle="', text) or "<style" in text or re.search(r"<script(?![^>]*\bsrc=)", text):
        problems.append(f"{name}: inline style or script (blocked by the Content-Security-Policy in _headers)")
    if re.search(r'\son[a-z]+="', text):
        problems.append(f"{name}: inline event handler (blocked by the Content-Security-Policy)")

print(f"{len(pages)} pages, {sum(len(v) for v in ids.values())} anchors, {len(casts)} recordings checked")
for problem in problems:
    print("PROBLEM:", problem)
sys.exit(1 if problems else 0)
