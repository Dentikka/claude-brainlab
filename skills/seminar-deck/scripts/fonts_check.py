"""Fonts a deck uses vs fonts installed on this Windows machine (system and per-user).
A missing font renders as substitute glyphs in PowerPoint. Install it per user: copy the files
to %LOCALAPPDATA%/Microsoft/Windows/Fonts, add HKCU Fonts registry values, restart PowerPoint.
For a deck that travels, save it with embedded fonts.

    python fonts_check.py deck.pptx
"""
import re
import sys
import winreg
import zipfile

KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def deck_fonts(path):
    """Fonts named on slides, layouts, masters and notes, plus the theme's Latin major/minor
    fonts. The theme's per-script fallbacks (Thai, CJK, ...) are ignored: they are never
    used by Latin or Cyrillic text."""
    names = set()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not n.endswith(".xml"):
                continue
            if n.startswith(("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/notesSlides/")):
                names |= set(re.findall(rb'typeface="([^"+][^"]*)"', z.read(n)))
            elif n.startswith("ppt/theme/"):
                names |= set(re.findall(rb'<a:latin typeface="([^"+][^"]*)"', z.read(n)))
    return {x.decode("utf-8") for x in names}


def installed():
    out = set()
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(root, KEY) as k:
                for i in range(winreg.QueryInfoKey(k)[1]):
                    out.add(re.sub(r"\s*\(.*\)$", "", winreg.EnumValue(k, i)[0]).strip().lower())
        except OSError:
            pass
    return out


if __name__ == "__main__":
    have = installed()
    for f in sorted(deck_fonts(sys.argv[1])):
        ok = any(h == f.lower() or h.startswith(f.lower() + " ") for h in have)
        print(f"{'ok     ' if ok else 'MISSING'}  {f}")
