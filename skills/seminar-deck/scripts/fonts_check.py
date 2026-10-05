"""Fonts a deck uses vs fonts installed on this Windows machine (system and per-user).
A missing font renders as substitute glyphs in PowerPoint. Install it per user: copy the files
to %LOCALAPPDATA%/Microsoft/Windows/Fonts, add HKCU Fonts registry values, restart PowerPoint.
For a deck that travels, save it with embedded fonts.

Fonts that appear only in speaker notes (notes pages and the theme of the notes master) are
listed apart: they never reach a slide, so a missing one is harmless. A typical case is Aptos
in the default notes master that python-pptx brings in.

    python fonts_check.py deck.pptx
"""
import posixpath
import re
import sys
import winreg
import zipfile

KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
SLIDE_PARTS = ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/")
NOTES_PARTS = ("ppt/notesSlides/", "ppt/notesMasters/", "ppt/handoutMasters/")


def _themes_of(z, master_dir):
    """Theme parts referenced by the masters in `master_dir` (e.g. 'ppt/notesMasters')."""
    out = set()
    for n in z.namelist():
        if n.startswith(master_dir + "/_rels/") and n.endswith(".rels"):
            for target in re.findall(rb'Target="([^"]*theme[^"]*)"', z.read(n)):
                out.add(posixpath.normpath(posixpath.join(master_dir, target.decode("utf-8"))))
    return out


def deck_fonts(path):
    """(slide fonts, notes-only fonts). Slide fonts: named on slides, layouts and slide masters,
    plus the Latin major/minor fonts of the slide masters' themes. The themes' per-script
    fallbacks (Thai, CJK, ...) are ignored: Latin and Cyrillic text never uses them."""
    slide, notes = set(), set()
    with zipfile.ZipFile(path) as z:
        slide_themes = _themes_of(z, "ppt/slideMasters")
        for n in z.namelist():
            if not n.endswith(".xml"):
                continue
            if n.startswith(SLIDE_PARTS + NOTES_PARTS):
                found = set(re.findall(rb'typeface="([^"+][^"]*)"', z.read(n)))
            elif n.startswith("ppt/theme/"):
                found = set(re.findall(rb'<a:latin typeface="([^"+][^"]*)"', z.read(n)))
            else:
                continue
            (slide if n.startswith(SLIDE_PARTS) or n in slide_themes else notes).update(found)
    dec = lambda names: {x.decode("utf-8") for x in names}
    slide = dec(slide)
    return slide, dec(notes) - slide


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
    ok = lambda f: any(h == f.lower() or h.startswith(f.lower() + " ") for h in have)
    slide, notes_only = deck_fonts(sys.argv[1])
    for f in sorted(slide):
        print(f"{'ok     ' if ok(f) else 'MISSING'}  {f}")
    for f in sorted(notes_only):
        print(f"{'ok     ' if ok(f) else 'notes  '}  {f}{'' if ok(f) else '  (only in speaker notes; slides unaffected)'}")
