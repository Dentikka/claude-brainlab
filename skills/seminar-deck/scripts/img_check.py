"""Check downloaded images and re-encode them as pixels only (drops metadata and anything
appended after the end of the image). Run it on every meme or figure taken from the web.

    python img_check.py in1.jpg in2.png --out-dir figs/memes
"""
import argparse
import os
import struct
import sys

from PIL import Image

SUSPICIOUS = (b"<script", b"<?php", b"PK\x03\x04", b"<html", b"%PDF", b"#!/", b"This program")
PNG_STANDARD = {"IHDR", "IDAT", "IEND", "PLTE", "tRNS", "gAMA", "sRGB", "pHYs", "cHRM", "iCCP"}


def png_chunks(b: bytes):
    p, out = 8, []
    while p + 8 <= len(b):
        n, t = struct.unpack(">I4s", b[p:p + 8])
        out.append(t.decode("latin1"))
        p += 12 + n
    return out


def check(path: str) -> bool:
    """Decodes fully, nothing after the image end, no script/archive/PE signatures.
    A bare 'MZ' is not checked: two bytes occur by chance inside compressed pixel data."""
    b = open(path, "rb").read()
    im = Image.open(path)
    im.load()
    if im.format == "PNG":
        end = b.rfind(b"IEND") + 8
        extra = sorted(set(png_chunks(b)) - PNG_STANDARD)
    else:
        end = b.rfind(b"\xff\xd9") + 2
        extra = []
    trailing = len(b) - end
    hits = [k.decode("latin1") for k in SUSPICIOUS if k in b]
    ok = trailing == 0 and not hits
    print(f"{path}: {im.format} {im.size} {len(b)} B, trailing={trailing}, suspicious={hits}, "
          f"extra chunks={extra} -> {'ok' if ok else 'LOOK CLOSER'}")
    return ok


def reencode(path: str, out_dir: str) -> str:
    im = Image.open(path).convert("RGB")
    clean = Image.new("RGB", im.size)
    clean.paste(im)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".png")
    clean.save(out, "PNG", optimize=True)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out-dir", required=True)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    bad = 0
    for f in a.files:
        if check(f):
            print("  ->", reencode(f, a.out_dir))
        else:
            bad += 1
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
