"""Contact sheet of rendered slides for a quick visual check.

    python sheet.py png out.png 33 36 40-42           # chosen slides, 2 per row
    python sheet.py png out.png all --cols 4 --scale 0.5
    python sheet.py png out.png 5 41 --strip bottom   # only the citation strip of each slide
"""
import argparse
import glob
import os
import re

from PIL import Image


def pick(pngs, spec):
    by_n = {int(re.findall(r"(\d+)\.png$", p)[0]): p for p in pngs}
    if spec == ["all"]:
        return [by_n[k] for k in sorted(by_n)]
    out = []
    for tok in spec:
        a, _, b = tok.partition("-")
        out += [by_n[i] for i in range(int(a), int(b or a) + 1) if i in by_n]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pngdir")
    ap.add_argument("out")
    ap.add_argument("slides", nargs="+")
    ap.add_argument("--cols", type=int, default=2)
    ap.add_argument("--scale", type=float, default=1.0)
    ap.add_argument("--strip", choices=["bottom", "top"], help="keep only a 17%% strip of each slide")
    a = ap.parse_args()
    ims = []
    for p in pick(sorted(glob.glob(os.path.join(a.pngdir, "s*.png"))), a.slides):
        im = Image.open(p).convert("RGB")
        if a.strip:
            h = int(im.height * 0.17)
            im = im.crop((0, im.height - h, im.width, im.height) if a.strip == "bottom" else (0, 0, im.width, h))
        if a.scale != 1.0:
            im = im.resize((int(im.width * a.scale), int(im.height * a.scale)))
        ims.append(im)
    cols = 1 if a.strip else a.cols
    w, h = ims[0].size
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (w * cols, h * rows), "white")
    for i, im in enumerate(ims):
        sheet.paste(im, ((i % cols) * w, (i // cols) * h))
    sheet.save(a.out)
    print(a.out, sheet.size, len(ims), "slides")


if __name__ == "__main__":
    main()
