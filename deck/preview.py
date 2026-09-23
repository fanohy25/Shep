#!/usr/bin/env python3
"""
Rendu approximatif du deck en PNG avec Pillow.
Outil de contrôle qualité interne — ne fait pas partie du livrable.
Usage : python3 preview.py [fichier.pptx] [dossier_sortie]
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.oxml.ns import qn

SRC = sys.argv[1] if len(sys.argv) > 1 else "../Shep_Presentation_Strategique_2026.pptx"
OUTDIR = sys.argv[2] if len(sys.argv) > 2 else "/tmp/preview"
os.makedirs(OUTDIR, exist_ok=True)

prs = Presentation(SRC)
SW, SH = prs.slide_width, prs.slide_height
W = 1400
SCALE = W / SW
H = int(SH * SCALE)

FDIR = "/usr/share/fonts"
def find_font(bold=False):
    cands = (["DejaVuSans-Bold.ttf"] if bold else ["DejaVuSans.ttf"])
    for root, _, files in os.walk(FDIR):
        for c in cands:
            if c in files:
                return os.path.join(root, c)
    return None

FB, FR = find_font(True), find_font(False)
_cache = {}
def font(size, bold):
    key = (int(size), bold)
    if key not in _cache:
        path = FB if bold else FR
        _cache[key] = ImageFont.truetype(path, max(6, int(size))) if path \
            else ImageFont.load_default()
    return _cache[key]


def px(v):
    return v * SCALE


def get_fill(sp):
    spPr = sp._element.spPr
    g = spPr.find(qn('a:gradFill'))
    if g is not None:
        c = g.find('.//' + qn('a:srgbClr'))
        if c is not None:
            return tuple(int(c.get('val')[i:i+2], 16) for i in (0, 2, 4)) + (255,)
    sf = spPr.find(qn('a:solidFill'))
    if sf is not None:
        c = sf.find(qn('a:srgbClr'))
        if c is not None:
            rgb = tuple(int(c.get('val')[i:i+2], 16) for i in (0, 2, 4))
            a = c.find(qn('a:alpha'))
            alpha = int(int(a.get('val')) / 1000 * 2.55) if a is not None else 255
            return rgb + (alpha,)
    return None


def wrap(draw, txt, fnt, maxw):
    words, lines, cur = txt.split(), [], ""
    for w_ in words:
        t = (cur + " " + w_).strip()
        if draw.textlength(t, font=fnt) <= maxw or not cur:
            cur = t
        else:
            lines.append(cur); cur = w_
    if cur:
        lines.append(cur)
    return lines


for idx, slide in enumerate(prs.slides, 1):
    img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    for sp in slide.shapes:
        if sp.left is None:
            continue
        x, y, w, h = px(sp.left), px(sp.top), px(sp.width), px(sp.height)
        layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        if sp.shape_type == 13:  # picture
            d.rectangle([x, y, x + w, y + h], fill=(18, 34, 54, 255))
        elif getattr(sp, "has_chart", False) and sp.has_chart:
            d.rectangle([x, y, x + w, y + h], fill=(219, 231, 245, 255),
                        outline=(140, 175, 215, 255))
            d.text((x + w / 2, y + h / 2), "[GRAPHIQUE]", fill=(50, 100, 160),
                   font=font(14, True), anchor="mm")
        else:
            f = get_fill(sp)
            if f:
                d.rounded_rectangle([x, y, x + w, y + h],
                                    radius=min(8, w / 2, h / 2) if 'ROUNDED' in
                                    str(sp.shape_type) else 0, fill=f)
        img = Image.alpha_composite(img, layer)

        if sp.has_text_frame and sp.text_frame.text.strip():
            d = ImageDraw.Draw(img)
            cy = y
            for p in sp.text_frame.paragraphs:
                if not p.runs:
                    continue
                al = p.alignment.value if p.alignment else 1
                # composition run par run sur une même ligne logique
                segs = []
                for r in p.runs:
                    size = (r.font.size.pt if r.font.size else 12) * 12700 * SCALE
                    col = (17, 17, 17)
                    try:
                        col = tuple(r.font.color.rgb)
                    except Exception:
                        pass
                    segs.append((r.text, font(size, bool(r.font.bold)), col, size))
                full = "".join(s[0] for s in segs)
                fnt = segs[0][1]
                lines = wrap(d, full, fnt, max(20, w))
                lh = segs[0][3] * (p.line_spacing or 1.25)
                for ln in lines:
                    tw = d.textlength(ln, font=fnt)
                    tx = x if al == 1 else (x + (w - tw) / 2 if al == 2 else x + w - tw)
                    d.text((tx, cy), ln, font=fnt, fill=segs[0][2])
                    cy += lh
                cy += (p.space_after.pt * 12700 * SCALE) if p.space_after else 0
    img.convert("RGB").save(os.path.join(OUTDIR, f"slide_{idx:02d}.png"), quality=92)

print(f"{len(prs.slides._sldIdLst)} PNG -> {OUTDIR}")
