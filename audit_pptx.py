"""Audit the built PPTX against its rendered PDF.

The PDF gives the *real* rendered text extents, so comparing them with the
textbox rectangles declared in the PPTX catches text that overflows its box,
runs off the slide, or collides with a neighbouring element.

Rendered extents are computed per text span and assigned to the textbox whose
text best contains them, which avoids false positives from the PDF merging
neighbouring blocks into one.

Also checks that every shape stays inside the safe margins.
"""

import re
import sys
import unicodedata

import pymupdf
from pptx import Presentation

PPTX = "Будущее_труда_в_автоматизированной_экономике.pptx"
PDF = "Будущее_труда_в_автоматизированной_экономике.pdf"

EMU_PT = 12700.0
PW, PH = 959.98, 540.0
MARGIN_PT = 0.62 * 72
SAFE_L, SAFE_R = MARGIN_PT, PW - MARGIN_PT
SAFE_T, SAFE_B = 0.24 * 72, PH - 0.14 * 72
PAGE_AREA = PW * PH

problems = []


def norm(s):
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u200b", "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def token_set(s):
    return set(re.findall(r"[\w%$,.]+", norm(s)))


def chart_xrefs(doc, page):
    """Return the xrefs of genuinely complex raster images.

    LibreOffice exports every solid rectangle (slide background, accent bars,
    cards, the blue table header) as an image too, so size alone is not enough
    to tell a chart from a filled box. Counting distinct colours does the job:
    a filled rectangle has one colour, a chart has many.
    """
    out = []
    for img in page.get_images(full=True):
        xref = img[0]
        try:
            pix = pymupdf.Pixmap(doc, xref)
        except Exception:
            continue
        if pix.n > 4:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        n = pix.n
        total = pix.width * pix.height
        step = max(1, total // 4000)
        colors = set()
        data = pix.samples
        for i in range(0, total, step):
            colors.add(data[i * n:i * n + 3])
            if len(colors) > 24:
                break
        if len(colors) > 24:
            out.append(xref)
    return out


def main():
    prs = Presentation(PPTX)
    doc = pymupdf.open(PDF)

    for idx, slide in enumerate(prs.slides):
        page = doc[idx]
        pno = idx + 1
        issues = []

        # --- rendered text spans
        spans = []
        for block in page.get_text("dict")["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for sp in line["spans"]:
                    t = sp["text"].strip()
                    if t:
                        spans.append((pymupdf.Rect(sp["bbox"]), t))

        # --- charts
        charts = []
        for xref in chart_xrefs(doc, page):
            for r in page.get_image_rects(xref):
                if r.width > 40 and not any(
                        abs(r.x0 - c.x0) < 1.5 and abs(r.y0 - c.y0) < 1.5 for c in charts):
                    charts.append(r)

        # --- collect declared textboxes
        boxes = []
        for shape in slide.shapes:
            box = pymupdf.Rect(shape.left / EMU_PT, shape.top / EMU_PT,
                               (shape.left + shape.width) / EMU_PT,
                               (shape.top + shape.height) / EMU_PT)
            if shape.__class__.__name__ == "Picture":
                if box.x0 < SAFE_L - 2 or box.x1 > SAFE_R + 2 or box.y0 < SAFE_T or box.y1 > SAFE_B + 2:
                    issues.append(f"  [PIC-BOUNDS] {box.x0:.0f},{box.y0:.0f}->{box.x1:.0f},{box.y1:.0f}")
                continue
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                if box.x0 < -1 or box.y0 < -1 or box.x1 > PW + 1 or box.y1 > PH + 1:
                    issues.append(f"  [RECT-BOUNDS] {box.x0:.0f},{box.y0:.0f}->{box.x1:.0f},{box.y1:.0f}")
                continue
            boxes.append((box, shape.text_frame.text))

        # Tables also carry text. Their cells must take part in span
        # assignment, otherwise table text gets attributed to a textbox that
        # merely shares vocabulary and produces bogus overflow reports.
        assign = list(boxes)
        for shape in slide.shapes:
            if shape.has_table:
                tbl = shape.table
                x = shape.left / EMU_PT
                y = shape.top / EMU_PT
                for r, row in enumerate(tbl.rows):
                    cy = y + sum(tbl.rows[k].height for k in range(r)) / EMU_PT
                    cx = x
                    for c, cell in enumerate(row.cells):
                        cw = tbl.columns[c].width / EMU_PT
                        ch = row.height / EMU_PT
                        assign.append((pymupdf.Rect(cx, cy, cx + cw, cy + ch), cell.text))
                        cx += cw

        # --- assign each span to the best-matching box
        rendered = {i: None for i in range(len(boxes))}
        for rect, txt in spans:
            st = token_set(txt)
            if not st:
                continue
            # Prefer the box that both shares tokens and geometrically contains
            # the span; several boxes share vocabulary (e.g. "Роботы:").
            bi, bs = None, 0.0
            for i, (box, btxt) in enumerate(assign):
                bt = token_set(btxt)
                if not bt:
                    continue
                tok = len(st & bt) / len(st)
                if tok < 0.5:
                    continue
                ov = (max(0, min(rect.x1, box.x1) - max(rect.x0, box.x0))
                      * max(0, min(rect.y1, box.y1) - max(rect.y0, box.y0)))
                contain = ov / max(1.0, rect.width * rect.height)
                # Tie-break towards the more specific (smaller) box: a table
                # cell shares vocabulary with a wide bullet list.
                sc = tok + 0.6 * contain - 1e-6 * (box.width * box.height)
                if sc > bs:
                    bi, bs = i, sc
            if bi is not None and bi < len(boxes):
                rendered[bi] = rect if rendered[bi] is None else (rendered[bi] | rect)

        for i, (box, btxt) in enumerate(boxes):
            label = norm(btxt)[:42]
            r = rendered[i]
            if r is None:
                issues.append(f"  [NOT-RENDERED] '{label}'")
                continue
            if r.x1 > box.x1 + 4:
                issues.append(f"  [H-OVERFLOW] text {r.x1:.0f} > box {box.x1:.0f} :: {label}")
            if r.y1 > box.y1 + 5:
                issues.append(f"  [V-OVERFLOW] text {r.y1:.0f} > box {box.y1:.0f} :: {label}")
            if r.x0 < SAFE_L - 3 or r.x1 > SAFE_R + 3:
                issues.append(f"  [H-MARGIN] {r.x0:.0f}-{r.x1:.0f} :: {label}")
            if r.y1 > SAFE_B + 3:
                issues.append(f"  [V-MARGIN] bottom {r.y1:.0f} :: {label}")

        # --- text overlapping a chart
        for rect, txt in spans:
            for im in charts:
                ov = (max(0, min(rect.x1, im.x1) - max(rect.x0, im.x0))
                      * max(0, min(rect.y1, im.y1) - max(rect.y0, im.y0)))
                if ov > 0.45 * rect.width * rect.height and ov > 40:
                    issues.append(f"  [TEXT-ON-CHART] {ov:.0f}pt² :: {norm(txt)[:36]}")

        # --- text/text collisions between different boxes
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = rendered[i], rendered[j]
                if a is None or b is None:
                    continue
                ov = (max(0, min(a.x1, b.x1) - max(a.x0, b.x0))
                      * max(0, min(a.y1, b.y1) - max(a.y0, b.y0)))
                if ov > 10:
                    issues.append(f"  [COLLISION] {ov:.0f}pt² :: "
                                  f"'{norm(boxes[i][1])[:24]}' vs '{norm(boxes[j][1])[:24]}'")

        print(f"--- slide {pno}: {len(spans)} spans, {len(charts)} charts, {len(boxes)} boxes")
        for line in issues:
            print(line)
        if not issues:
            print("  ok")
        problems.extend(issues)

    print("\nTOTAL ISSUES:", len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
