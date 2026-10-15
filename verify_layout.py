"""Verify the rendered PDF: bounds, overlap and text-clipping checks.

Run after converting the PPTX to PDF. Prints a report; exits non-zero on any
hard failure so the build can be gated on it.
"""

import sys

import pymupdf

PDF = "Будущее_труда_в_автоматизированной_экономике.pdf"
PW, PH = 959.98, 540.0
MARGIN = 0.62 * 72  # 44.6 pt
TOP = 0.30 * 72
BOTTOM = PH - 0.20 * 72

problems = []


def rects_overlap(a, b, tol=1.0):
    return (a.x0 < b.x1 - tol and b.x0 < a.x1 - tol and
            a.y0 < b.y1 - tol and b.y0 < a.y1 - tol)


def area(r):
    return max(0.0, r.width) * max(0.0, r.height)


def inter_area(a, b):
    x0, y0 = max(a.x0, b.x0), max(a.y0, b.y0)
    x1, y1 = min(a.x1, b.x1), min(a.y1, b.y1)
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def main():
    doc = pymupdf.open(PDF)
    print(f"pages: {doc.page_count}\n")

    for pno, page in enumerate(doc, start=1):
        blocks = [b for b in page.get_text("blocks") if b[6] == 0]
        text_blocks = []
        for x0, y0, x1, y1, txt, *_ in blocks:
            txt = txt.strip()
            if not txt:
                continue
            text_blocks.append((pymupdf.Rect(x0, y0, x1, y1), txt))

        images = [pymupdf.Rect(r) for r in
                  [page.get_image_bbox(i) for i in page.get_images(full=True)]
                  if r is not None and r.width > 0]
        # Deduplicate: LibreOffice may emit nested image rects.
        uniq = []
        for r in images:
            if not any(abs(r.x0 - u.x0) < 1 and abs(r.y0 - u.y0) < 1 and
                       abs(r.x1 - u.x1) < 1 and abs(r.y1 - u.y1) < 1 for u in uniq):
                uniq.append(r)
        images = uniq

        issues = []
        for r, txt in text_blocks:
            label = txt.replace("\n", " ")[:48]
            if r.x0 < MARGIN - 2 or r.x1 > PW - MARGIN + 2:
                issues.append(f"  [H-OUT] x {r.x0:.0f}-{r.x1:.0f} :: {label}")
            if r.y0 < TOP - 2 or r.y1 > BOTTOM + 2:
                issues.append(f"  [V-OUT] y {r.y0:.0f}-{r.y1:.0f} :: {label}")
        for r in images:
            if r.x0 < MARGIN - 3 or r.x1 > PW - MARGIN + 3 or r.y0 < TOP or r.y1 > BOTTOM + 3:
                issues.append(f"  [IMG-OUT] {r.x0:.0f},{r.y0:.0f} -> {r.x1:.0f},{r.y1:.0f}")

        # Text-on-text collisions (ignore tiny overlaps).
        for i in range(len(text_blocks)):
            for j in range(i + 1, len(text_blocks)):
                ra, ta = text_blocks[i]
                rb, tb = text_blocks[j]
                ov = inter_area(ra, rb)
                if ov > 12 and ov / min(area(ra), area(rb)) > 0.18:
                    issues.append(
                        f"  [OVERLAP] {ov:.0f}pt² :: '{ta.replace(chr(10),' ')[:26]}' vs '{tb.replace(chr(10),' ')[:26]}'")

        # Text sitting on top of a chart image (only flag when the chart is wide
        # enough that it is not an intentional label strip).
        for r, txt in text_blocks:
            for im in images:
                ov = inter_area(r, im)
                if ov > 0.35 * area(r) and im.width > 120 and im.height > 60:
                    issues.append(f"  [TEXT-ON-IMAGE] {ov:.0f}pt² :: {txt.replace(chr(10),' ')[:34]}")

        print(f"--- slide {pno}: {len(text_blocks)} text blocks, {len(images)} images")
        for line in issues:
            print(line)
        if not issues:
            print("  ok")
        problems.extend(issues)

    print("\nTOTAL ISSUES:", len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
