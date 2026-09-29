# -*- coding: utf-8 -*-
"""xlsx → PDF → PNG(쪽별). 화면 확인용."""
import subprocess, os, sys, pymupdf
import tempfile
W = os.path.join(tempfile.gettempdir(), "supervision_work")
os.makedirs(W, exist_ok=True)


def render(xlsx, tag, dpi=70):
    out = os.path.join(W, "pdf"); os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out, xlsx], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=200)
    pdf = os.path.join(out, os.path.basename(xlsx).replace(".xlsx", ".pdf"))
    doc = pymupdf.open(pdf)
    imgs = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        p = os.path.join(W, f"{tag}_p{i+1:02d}.png")
        pix.save(p); imgs.append(p)
    return pdf, imgs, [pg.get_text()[:40].replace("\n", " ") for pg in doc]

if __name__ == "__main__":
    pdf, imgs, heads = render(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 70)
    for i, h in zip(imgs, heads): print(i, "|", h)


def review_copy(xlsx, out):
    """검토용: 모든 시트를 가로 1쪽 너비로 맞춰 PDF 쪽수를 줄인다(원본 서식은 그대로)."""
    import openpyxl
    from openpyxl.worksheet.properties import PageSetupProperties
    wb = openpyxl.load_workbook(xlsx)
    for ws in wb:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = "landscape"
        if ws.title in ("2.기록지",):
            ws.print_area = f"A1:{openpyxl.utils.get_column_letter(ws.max_column)}45"
        if ws.title == "3.요약":
            ws.print_area = "A1:AC70"
    wb.save(out)
    return out
