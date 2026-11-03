"""Tạo templates/outbound_forms.xlsx từ file 'Database outbound shipping plan.xlsx'.

Giữ nguyên các sheet biểu mẫu (kèm công thức VLOOKUP gốc, định dạng, ô gộp, vùng in, hình ảnh) và bỏ các sheet dữ liệu —
dữ liệu nay nằm trong cơ sở dữ liệu của hệ thống. Chạy lại khi bạn đổi mẫu biểu trong file Excel gốc:

    python tools/make_outbound_template.py "đường dẫn tới Database outbound shipping plan.xlsx"
"""
import sys
import warnings
from pathlib import Path

import openpyxl

FORMS = ["SI Form", "SI Non-DG", "DGD WH", "DGD air", "VGM", "Manifest", "CMA", "ZIM", "IMDG ZIM 2cont", "MSC", "ONE", "OOCL",
         "Hapag-L", "CVC", "PK", "INV"]


def extract_shape_drawings(src, out_dir):
    """openpyxl bỏ các hình vẽ (shape) — ví dụ khung viền đỏ của DGD WH. Lưu lại XML gốc của sheet có shape để ghép vào file xuất."""
    import re
    import zipfile
    z = zipfile.ZipFile(src)
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    target = {}
    for m in re.finditer(r"<Relationship [^>]*>", rels):
        i, t = re.search(r"Id=\"(rId\d+)\"", m.group()), re.search(r"Target=\"([^\"]+)\"", m.group())
        if i and t:
            target[i.group(1)] = t.group(1).lstrip("/")
    out_dir.mkdir(exist_ok=True)
    for old in out_dir.glob("*.xml"):
        old.unlink()
    for m in re.finditer(r"<sheet [^>]*name=\"([^\"]+)\"[^>]*r:id=\"(rId\d+)\"", wbx):
        name, rid = m.group(1), m.group(2)
        if name not in FORMS:
            continue
        sheet_path = target[rid] if target[rid].startswith("xl/") else "xl/" + target[rid]
        rel_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rel_path not in z.namelist():
            continue
        for dm in re.finditer(r"Target=\"([^\"]*drawings/[^\"]+)\"", z.read(rel_path).decode("utf8")):
            dpath = "xl/drawings/" + dm.group(1).split("/")[-1]
            xml = z.read(dpath).decode("utf8")
            if "<xdr:sp>" in xml or "<xdr:sp " in xml or "<xdr:grpSp" in xml:
                (out_dir / f"{name}.xml").write_text(xml, encoding="utf8")
                print("drawing saved for", name, len(xml), "bytes")


def main(src):
    warnings.filterwarnings("ignore")  # openpyxl bỏ qua ảnh định dạng .wmf (logo cũ)
    wb = openpyxl.load_workbook(src)
    for name in list(wb.sheetnames):
        if name not in FORMS:
            del wb[name]
    wb._sheets.sort(key=lambda ws: FORMS.index(ws.title))
    out = Path(__file__).resolve().parent.parent / "templates" / "outbound_forms.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    extract_shape_drawings(src, out.parent / "drawings")
    print("saved", out, out.stat().st_size, "bytes;", len(wb.sheetnames), "sheets")


if __name__ == "__main__":
    main(sys.argv[1])
