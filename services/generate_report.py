"""
Sinh BẢNG ĐỀ NGHỊ THANH TOÁN từ Data.xlsx và Template.docx
Yêu cầu: pip install python-docx openpyxl lxml

Cấu trúc cột Data.xlsx (bắt đầu từ dòng 3):
  col 0: Mã GV
  col 1: Họ tên GV
  col 2: Trừ thuế TNCN  ("X" = trừ 10%)
  col 3: Tự luận        (số bài)
  col 4: Thực hành      (số bài)
  col 5: Tiểu luận      (số bài)
  col 6: Ra đề tự luận  (số đề)
  col 7: Duyệt đề tự luận (số đề)
"""

import copy, os
import openpyxl
import lxml.etree as etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# =================== CẤU HÌNH GIÁ ===================
GIA = {
    "tu_luan":   3000,
    "thuc_hanh": 1800,
    "tieu_luan": 3000,
    "ra_de":     90000,
    "duyet_de":  30000,
}

ACT_NAME = {
    "ra_de":     "Ra đề và đáp án thi viết 90-120 phút",
    "duyet_de":  "Duyệt đề và ĐA thi viết 90-120 phút",
    "tu_luan":   "Chấm bài thi hết HP tự luận bậc ĐH KHTN",
    "thuc_hanh": "Chấm bài thi thực hành bậc ĐH KHTN",
    "tieu_luan": "Chấm tiểu luận hết học phần bậc DH",
}

ACT_UNIT = {
    "ra_de":     "đề+ĐA",
    "duyet_de":  "đề",
    "tu_luan":   "bài",
    "thuc_hanh": "bài",
    "tieu_luan": "bài",
}

# =================== ĐỌC DỮ LIỆU ===================
def load_data(xlsx_path: str) -> list:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    def _int(v): return int(v) if v else 0
    teachers = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None:
            break
        teachers.append({
            "ma":        str(row[0]).zfill(4),
            "ten":       row[1],
            "co_thue":   str(row[2]).strip().upper() == "X" if row[2] else False,
            "tu_luan":   _int(row[3]),
            "thuc_hanh": _int(row[4]),
            "tieu_luan": _int(row[5]),
            "ra_de":     _int(row[6]),
            "duyet_de":  _int(row[7]),
        })
    teachers.sort(key=lambda x: x['ma'])
    return {
        "teachers": teachers,
        "nam_hoc" : ws["D1"].value,
        "hoc_ky" : ws["F1"].value
    }


# =================== HELPERS ===================
def fmt_money(val: int) -> str:
    return f"{val:,.0f}đ".replace(",", ".")


def compute_acts(t: dict) -> list:
    rows = []
    for key in ["ra_de", "duyet_de", "tu_luan", "thuc_hanh", "tieu_luan"]:
        qty = t[key]
        if qty == 0:
            continue
        vat = qty * GIA[key] * 0.1 if t["co_thue"] else 0
        rows.append({
            "act":  ACT_NAME[key],
            "qty":  qty,
            "unit": ACT_UNIT[key],
            "gia":  GIA[key],
            "tien": qty * GIA[key],
            "vat": vat,
            "thanh_tien": qty * GIA[key] - vat
        })
    return rows


def remove_highlight(tr):
    for rPr in tr.findall(".//" + qn("w:rPr")):
        hl = rPr.find(qn("w:highlight"))
        if hl is not None:
            rPr.remove(hl)

def set_tc_text(tc_elem, text: str, bold: bool = False):
    """Xóa toàn bộ runs cũ, tạo lại một run duy nhất để tránh text thừa."""
    for p in tc_elem.findall(qn("w:p")):
        first_run = p.find(qn("w:r"))
        old_rPr = None
        if first_run is not None:
            rp = first_run.find(qn("w:rPr"))
            if rp is not None:
                old_rPr = copy.deepcopy(rp)
                hl = old_rPr.find(qn("w:highlight"))
                if hl is not None:
                    old_rPr.remove(hl)

        # Xóa tất cả runs cũ
        for r in p.findall(qn("w:r")):
            p.remove(r)

        # Tạo run mới
        new_r = OxmlElement("w:r")
        if old_rPr is None:
            old_rPr = OxmlElement("w:rPr")

        b_el = old_rPr.find(qn("w:b"))
        if bold:
            if b_el is None: etree.SubElement(old_rPr, qn("w:b"))
        else:
            if b_el is not None: old_rPr.remove(b_el)

        new_r.append(old_rPr)

        t_el = OxmlElement("w:t")
        t_el.text = str(text)
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        new_r.append(t_el)
        p.append(new_r)
        return

def set_cols(tr, col_texts: dict, bold: bool = False):
    tcs = tr.findall(qn("w:tc"))
    for idx, text in col_texts.items():
        if idx < len(tcs):
            set_tc_text(tcs[idx], text, bold)

# =================== SINH REPORT ===================
def generate_report(xlsx_path: str, template_path: str):
    imported_data = load_data(xlsx_path)

    doc = Document(template_path)
    dt  = doc.tables[0].rows[0].cells[0].tables[1]

    # Clone các dòng mẫu từ template
    # Lưu ý cấu trúc merged cells:
    #   Dòng hoạt động (row 2): 8 cells  -> col 0..7
    #   Dòng "Cộng"    (row 4): 4 cells  -> col 0(label), 1(Tiền), 2(Thuế), 3(Thực lĩnh)
    #   Dòng "Cộng toàn bảng" (row 10): 4 cells -> tương tự
    tmpl_gv    = copy.deepcopy(dt.rows[1]._tr)
    tmpl_act   = copy.deepcopy(dt.rows[2]._tr)
    tmpl_cong  = copy.deepcopy(dt.rows[4]._tr)
    tmpl_grand = copy.deepcopy(dt.rows[10]._tr)

    tbl = dt._tbl

    for tr in tbl.findall(qn("w:tr"))[1:]:
        tbl.remove(tr)

    grand_total = 0
    grand_tax = 0

    for t in imported_data["teachers"]:
        acts = compute_acts(t)
        if not acts:
            continue

        subtotal = sum(a["tien"] for a in acts)
        t_tax = round(subtotal * 0.1) if t["co_thue"] else 0
        net = subtotal - t_tax
        grand_total += subtotal
        grand_tax += t_tax

        # Dòng tên giảng viên
        tr_gv = copy.deepcopy(tmpl_gv)
        remove_highlight(tr_gv)
        tcs = tr_gv.findall(qn("w:tc"))
        if tcs:
            set_tc_text(tcs[0], f"♦ {t['ma']} - {t['ten']}", bold=False)
        tbl.append(tr_gv)

         # Dòng hoạt động (8 cells)
        for stt, a in enumerate(acts, start=1):
            tr_a = copy.deepcopy(tmpl_act)
            remove_highlight(tr_a)
            set_cols(tr_a, {
                0: str(stt),
                1: "1",
                2: a["act"],
                3: f"{a['qty']} ({a['unit']})",
                4: fmt_money(a["gia"]),
                5: fmt_money(a["tien"]),
                6: fmt_money(a["vat"]),
                7: fmt_money(a["thanh_tien"]),
            })
            tbl.append(tr_a)

        # Dòng "Cộng" (4 cells merged)
        tr_c = copy.deepcopy(tmpl_cong)
        remove_highlight(tr_c)
        set_cols(tr_c, {
            1: fmt_money(subtotal),
            2: fmt_money(t_tax),
            3: fmt_money(net),
        }, bold=True)
        tbl.append(tr_c)

    tr_g = copy.deepcopy(tmpl_grand)
    remove_highlight(tr_g)
    set_cols(tr_g, {
        1: fmt_money(grand_total),
        2: fmt_money(grand_tax),
        3: fmt_money(grand_total - grand_tax),
    }, bold=True)
    tbl.append(tr_g)
    output_path = f"FIT_TTHDK_HK{imported_data["hoc_ky"]}_{imported_data["nam_hoc"].replace(" ", "")}.docx"
    doc.save(output_path)

    # ===== SAVE TO MEMORY =====
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if __name__ == "__main__":
    generate_report(
        xlsx_path="Data.xlsx",
        template_path="Template.docx"
    )