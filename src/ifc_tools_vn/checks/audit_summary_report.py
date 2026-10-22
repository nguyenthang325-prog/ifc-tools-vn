"""
Tổng hợp kết quả BIM Audit để gửi lead discipline.
Đọc 3 file Excel từ thư mục kết quả (IDS_KetQua, LOD_Audit, IFC_QTO) sinh ra báo cáo tóm tắt.

Output: ifc_tools_output/<ngày>/AuditSummary_{project}_{discipline}_{stage}_{YYYYMMDD}.xlsx
  Sheet "Tong_Hop"    — Compliance IDS, LOD, bảng KL chính, kết luận
  Sheet "Loi_Uu_Tien" — Danh sách lỗi cần xử lý (ERROR trước, sort by count desc)

Cách dùng:
    python audit_summary_report.py --project Duplex --discipline KT --stage TKKT
    python audit_summary_report.py --project "TenDuAn" --discipline ALL --stage TKTC --top 15

    # Chỉ định file cụ thể thay vì tự tìm:
    python audit_summary_report.py --project Duplex --discipline KT --stage TKKT --ids IDS_KetQua_Duplex_20261022.xlsx --lod LOD_Audit_Duplex_TKKT_20261022.xlsx --qto IFC_QTO_Duplex_KT_20261022.xlsx
"""
import sys
import re
import argparse
from pathlib import Path
from datetime import date

sys.stdout.reconfigure(encoding='utf-8')

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Lỗi: pip install openpyxl")
    sys.exit(1)

from ifc_tools_vn.config import BASE, KET_QUA_DIR, OUTPUT_DIR  # đường dẫn: xem config.py

# ── Excel styles ──────────────────────────────────────────────────────────────
_H_FILL  = PatternFill("solid", fgColor="1F4E79")
_S_FILL  = PatternFill("solid", fgColor="2E75B6")
_OK_FILL = PatternFill("solid", fgColor="E2EFDA")
_WA_FILL = PatternFill("solid", fgColor="FFF2CC")
_ER_FILL = PatternFill("solid", fgColor="FCE4D6")
_NA_FILL = PatternFill("solid", fgColor="F2F2F2")
_thin    = Side(style="thin", color="CCCCCC")
_border  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_center  = Alignment(horizontal="center", vertical="center", wrap_text=True)
_left    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_H_FONT  = Font(bold=True, color="FFFFFF", size=10)
_B_FONT  = Font(bold=True, size=10)
_N_FONT  = Font(size=10)
_S_FONT  = Font(size=9)
_I_FONT  = Font(italic=True, size=9)


def _c(ws, row, col, value=None, fill=None, font=None, align=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.border    = _border
    cell.fill      = fill or PatternFill()
    cell.font      = font or _N_FONT
    cell.alignment = align or _center
    return cell


def _col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── File discovery ────────────────────────────────────────────────────────────

def _find_latest(prefix: str, slug: str) -> Path | None:
    """Tìm file mới nhất khớp tiền tố: ưu tiên khớp slug project, fallback về prefix chung."""
    for want_slug in (True, False):
        pattern = f"{prefix}*{slug}*.xlsx" if want_slug else f"{prefix}*.xlsx"
        candidates = list(KET_QUA_DIR.rglob(pattern))
        if candidates:
            return sorted(candidates, key=lambda f: f.stat().st_mtime, reverse=True)[0]
    return None


# ── Excel readers ─────────────────────────────────────────────────────────────

def _pct(val) -> float | None:
    """'87.5%' hoặc 87.5 hoặc None → float hoặc None."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("%", "")
    try:
        return float(s)
    except ValueError:
        return None


def _nonempty_rows(ws, min_row: int):
    return [r for r in ws.iter_rows(min_row=min_row, values_only=True)
            if any(v is not None for v in r)]


def read_ids(path: Path) -> dict:
    """
    Đọc kết quả engine chuẩn buildingSMART (ids_run.py). Sheet "TongHop_Spec":
    STT, Mã EIR, Yêu cầu, Áp dụng, Đạt, Trượt, % Đạt (row1=title, row2=headers, row4+=data).
    Không có khái niệm severity ERROR/WARNING.
    CẢNH BÁO (đã kiểm chứng 29/09/2026): KHÔNG phải mọi yêu cầu đều 'required'. Spec do
    eir_to_ids.py sinh ra mang minOccurs=0 (mặc định của ifctester) = 'optional', nên spec
    có 0 cấu kiện áp dụng vẫn được tính ĐẠT. Vì vậy phải đọc thêm out["khong_kiem"] /
    out["khong_phu"], đừng kết luận chỉ từ pct.

    Trả về: pct, total, passed, errors, warnings (luôn 0), top_rules (sorted by failed desc).
    """
    out = {"pct": None, "total": 0, "passed": 0, "errors": 0, "warnings": 0,
           "top_rules": [], "file": path.name}
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠ Không đọc được IDS file: {e}")
        return out

    if "TongHop_Spec" not in wb.sheetnames:
        print(f"  ⚠ File IDS không đúng định dạng engine chuẩn (thiếu sheet TongHop_Spec): {path.name}")
        wb.close()
        return out

    rows = _nonempty_rows(wb["TongHop_Spec"], min_row=4)
    total = passed = 0
    for row in rows:
        if len(row) < 6:
            continue
        code   = row[1] or ""
        name   = row[2] or ""
        appl   = row[3] or 0
        psd    = row[4] or 0
        failed = row[5] or 0
        total  += appl
        passed += psd
        if failed:
            out["top_rules"].append({
                "rule_id":  code,
                "severity": "ERROR",
                "name":     name,
                "failed":   int(failed),
                "total":    int(appl),
            })
    out["total"]    = total
    out["passed"]   = passed
    out["errors"]   = total - passed
    out["warnings"] = 0
    out["pct"]      = round(passed / total * 100, 1) if total else None
    out["top_rules"].sort(key=lambda r: -r["failed"])

    # Yêu cầu có 0 cấu kiện áp dụng: chuẩn IDS tính là ĐẠT, nhưng thực chất KHÔNG kiểm gì.
    # Đếm ở đây để báo cáo tổng hợp không kết luận "đạt" trên phần chưa hề kiểm.
    out["khong_kiem"] = sum(1 for row in rows if len(row) >= 4 and not (row[3] or 0))
    # Số cấu kiện không yêu cầu nào phủ — lấy từ sheet DoPhu_CanhBao do ids_run.py sinh.
    out["khong_phu"] = None
    if "DoPhu_CanhBao" in wb.sheetnames:
        tong = 0
        for row in _nonempty_rows(wb["DoPhu_CanhBao"], min_row=2):
            if len(row) >= 3 and isinstance(row[0], int) and isinstance(row[2], int):
                tong += row[2]
        out["khong_phu"] = tong
    wb.close()
    return out


def read_lod(path: Path) -> dict:
    """
    Trả về:
      pct, total, ok, fail, required, by_group (list sorted by fail desc)
    Sheet "LOD_Audit": row3=headers, row4+=data
    Sheet "Thong_Ke" : row1=title, row3=headers, row4+=data, last=TỔNG
    """
    out = {"pct": None, "total": 0, "ok": 0, "fail": 0, "required": "—",
           "by_group": [], "file": path.name}
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠ Không đọc được LOD file: {e}")
        return out

    # Lấy LOD yêu cầu từ LOD_Audit sheet (col F = LOD YC)
    if "LOD_Audit" in wb.sheetnames:
        for row in _nonempty_rows(wb["LOD_Audit"], min_row=4):
            if len(row) >= 6 and row[5] is not None:
                out["required"] = str(int(row[5])) if isinstance(row[5], float) else str(row[5])
                break

    if "Thong_Ke" in wb.sheetnames:
        rows = _nonempty_rows(wb["Thong_Ke"], min_row=4)
        if rows:
            tot = rows[-1]   # last row = TỔNG
            out["total"] = tot[1] or 0
            out["ok"]    = tot[2] or 0
            out["fail"]  = tot[3] or 0
            out["pct"]   = _pct(tot[4])
            for row in rows[:-1]:
                nhom, total_g, ok_g, fail_g, pct_g = row[0], row[1], row[2], row[3], row[4]
                if nhom and fail_g:
                    out["by_group"].append({
                        "group": nhom,
                        "fail":  int(fail_g),
                        "total": int(total_g or 0),
                        "pct":   _pct(pct_g),
                    })
            out["by_group"].sort(key=lambda r: -r["fail"])

    wb.close()
    return out


def read_qto(path: Path) -> list:
    """
    Trả về list {group, prop, unit, total, n} từ sheet "Tom_Tat_BoQ".
    Row1=title, row2=note, row3=headers, row4+=data
    """
    items = []
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:
        print(f"  ⚠ Không đọc được QTO file: {e}")
        return items

    if "Tom_Tat_BoQ" in wb.sheetnames:
        for row in _nonempty_rows(wb["Tom_Tat_BoQ"], min_row=4):
            if len(row) < 5 or not row[1]:
                continue
            try:
                total_kl = float(row[4]) if row[4] is not None else None
            except (TypeError, ValueError):
                total_kl = None
            items.append({
                "group": row[1],
                "prop":  row[2] or "",
                "unit":  row[3] or "",
                "total": total_kl,
                "n":     int(row[5] or 0),
            })

    wb.close()
    return items


# ── Sheet builders ────────────────────────────────────────────────────────────

def _sheet_tong_hop(wb, ids: dict, lod: dict, qto: list,
                    project: str, discipline: str, stage: str,
                    ids_threshold: int = 90):
    ws = wb.active
    ws.title = "Tong_Hop"
    _col_widths(ws, [4, 34, 16, 42])
    ws.row_dimensions[1].height = 38

    # ── Tiêu đề chính ──
    ws.merge_cells("A1:D1")
    t = ws.cell(1, 1,
        f"BÁO CÁO TỔNG HỢP BIM AUDIT — {project.upper()}"
        f"  |  {discipline}  |  {stage}  |  {date.today().strftime('%d/%m/%Y')}"
    )
    t.fill = _H_FILL; t.font = Font(bold=True, color="FFFFFF", size=12)
    t.alignment = _center; t.border = _border

    row = 3

    def _section_hdr(label):
        nonlocal row
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row, 1, label)
        c.fill = _S_FILL; c.font = _H_FONT; c.alignment = _left; c.border = _border
        ws.row_dimensions[row].height = 22
        row += 1

    def _kv(label, value, fill=None):
        nonlocal row
        _c(ws, row, 1, "", fill=_NA_FILL)
        ws.merge_cells(f"B{row}:C{row}")
        _c(ws, row, 2, label, font=_B_FONT, align=_left)
        _c(ws, row, 4, value, fill=fill, font=_B_FONT)
        ws.row_dimensions[row].height = 20
        row += 1

    def _blank():
        nonlocal row
        ws.row_dimensions[row].height = 10
        row += 1

    # ── Khối IDS ──
    _section_hdr("1. KIỂM TRA THÔNG TIN (IDS / QD347 — 21 rules)")
    if ids["total"]:
        pct_ids = ids["pct"]
        ids_status_fill = _OK_FILL if (pct_ids or 0) >= ids_threshold and ids["errors"] == 0 else _ER_FILL
        _kv("Compliance tổng",    f"{pct_ids}% (ngưỡng đạt: ≥{ids_threshold}%)" if pct_ids is not None else "N/A",
            fill=ids_status_fill)
        _kv("Checks thực hiện",   f"{ids['total']} checks ({ids['passed']} passed)")
        _kv("Lỗi ERROR (bắt buộc)", str(ids["errors"]),
            fill=_ER_FILL if ids["errors"] else _OK_FILL)
        _kv("Cảnh báo WARNING",  str(ids["warnings"]),
            fill=_WA_FILL if ids["warnings"] else _OK_FILL)
        verdict_ids = "✅ ĐẠT" if (pct_ids or 0) >= ids_threshold and ids["errors"] == 0 else "❌ CHƯA ĐẠT"
        _kv("Kết luận IDS",      verdict_ids,
            fill=_OK_FILL if "ĐẠT" in verdict_ids else _ER_FILL)
        _kv("File nguồn",        ids.get("file", "—"), fill=None)
    else:
        _kv("Trạng thái", "Không tìm thấy file IDS_Check", fill=_NA_FILL)
    _blank()

    # ── Khối LOD ──
    _section_hdr("2. KIỂM TRA LOD (QD347 / AIA G202)")
    if lod["total"]:
        pct_lod = lod["pct"]
        lod_status_fill = _OK_FILL if (pct_lod or 0) == 100 else (
            _WA_FILL if (pct_lod or 0) >= 80 else _ER_FILL)
        _kv("LOD yêu cầu",    f"LOD {lod['required']}")
        _kv("Tỷ lệ đạt LOD",  f"{pct_lod}%" if pct_lod is not None else "N/A",
            fill=lod_status_fill)
        _kv("Tổng elements",  str(lod["total"]))
        _kv("Không đạt LOD",  str(lod["fail"]),
            fill=_ER_FILL if lod["fail"] else _OK_FILL)
        verdict_lod = "✅ ĐẠT" if lod["fail"] == 0 else "❌ CHƯA ĐẠT"
        _kv("Kết luận LOD",   verdict_lod,
            fill=_OK_FILL if "ĐẠT" in verdict_lod else _ER_FILL)
        _kv("File nguồn",     lod.get("file", "—"))
    else:
        _kv("Trạng thái", "Không tìm thấy file LOD_Audit", fill=_NA_FILL)
    _blank()

    # ── Khối QTO ──
    _section_hdr("3. KHỐI LƯỢNG CHÍNH (IFC QTO — BaseQuantities)")
    if qto:
        ws.merge_cells(f"A{row}:D{row}")
        _c(ws, row, 1, "⚠ Đơn vị theo project units trong IFC (Revit SI → m / m² / m³). "
           "Kiểm tra lại nếu file dùng mm hoặc imperial.",
           font=Font(italic=True, color="C00000", size=9), align=_left)
        ws.row_dimensions[row].height = 18
        row += 1

        for j, h in enumerate(["", "Nhóm element", "Đại lượng", "Tổng KL (ĐVT)"], 1):
            _c(ws, row, j, h, fill=_S_FILL, font=_H_FONT,
               align=_left if j > 1 else _center)
        ws.row_dimensions[row].height = 20
        row += 1

        for item in qto[:12]:
            total_str = (f"{item['total']:,.2f} {item['unit']}"
                         if item["total"] is not None else "—")
            _c(ws, row, 1, "", fill=_NA_FILL)
            _c(ws, row, 2, item["group"], font=_N_FONT, align=_left)
            _c(ws, row, 3, item["prop"],  font=_S_FONT, align=_left)
            _c(ws, row, 4, total_str,     font=_B_FONT)
            ws.row_dimensions[row].height = 18
            row += 1
    else:
        _kv("Trạng thái", "Không tìm thấy file IFC_QTO hoặc IFC không có BaseQuantities", fill=_NA_FILL)
    _blank()

    # ── Kết luận tổng ──
    ids_ok = ids["total"] > 0 and (ids["pct"] or 0) >= ids_threshold and ids["errors"] == 0
    lod_ok = lod["total"] > 0 and lod["fail"] == 0
    all_ok = ids_ok and lod_ok

    _section_hdr("4. KẾT LUẬN VÀ ĐỀ NGHỊ")
    verdict_fill = _OK_FILL if all_ok else _ER_FILL
    verdict_txt  = (
        f"✅ ĐẠT — Mô hình {discipline} đủ điều kiện nộp hồ sơ {stage}"
        if all_ok else
        f"❌ CHƯA ĐẠT — Cần xử lý lỗi trước khi nộp hồ sơ {stage}"
    )
    ws.merge_cells(f"A{row}:D{row}")
    c_v = ws.cell(row, 1, verdict_txt)
    c_v.fill = verdict_fill; c_v.font = Font(bold=True, size=11)
    c_v.alignment = _left; c_v.border = _border
    ws.row_dimensions[row].height = 28
    row += 1

    if not ids_ok and ids["errors"] > 0:
        ws.merge_cells(f"A{row}:D{row}")
        c_r = ws.cell(row, 1,
            f"→ Ưu tiên 1: Sửa {ids['errors']} lỗi ERROR trong IDS Check "
            "(thiếu thông tin bắt buộc — xem sheet Loi_Uu_Tien)")
        c_r.fill = _ER_FILL; c_r.font = Font(bold=True, size=10)
        c_r.alignment = _left; c_r.border = _border
        ws.row_dimensions[row].height = 22
        row += 1

    if not lod_ok and lod["fail"] > 0:
        ws.merge_cells(f"A{row}:D{row}")
        c_r2 = ws.cell(row, 1,
            f"→ Ưu tiên 2: Bổ sung thông tin LOD cho {lod['fail']} elements "
            f"(cần LOD {lod['required']}) trong Revit → export IFC lại")
        c_r2.fill = _WA_FILL; c_r2.font = Font(bold=True, size=10)
        c_r2.alignment = _left; c_r2.border = _border
        ws.row_dimensions[row].height = 22
        row += 1


def _sheet_loi_uu_tien(wb, ids: dict, lod: dict, top_n: int,
                        project: str, discipline: str, stage: str):
    ws = wb.create_sheet("Loi_Uu_Tien")
    _col_widths(ws, [4, 10, 10, 52, 10, 30])

    ws.merge_cells("A1:F1")
    t = ws.cell(1, 1,
        f"LỖI ƯU TIÊN XỬ LÝ — {project.upper()} | {discipline} | {stage}"
        f"  |  Ngày: {date.today().strftime('%d/%m/%Y')}"
    )
    t.fill = PatternFill("solid", fgColor="C00000")
    t.font = Font(bold=True, color="FFFFFF", size=12)
    t.alignment = _center; t.border = _border
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:F2")
    s = ws.cell(2, 1,
        "Sắp xếp: ERROR trước (bắt buộc sửa) → WARNING sau. "
        "Số lượng = số element bị lỗi theo rule/nhóm đó.")
    s.font = _I_FONT; s.alignment = _left
    ws.row_dimensions[2].height = 16

    headers = ["#", "Nguồn", "Mức", "Mô tả lỗi", "Số element", "Hành động đề xuất"]
    for j, h in enumerate(headers, 1):
        _c(ws, 3, j, h, fill=_S_FILL, font=_H_FONT)
    ws.row_dimensions[3].height = 22
    ws.freeze_panes = "A4"

    items = []

    # IDS errors first
    for r in ids.get("top_rules", []):
        action = _ids_action(r["rule_id"], r["name"])
        items.append({
            "src":    "IDS Check",
            "sev":    r["severity"],
            "desc":   f"[{r['rule_id']}] {r['name']}",
            "count":  r["failed"],
            "action": action,
        })

    # LOD errors
    for g in lod.get("by_group", []):
        items.append({
            "src":   "LOD Audit",
            "sev":   "ERROR" if (g["pct"] or 0) < 80 else "WARNING",
            "desc":  f"Nhóm '{g['group']}': {g['fail']}/{g['total']} elements không đạt LOD {lod.get('required','?')}",
            "count": g["fail"],
            "action": (
                f"Bổ sung Pset vào Revit (Name, Material, Pset theo loại); "
                f"export IFC lại với 'Export base quantities'"
            ),
        })

    # Trim to top_n
    items = items[:top_n]

    if not items:
        ws.merge_cells("A4:F4")
        _c(ws, 4, 1, "✅ Không có lỗi — tất cả kiểm tra đạt.",
           fill=_OK_FILL, font=_B_FONT, align=_left)
        return

    for i, it in enumerate(items, 1):
        sev_fill  = _ER_FILL if it["sev"] == "ERROR" else _WA_FILL
        sev_color = "C00000" if it["sev"] == "ERROR" else "7F6000"
        r = i + 3
        _c(ws, r, 1, i,            font=_N_FONT)
        _c(ws, r, 2, it["src"],    font=Font(size=9, italic=True), align=_left)
        _c(ws, r, 3, it["sev"],    fill=sev_fill,
           font=Font(bold=True, size=9, color=sev_color))
        _c(ws, r, 4, it["desc"],   font=_N_FONT, align=_left)
        _c(ws, r, 5, it["count"],  font=_B_FONT,
           fill=sev_fill if it["count"] else PatternFill())
        _c(ws, r, 6, it["action"], font=_S_FONT, align=_left)
        ws.row_dimensions[r].height = 24


def _ids_action(rule_id: str, rule_name: str) -> str:
    """Gợi ý hành động dựa trên rule ID."""
    r = rule_id.upper()
    if "PRJ-001" in r or "PRJ-002" in r:
        return "Mở IFC trong text editor, điền Name cho IfcProject/IfcBuilding"
    if "CMN-001" in r:
        return "Export IFC lại từ Revit để reset GUID (Revit tự tạo GUID hợp lệ)"
    if "CMN-002" in r:
        return "Điền tên cho element trong Revit trước khi export IFC"
    if "CMN-003" in r:
        return "Element không có geometry — xóa placeholder hoặc bổ sung hình học trong Revit"
    if "CMN-004" in r:
        return "Gán vật liệu cho element trong Revit → Material → Assign"
    if "FIRERATING" in rule_name.upper() or "KT-003" in r or "KT-005" in r:
        return "Điền FireRating vào Pset trong Revit: Element Properties → Psets → FireRating"
    if "LOADBEARING" in rule_name.upper() or "KC-002" in r or "KC-004" in r or "KC-005" in r:
        return "Điền LoadBearing (True/False) vào Pset trong Revit: Element Properties → Psets"
    if "PSET_WALL" in rule_name.upper() or "KT-001" in r:
        return "Bổ sung Pset_WallCommon vào tường trong Revit (dùng Shared Parameters hoặc Project Parameters)"
    if "KT-004" in r or "DOOR" in rule_name.upper():
        return "Bổ sung Pset_DoorCommon vào cửa trong Revit"
    if "KT-006" in r or "WINDOW" in rule_name.upper():
        return "Bổ sung Pset_WindowCommon vào cửa sổ trong Revit"
    if "MEP" in r:
        return "Điền Pset thiết bị MEP tương ứng trong Revit (Mechanical/Electrical properties)"
    return f"Xem sheet chi tiết IDS_KetQua_*.xlsx → Loi_Chi_Tiet để xác định element cụ thể"


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Tổng hợp kết quả BIM Audit gửi lead discipline"
    )
    parser.add_argument("--project",    required=True,
                        help="Tên dự án (dùng để tìm file trong thư mục kết quả và đặt tên output)")
    parser.add_argument("--discipline", default="ALL",
                        choices=["KT", "KC", "MEP", "ALL"],
                        help="Discipline (mặc định: ALL)")
    parser.add_argument("--stage",      default="TKKT",
                        choices=["TKCS", "TKKT", "TKTC", "AS-BUILT"],
                        help="Giai đoạn (mặc định: TKKT)")
    parser.add_argument("--top",        type=int, default=10,
                        help="Số lỗi tối đa trong sheet Loi_Uu_Tien (mặc định: 10)")
    parser.add_argument("--ids",        default=None,
                        help="Chỉ định trực tiếp file IDS_Check (bỏ qua auto-detect)")
    parser.add_argument("--lod",        default=None,
                        help="Chỉ định trực tiếp file LOD_Audit (bỏ qua auto-detect)")
    parser.add_argument("--qto",        default=None,
                        help="Chỉ định trực tiếp file IFC_QTO (bỏ qua auto-detect)")
    parser.add_argument("--ids-threshold", type=int, default=90,
                        help="Ngưỡng %% compliance IDS để đạt (mặc định: 90)")
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    slug = re.sub(r"[^\w]", "_", args.project)

    print(f"\n{'='*60}")
    print(f"  AUDIT SUMMARY REPORT — {args.project}")
    print(f"  Discipline : {args.discipline}  |  Stage: {args.stage}")
    print(f"{'='*60}")

    # ── Tìm file input ──
    def _resolve(cli_arg, prefix):
        if cli_arg:
            p = Path(cli_arg)
            if not p.is_absolute():
                p = BASE / cli_arg
            if not p.exists():
                print(f"  ⚠ File không tồn tại: {p}")
                return None
            return p
        found = _find_latest(prefix, slug)
        label = prefix.rstrip("_")
        if found:
            print(f"  ✓ {label}: {found.name}")
        else:
            print(f"  ⚠ Không tìm thấy {label}*.xlsx trong {KET_QUA_DIR}")
        return found

    ids_path = _resolve(args.ids, "IDS_KetQua_")
    lod_path = _resolve(args.lod, "LOD_Audit_")
    qto_path = _resolve(args.qto, "IFC_QTO_")

    if not any([ids_path, lod_path, qto_path]):
        print("\n  Lỗi: Không có file nào để đọc.")
        print("  → Chạy bim_audit_pipeline.py hoặc eir_to_ids+ids_run, lod_checker, ifc_qto trước.")
        sys.exit(1)

    # ── Đọc dữ liệu ──
    print("\n  Đang đọc kết quả audit...")
    _ids_empty = {"pct": None, "total": 0, "passed": 0, "errors": 0,
                  "warnings": 0, "top_rules": [], "file": "—"}
    _lod_empty = {"pct": None, "total": 0, "ok": 0, "fail": 0,
                  "required": "—", "by_group": [], "file": "—"}
    ids_data = read_ids(ids_path) if ids_path else _ids_empty
    lod_data = read_lod(lod_path) if lod_path else _lod_empty
    qto_data = read_qto(qto_path) if qto_path else []

    # ── Xuất Excel ──
    wb = openpyxl.Workbook()
    _sheet_tong_hop(wb, ids_data, lod_data, qto_data,
                    args.project, args.discipline, args.stage,
                    ids_threshold=args.ids_threshold)
    _sheet_loi_uu_tien(wb, ids_data, lod_data, args.top,
                       args.project, args.discipline, args.stage)

    out_path = OUTPUT_DIR / (
        f"AuditSummary_{slug}_{args.discipline}_{args.stage}"
        f"_{date.today().strftime('%Y%m%d')}.xlsx"
    )
    wb.save(str(out_path))

    # ── Summary ──
    ids_pct = ids_data.get("pct")
    lod_pct = lod_data.get("pct")
    ids_ok  = ids_data["total"] > 0 and (ids_pct or 0) >= args.ids_threshold and ids_data.get("errors", 0) == 0
    lod_ok  = lod_data["total"] > 0 and lod_data.get("fail", 0) == 0
    all_ok  = ids_ok and lod_ok

    print(f"\n  IDS : {ids_pct}% compliance, {ids_data.get('errors',0)} errors, "
          f"{ids_data.get('warnings',0)} warnings")
    # % compliance chỉ nói về phần ĐÃ kiểm. Hai số dưới nói về phần CHƯA kiểm —
    # không tính vào ngưỡng đạt/không đạt, chỉ cảnh báo để không đọc nhầm 100%.
    _kk, _kp = ids_data.get("khong_kiem") or 0, ids_data.get("khong_phu")
    if _kk or _kp:
        print(f"        ⚠ CHƯA KIỂM HẾT: {_kk} yêu cầu không có cấu kiện nào để kiểm"
              + (f", {_kp} cấu kiện không yêu cầu nào phủ" if _kp else "")
              + "  → xem sheet DoPhu_CanhBao")
    print(f"  LOD : {lod_pct}% đạt LOD {lod_data.get('required','—')}, "
          f"{lod_data.get('fail',0)} không đạt")
    print(f"  QTO : {len(qto_data)} nhóm khối lượng")
    print(f"\n  {'✅ ĐẠT — mô hình sẵn sàng nộp hồ sơ' if all_ok else '❌ CHƯA ĐẠT — xem sheet Loi_Uu_Tien'}")
    print(f"  Output: {out_path}")
    print(f"  Sheets: Tong_Hop | Loi_Uu_Tien")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
