"""
BIM Audit Pipeline — Chạy toàn bộ chuỗi kiểm tra IFC trong 1 lệnh.
Tự động gọi: IDS (eir_to_ids -> ids_run, chuẩn buildingSMART) → lod_checker → ifc_qto
→ tổng hợp Health Dashboard.

Output: ifc_tools_output/<ngày>/BIM_HealthDashboard_{project}_{stage}_{YYYYMMDD}.xlsx
  Sheet "Kiem_Tra_Tong_Quan" — Pass/Fail từng kiểm tra, metric chính, link file output
  Sheet "Hanh_Dong"          — Danh sách ưu tiên việc cần làm (ERROR trước, WARNING sau)

Cách dùng:
    python bim_audit_pipeline.py --ifc Duplex_A_20110907.ifc --project Duplex --stage TKKT --discipline KT
    python bim_audit_pipeline.py --ifc file.ifc --project "TenDuAn" --stage TKTC --discipline KC
"""
import sys
import re
import argparse
import subprocess
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

from ifc_tools_vn.config import BASE, KET_QUA_DIR  # đường dẫn: xem config.py
SCRIPT_DIR  = Path(__file__).parent                        # eir_to_ids, ids_run, lod_checker, ifc_qto

# ── Excel styles ──────────────────────────────────────────────────────────────
_H_FILL  = PatternFill("solid", fgColor="1F4E79")
_S_FILL  = PatternFill("solid", fgColor="2E75B6")
_OK_FILL = PatternFill("solid", fgColor="E2EFDA")
_WA_FILL = PatternFill("solid", fgColor="FFF2CC")
_ER_FILL = PatternFill("solid", fgColor="FCE4D6")
_thin    = Side(style="thin", color="CCCCCC")
_border  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_center  = Alignment(horizontal="center", vertical="center", wrap_text=True)
_left    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_H_FONT  = Font(bold=True, color="FFFFFF", size=10)
_B_FONT  = Font(bold=True, size=10)
_N_FONT  = Font(size=10)


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


# ── Run subprocess ────────────────────────────────────────────────────────────
def _run_script(script_name: str, extra_args: list):
    """Chạy script con, trả về (returncode, stdout, stderr)."""
    script_path = SCRIPT_DIR / script_name
    if not script_path.exists():
        return -1, "", f"Không tìm thấy script: {script_path}"
    cmd = [sys.executable, str(script_path)] + extra_args
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return result.returncode, result.stdout, result.stderr


# ── Parse metrics từ stdout ──────────────────────────────────────────────────
def _parse_ids_chuan(stdout: str) -> dict:
    """Parser cho ids_run.py (bộ chuẩn ifctester, buildingSMART IDS).

    Không có tầng WARNING nên warnings luôn 0. "errors" lấy từ số lượt kiểm (check) trượt.

    CẢNH BÁO (đã kiểm chứng 29/09/2026): spec IDS ở đây KHÔNG phải 'required' — chúng mang
    minOccurs=0 (mặc định ifctester) = 'optional', nên spec không tìm thấy cấu kiện nào vẫn
    được tính ĐẠT. pct 100% KHÔNG có nghĩa đã kiểm hết; phần "ĐÃ KIỂM ĐƯỢC NHỮNG GÌ" trong
    stdout của ids_run.py và sheet DoPhu_CanhBao mới cho biết phạm vi thật.
    """
    m = re.search(r"Compliance check\s*:\s*(\d+\.?\d*)%\s*\((\d+)/(\d+)\s*lượt kiểm\)", stdout)
    if not m:
        return {"pct": None, "errors": None, "warnings": None}
    pct, passed, total = float(m.group(1)), int(m.group(2)), int(m.group(3))
    return {"pct": pct, "errors": total - passed, "warnings": 0}


def _parse_ids_pham_vi(stdout: str) -> dict:
    """Rút PHẠM VI kiểm thật từ mục 'ĐÃ KIỂM ĐƯỢC NHỮNG GÌ' của ids_run.py.

    Vì sao cần: spec IDS mang minOccurs=0 nên spec không tìm thấy cấu kiện nào vẫn được
    tính ĐẠT → pct 100% có thể là 100% của phần chưa kiểm gì. Xem ids_dophu.py.
    """
    m_kk = re.search(r"còn\s+(\d+)\s+yêu cầu KHÔNG có cấu kiện", stdout)
    m_kp = re.search(r"KHÔNG YÊU CẦU NÀO PHỦ\s*—\s*(\d+)\s+cấu kiện", stdout)
    m_dp = re.search(r"Độ phủ cấu kiện\s*:\s*(\d+/\d+\s*\(\d+%\))", stdout)
    return {
        "khong_kiem": int(m_kk.group(1)) if m_kk else 0,
        "khong_phu":  int(m_kp.group(1)) if m_kp else 0,
        "do_phu":     m_dp.group(1).strip() if m_dp else None,
    }


def _parse_lod(stdout: str) -> dict:
    m_pct  = re.search(r"Đạt LOD \d+\s*:\s*\d+\s*\((\d+\.?\d*)%\)", stdout)
    m_fail = re.search(r"Không đạt\s*:\s*(\d+)", stdout)
    m_tot  = re.search(r"Tổng elements\s*:\s*(\d+)", stdout)
    return {
        "pct":   float(m_pct.group(1))  if m_pct  else None,
        "fail":  int(m_fail.group(1))   if m_fail else None,
        "total": int(m_tot.group(1))    if m_tot  else None,
    }


def _parse_qto(stdout: str) -> dict:
    m_tot   = re.search(r"Tổng elements\s*:\s*(\d+)", stdout)
    m_noqto = re.search(r"Không có Qto\s*:\s*(\d+)", stdout)
    return {
        "total": int(m_tot.group(1))   if m_tot   else None,
        "noqto": int(m_noqto.group(1)) if m_noqto else None,
    }


def _latest_file(prefix: str) -> str:
    """Tìm file output mới nhất trong thư mục kết quả theo prefix."""
    files = sorted(KET_QUA_DIR.rglob(f"{prefix}*.xlsx"),
                   key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0].name if files else "—"


# ── Build Excel dashboard ─────────────────────────────────────────────────────
def _build_dashboard(checks: list, project: str, stage: str, discipline: str, out_path: Path):
    wb = openpyxl.Workbook()

    # ── Sheet 1: Tổng quan ───────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Kiem_Tra_Tong_Quan"
    _col_widths(ws, [4, 30, 16, 28, 32, 42])
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A1:F1")
    c = ws.cell(1, 1, (
        f"BIM HEALTH DASHBOARD — {project.upper()}"
        f"  |  Stage: {stage}  |  Discipline: {discipline}"
        f"  |  {date.today().strftime('%d/%m/%Y')}"
    ))
    c.fill = _H_FILL; c.font = Font(bold=True, color="FFFFFF", size=12)
    c.alignment = _center; c.border = _border

    for j, h in enumerate(["#", "Kiểm Tra", "Trạng Thái", "Metric Chính", "Ghi Chú", "File Output"], 1):
        _c(ws, 2, j, h, fill=_S_FILL, font=_H_FONT)
    ws.row_dimensions[2].height = 24

    all_ok = True
    for i, chk in enumerate(checks, 3):
        ws.row_dimensions[i].height = 32
        ok     = chk["ok"]
        is_err = chk.get("script_error", False)
        fill   = _OK_FILL if ok else (_ER_FILL if is_err else _WA_FILL)
        status = "✅ ĐẠT" if ok else ("❌ LỖI SCRIPT" if is_err else "⚠ CẦN XỬ LÝ")
        if not ok:
            all_ok = False
        _c(ws, i, 1, i - 2,              font=_N_FONT)
        _c(ws, i, 2, chk["name"],        font=_B_FONT, align=_left)
        _c(ws, i, 3, status,             fill=fill,    font=_B_FONT)
        _c(ws, i, 4, chk["metric"],      font=_N_FONT, align=_left)
        _c(ws, i, 5, chk["note"],        font=_N_FONT, align=_left)
        _c(ws, i, 6, chk["output_file"], font=Font(size=9, italic=True), align=_left)

    r_sum = len(checks) + 3
    ws.row_dimensions[r_sum].height = 30
    ws.merge_cells(f"A{r_sum}:B{r_sum}")
    _c(ws, r_sum, 1, "KẾT LUẬN TỔNG THỂ", fill=_H_FILL, font=_H_FONT, align=_left)
    verdict_fill = _OK_FILL if all_ok else _ER_FILL
    verdict_text = (
        "✅ ĐẠT — Mô hình sẵn sàng nộp hồ sơ" if all_ok
        else "❌ CHƯA ĐẠT — Xem sheet Hanh_Dong để biết việc cần làm"
    )
    ws.merge_cells(f"C{r_sum}:F{r_sum}")
    _c(ws, r_sum, 3, verdict_text, fill=verdict_fill,
       font=Font(bold=True, size=11), align=_left)

    # ── Sheet 2: Hành động ───────────────────────────────────────────────────
    ws2 = wb.create_sheet("Hanh_Dong")
    _col_widths(ws2, [4, 16, 36, 52, 22])
    ws2.merge_cells("A1:E1")
    c2 = ws2.cell(1, 1, f"DANH SÁCH HÀNH ĐỘNG — {project.upper()}")
    c2.fill = _H_FILL; c2.font = Font(bold=True, color="FFFFFF", size=12)
    c2.alignment = _center; c2.border = _border
    for j, h in enumerate(["#", "Mức Độ", "Hạng Mục", "Hành Động Cần Thực Hiện", "Tool"], 1):
        _c(ws2, 2, j, h, fill=_S_FILL, font=_H_FONT)

    actions = [act for chk in checks for act in chk.get("actions", [])]

    if not actions:
        ws2.merge_cells("A3:E3")
        _c(ws2, 3, 1,
           "✅ Không có hành động cần thực hiện — tất cả kiểm tra đạt.",
           fill=_OK_FILL, font=_B_FONT, align=_left)
    else:
        for i, act in enumerate(actions, 3):
            sev_fill = _ER_FILL if act["severity"] == "ERROR" else _WA_FILL
            _c(ws2, i, 1, i - 2,          font=_N_FONT)
            _c(ws2, i, 2, act["severity"], fill=sev_fill, font=_B_FONT)
            _c(ws2, i, 3, act["item"],     font=_N_FONT, align=_left)
            _c(ws2, i, 4, act["action"],   font=_N_FONT, align=_left)
            _c(ws2, i, 5, act["tool"],     font=Font(size=9, italic=True), align=_left)

    wb.save(str(out_path))
    print(f"  ✅ Dashboard: {out_path.name}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="BIM Audit Pipeline — chạy toàn bộ chuỗi kiểm tra IFC"
    )
    parser.add_argument("--ifc",        required=True,
                        help="File IFC (path tuyệt đối hoặc tương đối từ thư mục hiện tại)")
    parser.add_argument("--project",    default="",
                        help="Tên dự án (mặc định: lấy từ tên file IFC)")
    parser.add_argument("--stage",      default="TKKT",
                        choices=["TKCS", "TKKT", "TKTC", "AS-BUILT"],
                        help="Giai đoạn: TKCS / TKKT / TKTC / AS-BUILT")
    parser.add_argument("--discipline", default="ALL",
                        choices=["KT", "KC", "MEP", "ALL"],
                        help="Discipline: KT / KC / MEP / ALL")
    parser.add_argument("--output",     default=None,
                        help="Thư mục lưu kết quả (mặc định: ifc_tools_output/{ngày}/ hoặc biến IFC_TOOLS_OUTPUT)")
    args = parser.parse_args()

    ifc_path = Path(args.ifc)
    if not ifc_path.is_absolute():
        ifc_path = BASE / args.ifc
    if not ifc_path.exists():
        print(f"Lỗi: không tìm thấy file IFC: {ifc_path}")
        sys.exit(1)

    project = args.project or ifc_path.stem
    OUTPUT_DIR = (Path(args.output) if args.output
                  else KET_QUA_DIR / date.today().strftime('%Y-%m-%d'))
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    ifc_str = str(ifc_path)
    common  = [
        "--ifc",        ifc_str,
        "--discipline", args.discipline,
        "--stage",      args.stage,
        "--project",    project,
    ]

    print(f"\n{'='*60}")
    print(f"  BIM AUDIT PIPELINE — {project}")
    print(f"  File       : {ifc_path.name}")
    print(f"  Stage      : {args.stage}  |  Discipline: {args.discipline}")
    print(f"{'='*60}\n")

    checks = []

    # ── 1. IDS Check (bộ chuẩn buildingSMART: eir_to_ids -> ids_run) ──────────
    print("[1/3] IDS Check (chuẩn buildingSMART, EIR/QD347)...")
    ids_cache_dir = KET_QUA_DIR / "_ids_cache"
    slug_ids = re.sub(r"[^\w]", "_", project)
    ids_spec_path = ids_cache_dir / f"EIR_{slug_ids}_{args.discipline}_{args.stage}.ids"
    rc0, out0, err0 = _run_script(
        "eir_to_ids.py",
        ["--du-an", project, "--discipline", args.discipline, "--stage", args.stage,
         "-o", str(ids_spec_path)],
    )
    if rc0 != 0:
        rc1, out1, err1 = rc0, "", err0 or "Không sinh được file .ids từ EIR"
    else:
        rc1, out1, err1 = _run_script(
            "ids_run.py",
            ["--ids", str(ids_spec_path), "--ifc", ifc_str, "--project", project],
        )
    m1     = _parse_ids_chuan(out1)
    pct1   = m1["pct"]; err1_n = m1["errors"] or 0; warn1_n = m1["warnings"] or 0
    ids_ok = rc1 == 0 and pct1 is not None and err1_n == 0
    metric1 = f"Compliance: {pct1}%"     if pct1 is not None else "Không đọc được"
    note1   = (f"{err1_n} Errors (bộ chuẩn IDS không có tầng Warning)" if pct1 is not None
               else err1.strip()[:100] or "Script lỗi — kiểm tra ifcopenshell/ifctester đã cài chưa")
    # % Compliance chỉ nói về phần ĐÃ kiểm. Lấy thêm phạm vi thật từ stdout của ids_run.py
    # để pipeline không in "100%" trơ trọi trên một mô hình chưa kiểm hết.
    _pv = _parse_ids_pham_vi(out1)
    if _pv["khong_kiem"]:
        note1 += f" | ⚠ {_pv['khong_kiem']} yêu cầu KHÔNG kiểm được"
    if _pv["khong_phu"]:
        note1 += f" | ⚠ {_pv['khong_phu']} cấu kiện không yêu cầu nào phủ"
    if _pv["do_phu"]:
        note1 += f" | Độ phủ {_pv['do_phu']}"
    acts1 = []
    if _pv["khong_kiem"] or _pv["khong_phu"]:
        acts1.append({"severity": "WARNING", "item": "IDS — phạm vi kiểm chưa đủ",
                      "action": (f"{_pv['khong_kiem']} yêu cầu 0 cấu kiện áp dụng, "
                                 f"{_pv['khong_phu'] or 0} cấu kiện không spec nào phủ → "
                                 f"xem IDS_KetQua_*.xlsx sheet DoPhu_CanhBao"),
                      "tool": "ids_dophu.py"})
    if err1_n > 0:
        acts1.append({"severity": "ERROR", "item": "IDS — EIR/QD347 (chuẩn buildingSMART)",
                      "action": f"{err1_n} lượt kiểm trượt → xem IDS_KetQua_*.xlsx sheet Loi_Chi_Tiet",
                      "tool": "ids_run.py"})
    checks.append({"name": "IDS Check (EIR/QD347 — chuẩn buildingSMART)", "ok": ids_ok,
                   "script_error": rc1 not in (0, None) and pct1 is None,
                   "metric": metric1, "note": note1,
                   "output_file": _latest_file("IDS_KetQua_"), "actions": acts1})
    print(f"   → {metric1}  |  {note1}")

    # ── 2. LOD Check ─────────────────────────────────────────────────────────
    print("[2/3] LOD Check...")
    rc2, out2, err2 = _run_script("lod_checker.py", common)
    m2     = _parse_lod(out2)
    pct2   = m2["pct"]; fail2 = m2["fail"] or 0
    lod_ok = rc2 == 0 and pct2 is not None and fail2 == 0
    metric2 = f"Đạt LOD: {pct2}%"        if pct2 is not None else "Không đọc được"
    note2   = (f"{fail2} elements dưới LOD yêu cầu" if pct2 is not None
               else err2.strip()[:100] or "Script lỗi — kiểm tra ifcopenshell đã cài chưa")
    acts2 = []
    if fail2 > 0:
        acts2.append({"severity": "ERROR", "item": "LOD Check",
                      "action": f"{fail2} elements thiếu thông tin LOD → bổ sung Pset trong Revit, export IFC lại",
                      "tool": "lod_checker.py"})
    checks.append({"name": "LOD Check (QD347/AIA G202)", "ok": lod_ok,
                   "script_error": rc2 not in (0, None) and pct2 is None,
                   "metric": metric2, "note": note2,
                   "output_file": _latest_file("LOD_Audit_"), "actions": acts2})
    print(f"   → {metric2}  |  {note2}")

    # ── 3. IFC QTO ───────────────────────────────────────────────────────────
    print("[3/3] IFC QTO...")
    rc3, out3, err3 = _run_script("ifc_qto.py", common)
    m3      = _parse_qto(out3)
    total3  = m3["total"]; noqto3 = m3["noqto"] or 0
    qto_ok  = rc3 == 0 and total3 is not None and noqto3 == 0
    metric3 = f"{total3} elements"        if total3 is not None else "Không đọc được"
    note3   = ("Đủ dữ liệu KL" if total3 is not None and noqto3 == 0
               else (f"{noqto3} elements thiếu BaseQuantities" if total3 is not None
                     else err3.strip()[:100] or "Script lỗi — kiểm tra ifcopenshell đã cài chưa"))
    acts3 = []
    if noqto3 > 0 and total3:
        pct_miss = round(noqto3 / total3 * 100, 1)
        acts3.append({"severity": "WARNING", "item": "IFC QTO",
                      "action": (f"{noqto3}/{total3} ({pct_miss}%) elements thiếu BaseQuantities"
                                 " → Revit: File→Export→IFC→Modify Setup→check 'Export base quantities'"),
                      "tool": "ifc_qto.py"})
    checks.append({"name": "IFC QTO (BaseQuantities)", "ok": qto_ok,
                   "script_error": rc3 not in (0, None) and total3 is None,
                   "metric": metric3, "note": note3,
                   "output_file": _latest_file("IFC_QTO_"), "actions": acts3})
    print(f"   → {metric3}  |  {note3}")

    # ── Dashboard ─────────────────────────────────────────────────────────────
    slug     = re.sub(r"[^\w]", "_", project)
    out_path = OUTPUT_DIR / f"BIM_HealthDashboard_{slug}_{args.stage}_{date.today().strftime('%Y%m%d')}.xlsx"
    print("\n  Đang tổng hợp dashboard...")
    _build_dashboard(checks, project, args.stage, args.discipline, out_path)

    all_ok = all(c["ok"] for c in checks)
    n_fail = sum(1 for c in checks if not c["ok"])
    print(f"\n{'='*60}")
    if all_ok:
        print(f"  ✅ Mô hình ĐẠT — sẵn sàng nộp hồ sơ {args.stage}")
    else:
        print(f"  ⚠  {n_fail}/3 kiểm tra CHƯA ĐẠT — xem sheet Hanh_Dong")
    print(f"  Dashboard: {out_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
