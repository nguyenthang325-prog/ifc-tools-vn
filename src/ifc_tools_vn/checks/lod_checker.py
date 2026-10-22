"""
Kiểm tra LOD compliance của file IFC theo stage dự án (QD347/AIA G202).

LOD Matrix áp dụng:
  LOD 100 — Element tồn tại (placeholder, không có geometry)
  LOD 200 — Có geometry (Representation không None)
  LOD 300 — LOD 200 + Name + Material           ← TKKT yêu cầu
  LOD 350 — LOD 300 + FireRating trong Psets    ← TKTC / KT+MEP
  LOD 400 — LOD 350 + LoadBearing trong Psets   ← TKTC / KC
            (KC bỏ qua mốc 350: cấu kiện kết cấu không mang FireRating,
             xét thẳng LoadBearing từ LOD 300 → 400)
  LOD 500 — LOD 400 + SerialNumber/Manufacturer ← AS-BUILT

Cách dùng:
    python lod_checker.py --ifc Duplex_A_20110907.ifc --stage TKKT --discipline KT
"""
import sys
import re
import argparse
from pathlib import Path
from datetime import date
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

# Giá trị Pset không hợp lệ — placeholder Revit thường điền tự động
_INVALID_STR = frozenset({
    "", "none", "undefined", "unknown", "—", "-", "--", "n/a", "na", "?", "tbd",
})


def _is_valid_pset_value(val) -> bool:
    """True nếu val là dữ liệu có nghĩa (không phải None, rỗng, hay ký hiệu chỗ trống)."""
    if val is None:
        return False
    if isinstance(val, bool):
        return True  # True hoặc False đều là Boolean hợp lệ
    return str(val).strip().lower() not in _INVALID_STR


try:
    import ifcopenshell
    import ifcopenshell.util.element as ifc_util
except ImportError:
    print("Lỗi: pip install ifcopenshell")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Lỗi: pip install openpyxl")
    sys.exit(1)

from ifc_tools_vn.config import BASE, KET_QUA_DIR, OUTPUT_DIR  # đường dẫn: xem config.py

# ── LOD Matrix ────────────────────────────────────────────────────────────────

STAGE_LOD = {
    "TKCS":     200,
    "TKKT":     300,
    "TKTC":     350,
    "AS-BUILT": 500,
}
# KC ở TKTC cần LOD 400 (thông tin LoadBearing)
STAGE_LOD_KC = {**STAGE_LOD, "TKTC": 400}

# Loại entity kiểm tra theo discipline
# Loại THUẦN kết cấu — không bao giờ mang FireRating dù nằm trong mô hình bộ môn nào.
# Cố tình KHÔNG gồm IfcSlab/IfcWall: hai loại này thuộc cả KT lẫn KC, và ở mô hình
# kiến trúc chúng là sàn/tường ngăn nên vẫn phải khai FireRating.
_STRUCTURAL_ONLY_TYPES = {
    "IfcBeam", "IfcColumn", "IfcPlate", "IfcMember", "IfcFooting", "IfcPile",
}

DISC_TYPES = {
    "KT": [
        "IfcWall", "IfcWallStandardCase", "IfcSlab", "IfcRoof",
        "IfcDoor", "IfcWindow", "IfcStair", "IfcRailing", "IfcRamp",
        "IfcCovering", "IfcFurnishingElement",
    ],
    "KC": [
        "IfcBeam", "IfcColumn", "IfcSlab",
        "IfcPlate", "IfcMember",
        "IfcWall", "IfcWallStandardCase",
    ],
    "MEP": [
        "IfcPipeSegment", "IfcPipeFitting",
        "IfcDuctSegment", "IfcDuctFitting",
        "IfcFlowSegment", "IfcFlowTerminal",
        "IfcFlowController", "IfcFlowFitting",
        "IfcLightFixture", "IfcAirTerminal",
    ],
}
ALL_TYPES = list({t for types in DISC_TYPES.values() for t in types})

TYPE_GROUP = {
    "IfcWall": "Tường", "IfcWallStandardCase": "Tường",
    "IfcSlab": "Sàn/Mái", "IfcRoof": "Sàn/Mái",
    "IfcBeam": "Dầm", "IfcColumn": "Cột",
    "IfcDoor": "Cửa", "IfcWindow": "Cửa sổ",
    "IfcStair": "Cầu thang", "IfcRailing": "Lan can", "IfcRamp": "Dốc",
    "IfcPlate": "Tấm KC", "IfcMember": "Thanh KC",
    "IfcPipeSegment": "Ống nước", "IfcPipeFitting": "Ống nước",
    "IfcDuctSegment": "Ống gió", "IfcDuctFitting": "Ống gió",
    "IfcFlowSegment": "MEP", "IfcFlowTerminal": "MEP",
    "IfcFlowController": "MEP", "IfcFlowFitting": "MEP",
    "IfcLightFixture": "Đèn chiếu sáng", "IfcAirTerminal": "Miệng gió",
    "IfcCovering": "Ốp lát", "IfcFurnishingElement": "Nội thất",
}

# ── Excel styles ──────────────────────────────────────────────────────────────
_H_FILL    = PatternFill("solid", fgColor="1F4E79")
_S_FILL    = PatternFill("solid", fgColor="2E75B6")
_OK_FILL   = PatternFill("solid", fgColor="E2EFDA")
_FAIL_FILL = PatternFill("solid", fgColor="FCE4D6")
_WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
_thin      = Side(style="thin", color="CCCCCC")
_border    = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_center    = Alignment(horizontal="center", vertical="center", wrap_text=True)
_left      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_H_FONT    = Font(bold=True, color="FFFFFF", size=10)
_B_FONT    = Font(bold=True, size=10)
_N_FONT    = Font(size=10)


def _c(ws, row, col, value=None, fill=None, font=None, align=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.border    = _border
    cell.fill      = fill  or PatternFill()
    cell.font      = font  or _N_FONT
    cell.alignment = align or _center
    return cell


def _col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── IFC helpers ───────────────────────────────────────────────────────────────

def get_storey(element) -> str:
    for rel in getattr(element, "ContainedInStructure", None) or []:
        container = rel.RelatingStructure
        if container.is_a("IfcBuildingStorey"):
            return container.Name or ""
    return ""


def get_achieved_lod(element, discipline: str = "ALL") -> tuple:
    """Trả về (lod_đạt_được: int, lý_do_chưa_đạt_cao_hơn: list[str])."""
    missing = []

    # LOD 200: cần geometry
    has_geom = element.Representation is not None
    if not has_geom:
        missing.append("Thiếu geometry (Representation = None)")
        return 100, missing
    lod = 200

    # LOD 300: cần Name + Material
    name = (element.Name or "").strip()
    has_name = bool(name) and name.lower() not in ("unnamed", "")
    has_material = any(
        assoc.is_a("IfcRelAssociatesMaterial")
        for assoc in (getattr(element, "HasAssociations", None) or [])
    )
    if not has_name:
        missing.append("Thiếu Name (Name rỗng hoặc 'Unnamed')")
    if not has_material:
        missing.append("Thiếu vật liệu (chưa gán IfcRelAssociatesMaterial)")
    if has_name and has_material:
        lod = 300
    if lod < 300:
        return lod, missing

    # Đọc toàn bộ Psets 1 lần (dùng cho các mốc LOD 350/400/500)
    try:
        psets = ifc_util.get_psets(element)
        flat_kv = {}  # {prop_name_lower: value}
        for pset in psets.values():
            if isinstance(pset, dict):
                for k, v in pset.items():
                    flat_kv[k.lower()] = v
    except Exception:
        flat_kv = {}

    # LOD 350: FireRating — CHỈ áp dụng cho KT/MEP. Cấu kiện kết cấu (dầm/cột/tường KC)
    # không mang FireRating (thuộc tính PCCC/kiến trúc); với KC bỏ qua mốc 350 và xét
    # thẳng LoadBearing để lên LOD 400 từ mốc 300.
    # Miễn trừ theo LOẠI CẤU KIỆN, không chỉ theo cờ --discipline: mô hình kiến trúc luôn
    # chứa dầm/cột, và trước đây chúng bị đòi FireRating nên bị chấm LOD 0% oan
    # (một mô hình KT thử: 337/396 lỗi là dầm+cột). IfcSlab/IfcWall KHÔNG miễn trừ — trong mô hình
    # kiến trúc chúng là sàn/tường ngăn và vẫn cần FireRating.
    if discipline.upper() != "KC" and element.is_a() not in _STRUCTURAL_ONLY_TYPES:
        fire_key  = next((k for k in flat_kv if "firerating" in k or "fire_rating" in k), None)
        has_fire  = fire_key is not None and _is_valid_pset_value(flat_kv[fire_key])
        if has_fire:
            lod = 350
        else:
            fire_val = flat_kv.get(fire_key) if fire_key else None
            if fire_key and not _is_valid_pset_value(fire_val):
                missing.append(f"FireRating = '{fire_val}' (giá trị không hợp lệ — cần điền EI60, REI90… theo TCVN 9385)")
            else:
                missing.append("Thiếu FireRating trong Psets (để đạt LOD 350)")
            return lod, missing

    # LOD 400: cần LoadBearing có giá trị Boolean hợp lệ trong Psets
    load_key  = next((k for k in flat_kv if "loadbearing" in k or "load_bearing" in k), None)
    has_load  = load_key is not None and _is_valid_pset_value(flat_kv[load_key])
    if has_load:
        lod = 400
    else:
        load_val = flat_kv.get(load_key) if load_key else None
        if load_key and not _is_valid_pset_value(load_val):
            missing.append(f"LoadBearing = '{load_val}' (giá trị không hợp lệ — cần True hoặc False)")
        else:
            missing.append("Thiếu LoadBearing trong Psets (để đạt LOD 400)")
        return lod, missing

    # LOD 500: cần SerialNumber hoặc Manufacturer có giá trị hợp lệ
    serial_key = next((k for k in flat_kv if "serialnumber" in k or "manufacturer" in k
                       or "modelnumber" in k), None)
    has_serial = serial_key is not None and _is_valid_pset_value(flat_kv[serial_key])
    if has_serial:
        lod = 500
    else:
        serial_val = flat_kv.get(serial_key) if serial_key else None
        if serial_key and not _is_valid_pset_value(serial_val):
            missing.append(f"SerialNumber/Manufacturer = '{serial_val}' (giá trị không hợp lệ)")
        else:
            missing.append("Thiếu SerialNumber/Manufacturer trong Psets (để đạt LOD 500)")

    return lod, missing


def fix_hint(missing_reasons: list) -> str:
    hints = []
    for r in missing_reasons:
        if "Representation" in r:
            hints.append("Kiểm tra visibility/workset; đảm bảo element có geometry trong Revit")
        elif "Name" in r:
            hints.append("Properties → Identity Data → Type Name hoặc Mark")
        elif "vật liệu" in r.lower():
            hints.append("Gán Material trong Element/Type Properties")
        elif "FireRating" in r:
            hints.append("Thêm Pset_WallCommon → FireRating (hoặc Pset_DoorCommon)")
        elif "LoadBearing" in r:
            hints.append("Thêm Pset_WallCommon → LoadBearing = True")
        elif "SerialNumber" in r:
            hints.append("Thêm Pset_ManufacturerTypeInformation → Manufacturer + SerialNumber")
    return " | ".join(hints) if hints else ""


# ── Audit ─────────────────────────────────────────────────────────────────────

def audit_ifc(ifc_path: Path, discipline: str, stage: str) -> list:
    ifc = ifcopenshell.open(str(ifc_path))
    target_types = DISC_TYPES.get(discipline.upper(), ALL_TYPES)
    lod_map      = STAGE_LOD_KC if discipline.upper() == "KC" else STAGE_LOD
    required     = lod_map.get(stage.upper(), 300)

    records = []
    seen_guid: set = set()  # tránh đếm trùng khi 1 entity khớp cả type cha lẫn subtype
                            # (vd IfcWallStandardCase vừa khớp "IfcWall" vừa khớp chính nó)
    for ifc_type in target_types:
        try:
            elements = ifc.by_type(ifc_type)
        except Exception:
            continue
        for el in elements:
            guid = el.GlobalId or ""
            if guid:
                if guid in seen_guid:
                    continue
                seen_guid.add(guid)
            achieved, missing = get_achieved_lod(el, discipline)
            records.append({
                "name":     (el.Name or "").strip() or "—",
                "type":     ifc_type,
                "group":    TYPE_GROUP.get(ifc_type, ifc_type),
                "storey":   get_storey(el),
                "required": required,
                "achieved": achieved,
                "passed":   achieved >= required,
                "missing":  " | ".join(missing),
                "hint":     fix_hint(missing) if achieved < required else "",
            })
    return records


# ── Excel export ──────────────────────────────────────────────────────────────

def export_excel(records: list, ifc_path: Path, discipline: str, stage: str,
                 project: str, out_path: Path):
    wb    = openpyxl.Workbook()
    total = len(records)
    n_ok  = sum(1 for r in records if r["passed"])
    pct   = round(n_ok / total * 100, 1) if total else 0
    req   = records[0]["required"] if records else "—"

    # ── Sheet 1: LOD Audit ──
    ws = wb.active
    ws.title = "LOD_Audit"
    _col_widths(ws, [6, 34, 22, 14, 12, 10, 10, 8, 48])

    ws.merge_cells("A1:I1")
    c = ws.cell(1, 1, f"LOD AUDIT — {project.upper()} | {discipline} | Stage: {stage} | LOD yêu cầu: {req}")
    c.fill = _H_FILL; c.font = Font(bold=True, color="FFFFFF", size=12)
    c.alignment = _center; c.border = _border; ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:I2")
    c2 = ws.cell(2, 1, f"File: {ifc_path.name}  |  {date.today().strftime('%d/%m/%Y')}  |  Compliance: {n_ok}/{total} ({pct}%)")
    c2.font = Font(italic=True, size=9); c2.alignment = _center
    ws.row_dimensions[2].height = 16

    for j, h in enumerate(["STT", "Tên element", "Loại entity", "Nhóm", "Tầng", "LOD YC", "LOD Đạt", "Pass?", "Lý do chưa đạt"], 1):
        _c(ws, 3, j, h, fill=_S_FILL, font=_H_FONT)
    ws.row_dimensions[3].height = 22

    for i, r in enumerate(records, 1):
        fill = _OK_FILL if r["passed"] else _FAIL_FILL
        _c(ws, i+3, 1, i,              font=_N_FONT)
        _c(ws, i+3, 2, r["name"],      font=_N_FONT,      align=_left, fill=fill)
        _c(ws, i+3, 3, r["type"],      font=Font(size=9), align=_left)
        _c(ws, i+3, 4, r["group"],     font=_N_FONT)
        _c(ws, i+3, 5, r["storey"],    font=_N_FONT)
        _c(ws, i+3, 6, r["required"],  font=_N_FONT)
        _c(ws, i+3, 7, r["achieved"],  font=_N_FONT)
        _c(ws, i+3, 8, "✅" if r["passed"] else "❌", font=_N_FONT)
        _c(ws, i+3, 9, r["missing"],   font=Font(size=9), align=_left,
           fill=_FAIL_FILL if r["missing"] else PatternFill())
    ws.freeze_panes = "A4"

    # ── Sheet 2: Lỗi LOD ──
    failures = [r for r in records if not r["passed"]]
    ws2 = wb.create_sheet("Loi_LOD")
    _col_widths(ws2, [6, 34, 22, 14, 12, 10, 10, 48, 55])
    ws2.merge_cells("A1:I1")
    c3 = ws2.cell(1, 1, f"ELEMENT KHÔNG ĐẠT LOD ({len(failures)}/{total} elements — {100 - pct}% fail)")
    c3.fill = PatternFill("solid", fgColor="C00000")
    c3.font = Font(bold=True, color="FFFFFF", size=12)
    c3.alignment = _center; c3.border = _border; ws2.row_dimensions[1].height = 28

    for j, h in enumerate(["STT", "Tên element", "Loại entity", "Nhóm", "Tầng", "LOD YC", "LOD Đạt", "Lý do thiếu", "Gợi ý sửa trong Revit"], 1):
        _c(ws2, 2, j, h, fill=_S_FILL, font=_H_FONT)
    for i, r in enumerate(failures, 1):
        _c(ws2, i+2, 1, i,             font=_N_FONT)
        _c(ws2, i+2, 2, r["name"],     font=_N_FONT,      align=_left, fill=_FAIL_FILL)
        _c(ws2, i+2, 3, r["type"],     font=Font(size=9), align=_left)
        _c(ws2, i+2, 4, r["group"],    font=_N_FONT)
        _c(ws2, i+2, 5, r["storey"],   font=_N_FONT)
        _c(ws2, i+2, 6, r["required"], font=_N_FONT)
        _c(ws2, i+2, 7, r["achieved"], font=_N_FONT)
        _c(ws2, i+2, 8, r["missing"],  font=Font(size=9), align=_left, fill=_WARN_FILL)
        _c(ws2, i+2, 9, r["hint"],     font=Font(size=9), align=_left)

    # ── Sheet 3: Thống kê ──
    ws3 = wb.create_sheet("Thong_Ke")
    _col_widths(ws3, [28, 10, 10, 10, 14])
    ws3.merge_cells("A1:E1")
    c4 = ws3.cell(1, 1, f"THỐNG KÊ LOD — {project.upper()} | {discipline} | {stage}")
    c4.fill = _H_FILL; c4.font = Font(bold=True, color="FFFFFF", size=13)
    c4.alignment = _center; c4.border = _border; ws3.row_dimensions[1].height = 30

    for j, h in enumerate(["Nhóm element", "Tổng", "Đạt", "Không đạt", "% Compliance"], 1):
        _c(ws3, 3, j, h, fill=_S_FILL, font=_H_FONT)

    by_group = defaultdict(lambda: {"total": 0, "pass": 0})
    for r in records:
        by_group[r["group"]]["total"] += 1
        if r["passed"]:
            by_group[r["group"]]["pass"] += 1

    for i, (grp, cnt) in enumerate(sorted(by_group.items()), 4):
        t = cnt["total"]; p = cnt["pass"]; f = t - p
        pg = round(p / t * 100, 1) if t else 0
        row_fill = _OK_FILL if pg == 100 else (_WARN_FILL if pg >= 80 else _FAIL_FILL)
        _c(ws3, i, 1, grp,         font=_B_FONT, align=_left)
        _c(ws3, i, 2, t,           font=_N_FONT)
        _c(ws3, i, 3, p,           font=_N_FONT, fill=_OK_FILL   if p > 0 else PatternFill())
        _c(ws3, i, 4, f,           font=_N_FONT, fill=_FAIL_FILL if f > 0 else PatternFill())
        _c(ws3, i, 5, f"{pg}%",    font=_N_FONT, fill=row_fill)

    r_tot = len(by_group) + 4
    _c(ws3, r_tot, 1, "TỔNG",          fill=_H_FILL, font=_H_FONT, align=_left)
    _c(ws3, r_tot, 2, total,            fill=_H_FILL, font=_H_FONT)
    _c(ws3, r_tot, 3, n_ok,             fill=_H_FILL, font=_H_FONT)
    _c(ws3, r_tot, 4, total - n_ok,     fill=_H_FILL, font=_H_FONT)
    _c(ws3, r_tot, 5, f"{pct}%",        fill=_H_FILL, font=_H_FONT)

    wb.save(str(out_path))
    print(f"  ✅ Excel: {out_path.name}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Kiểm tra LOD compliance IFC theo stage (QD347)")
    parser.add_argument("--ifc",        required=True,
                        help="File IFC cần kiểm tra (path tuyệt đối hoặc tương đối từ thư mục hiện tại)")
    parser.add_argument("--stage",      required=True,
                        choices=["TKCS", "TKKT", "TKTC", "AS-BUILT"],
                        help="Giai đoạn: TKCS / TKKT / TKTC / AS-BUILT")
    parser.add_argument("--discipline", default="ALL",
                        choices=["KT", "KC", "MEP", "ALL"],
                        help="Discipline: KT / KC / MEP / ALL (mặc định: ALL)")
    parser.add_argument("--project",    default="",
                        help="Tên dự án (mặc định: lấy từ tên file IFC)")
    args = parser.parse_args()

    ifc_path = Path(args.ifc)
    if not ifc_path.is_absolute():
        ifc_path = BASE / args.ifc
    if not ifc_path.exists():
        print(f"Lỗi: không tìm thấy file IFC: {ifc_path}")
        sys.exit(1)

    project = args.project or ifc_path.stem
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    lod_map  = STAGE_LOD_KC if args.discipline == "KC" else STAGE_LOD
    required = lod_map.get(args.stage, 300)
    size_mb  = ifc_path.stat().st_size // 1_048_576

    print(f"\n{'='*60}")
    print(f"  LOD CHECKER — {project}")
    print(f"  File        : {ifc_path.name}  ({size_mb} MB)")
    print(f"  Stage       : {args.stage}  |  Discipline: {args.discipline}")
    print(f"  LOD yêu cầu: {required}")
    print(f"{'='*60}")

    print("  Đang mở và phân tích IFC...")
    records = audit_ifc(ifc_path, args.discipline, args.stage)

    if not records:
        print("  Không tìm thấy element nào phù hợp với discipline đã chọn.")
        print("  Gợi ý: thử --discipline ALL để kiểm tra toàn bộ IFC.\n")
        sys.exit(0)

    total = len(records)
    n_ok  = sum(1 for r in records if r["passed"])
    pct   = round(n_ok / total * 100, 1)

    print(f"\n  Kết quả:")
    print(f"    Tổng elements : {total}")
    print(f"    ✅ Đạt LOD {required} : {n_ok} ({pct}%)")
    print(f"    ❌ Không đạt  : {total - n_ok}")

    slug     = re.sub(r"[^\w]", "_", project)
    out_path = OUTPUT_DIR / f"LOD_Audit_{slug}_{args.stage}_{date.today().strftime('%Y%m%d')}.xlsx"

    print("  Đang xuất báo cáo...")
    export_excel(records, ifc_path, args.discipline, args.stage, project, out_path)
    print(f"\n  ✅ Hoàn thành! Kết quả tại: {out_path}\n")


if __name__ == "__main__":
    main()
