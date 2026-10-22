"""
eir_to_ids.py — Sinh file IDS chuẩn buildingSMART từ bảng EIR (Yêu cầu trao đổi thông tin).
================================================================================
Thay cho 21 rule hard-code trong ids_checker.py: ở đây thước đo là FILE .ids CHUẨN
(buildingSMART Information Delivery Specification), chấm được bằng bất kỳ công cụ nào
tương thích IDS, và chính là đặc tả hợp đồng — không phải luật tự chế.

Nguồn yêu cầu: QĐ 347/QĐ-BXD (bảng LOD + thông tin theo giai đoạn) diễn đạt qua facet IDS.
Đã CỐ Ý loại các luật sai từng gây báo động giả:
  - KHÔNG đòi FireRating trên dầm/cột/sàn kết cấu (FireRating chỉ áp tường/cửa/sổ — KT).
  - Cốt thép / khối lượng KHÔNG thuộc IDS (đó là việc bóc QTO, không phải kiểm tra thông tin).

Cách dùng:
    # Sinh 1 file cho 1 bộ môn + giai đoạn
    python eir_to_ids.py --du-an "Nha pho mau" --discipline KC --stage TKTC -o out.ids
    # Sinh cả ma trận chuẩn (mọi bộ môn × mọi giai đoạn) vào 1 thư mục
    python eir_to_ids.py --du-an "Nha pho mau" --all -o EIR
"""
import argparse
import sys
from pathlib import Path
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")

try:
    from ifctester import ids
except ImportError:
    print("Lỗi: cần cài engine chuẩn -> python -m pip install ifctester")
    sys.exit(1)

try:
    from ifcopenshell import ifcopenshell_wrapper
except ImportError:
    ifcopenshell_wrapper = None

# Giai đoạn thiết kế VN, thứ tự tăng dần độ chi tiết
STAGE_ORDER = {"TKCS": 0, "TKKT": 1, "TKTC": 2, "AS-BUILT": 3}
IFC_VERSIONS = ["IFC2X3", "IFC4"]   # file Revit VN chủ yếu IFC2X3; thêm IFC4 cho tương lai


# ── Chống 2 lỗi im lặng của facet Entity (đã kiểm chứng 29/09/2026) ───────────
# (1) facet Entity khớp lớp CHÍNH XÁC, KHÔNG nhận lớp con (ifctester dùng
#     include_subtypes=False). Nên luật "IFCWALL" bỏ sót IfcWallStandardCase —
#     Revit xuất phần lớn tường ra lớp con này. Cùng bẫy với IfcBeamStandardCase,
#     IfcColumnStandardCase, IfcSlabStandardCase (IFC4).
#     => _mo_rong_lop() tự thêm MỌI lớp con có trong các schema đang khai, thay vì
#        liệt kê tay từng trường hợp rồi quên trường hợp sau.
# (2) Lớp có thể không tồn tại trong một schema (IFC2X3 KHÔNG có IfcPipeSegment).
#     ifctester tự lùi sang <Lớp>Type nên luật vẫn chạy được — chỉ khi CẢ HAI đều
#     thiếu thì luật mới chết. => _schema_dung_duoc() lọc đúng theo tiêu chí đó.
def _lop_con(ten_lop: str, schema: str) -> set:
    if ifcopenshell_wrapper is None:
        return set()
    try:
        S = ifcopenshell_wrapper.schema_by_name(schema)
        d = S.declaration_by_name(ten_lop)
    except Exception:
        return set()
    ra = set()
    try:
        for sub in d.subtypes():
            ten = sub.name()
            ra.add(ten.upper())
            ra |= _lop_con(ten, schema)      # đệ quy: lớp con của lớp con
    except Exception:
        pass
    return ra


def _mo_rong_lop(entity) -> list:
    """Danh sách lớp IFC cuối cùng cho applicability = lớp khai + mọi lớp con có thật."""
    goc = [entity] if isinstance(entity, str) else list(entity)
    ra = {str(e).strip().upper() for e in goc}
    for e in list(ra):
        for schema in IFC_VERSIONS:
            ra |= _lop_con(e, schema)
    return sorted(ra)


def _co_trong_schema(ten_lop: str, schema: str) -> bool:
    if ifcopenshell_wrapper is None:
        return True
    try:
        ifcopenshell_wrapper.schema_by_name(schema).declaration_by_name(ten_lop)
        return True
    except Exception:
        return False


def _schema_dung_duoc(lop_list: list) -> list:
    """Schema nào thật sự dùng được luật này (có lớp, hoặc có <Lớp>Type để ifctester lùi về)."""
    ra = []
    for schema in IFC_VERSIONS:
        if any(_co_trong_schema(l, schema) or _co_trong_schema(l + "Type", schema)
               for l in lop_list):
            ra.append(schema)
    return ra

# ── BẢNG EIR ──────────────────────────────────────────────────────────────────
# Mỗi mục = 1 Specification IDS: (mã, tiêu đề, bộ môn, giai đoạn tối thiểu,
#   applicability = entity IFC, requirement facet).
# facet: ("attribute", tên) | ("material",) | ("property", pset, prop, dataType)
#        | ("classification", hệ)
# Tất cả requirement là 'required' — vì đây là ĐẶC TẢ HỢP ĐỒNG, không có "warning".
# Nâng dần theo stage: cái gì chưa bắt buộc ở giai đoạn này thì đặt ở stage sau.

EIR = [
    # ── PRJ — thông tin dự án (mọi bộ môn) ──
    ("EIR-PRJ-01", "IfcProject phải có Tên dự án", "ALL", "TKCS",
     "IFCPROJECT", ("attribute", "Name")),
    ("EIR-PRJ-02", "IfcBuilding phải có Tên công trình", "ALL", "TKCS",
     "IFCBUILDING", ("attribute", "Name")),
    ("EIR-PRJ-03", "IfcSite phải có Tên địa điểm", "ALL", "TKKT",
     "IFCSITE", ("attribute", "Name")),
    ("EIR-PRJ-04", "Tầng (IfcBuildingStorey) phải có Tên", "ALL", "TKCS",
     "IFCBUILDINGSTOREY", ("attribute", "Name")),

    # ── KT — kiến trúc ──
    # Tường: Revit IFC2X3 xuất phần lớn thành IfcWallStandardCase (class con của IfcWall),
    # mà facet Entity của IDS khớp ĐÚNG class → phải liệt kê cả hai, nếu không bỏ sót ~98% tường.
    ("EIR-KT-01", "Tường phải có Tên", "KT", "TKKT",
     ["IFCWALL", "IFCWALLSTANDARDCASE"], ("attribute", "Name")),
    ("EIR-KT-02", "Tường phải gán Vật liệu", "KT", "TKKT",
     ["IFCWALL", "IFCWALLSTANDARDCASE"], ("material",)),
    ("EIR-KT-03", "Tường phải khai IsExternal (trong/ngoài)", "KT", "TKKT",
     ["IFCWALL", "IFCWALLSTANDARDCASE"], ("property", "Pset_WallCommon", "IsExternal", "IfcBoolean")),
    ("EIR-KT-04", "Tường phải khai FireRating (chống cháy)", "KT", "TKTC",
     ["IFCWALL", "IFCWALLSTANDARDCASE"], ("property", "Pset_WallCommon", "FireRating", None)),
    ("EIR-KT-05", "Cửa đi phải khai FireRating", "KT", "TKTC",
     "IFCDOOR", ("property", "Pset_DoorCommon", "FireRating", None)),
    ("EIR-KT-06", "Cửa sổ phải khai IsExternal", "KT", "TKKT",
     "IFCWINDOW", ("property", "Pset_WindowCommon", "IsExternal", "IfcBoolean")),

    # ── KC — kết cấu ── (KHÔNG có FireRating ở đây — bài học báo động giả)
    ("EIR-KC-01", "Dầm phải có Tên", "KC", "TKKT",
     "IFCBEAM", ("attribute", "Name")),
    ("EIR-KC-02", "Dầm phải gán Vật liệu", "KC", "TKKT",
     "IFCBEAM", ("material",)),
    ("EIR-KC-03", "Dầm phải khai LoadBearing (chịu lực)", "KC", "TKTC",
     "IFCBEAM", ("property", "Pset_BeamCommon", "LoadBearing", "IfcBoolean")),
    ("EIR-KC-04", "Cột phải gán Vật liệu", "KC", "TKKT",
     "IFCCOLUMN", ("material",)),
    ("EIR-KC-05", "Cột phải khai LoadBearing", "KC", "TKTC",
     "IFCCOLUMN", ("property", "Pset_ColumnCommon", "LoadBearing", "IfcBoolean")),
    ("EIR-KC-06", "Sàn phải gán Vật liệu", "KC", "TKKT",
     "IFCSLAB", ("material",)),
    ("EIR-KC-07", "Móng phải gán Vật liệu", "KC", "TKKT",
     "IFCFOOTING", ("material",)),

    # ── MEP — cơ điện ──
    ("EIR-MEP-01", "Ống (PipeSegment) phải có Tên", "MEP", "TKKT",
     "IFCPIPESEGMENT", ("attribute", "Name")),
    ("EIR-MEP-02", "Ống gió (DuctSegment) phải có Tên", "MEP", "TKKT",
     "IFCDUCTSEGMENT", ("attribute", "Name")),
]


def _make_facet(spec_facet):
    kind = spec_facet[0]
    if kind == "attribute":
        return ids.Attribute(name=spec_facet[1], cardinality="required")
    if kind == "material":
        return ids.Material(cardinality="required")
    if kind == "property":
        _, pset, prop, dtype = spec_facet
        kw = dict(propertySet=pset, baseName=prop, cardinality="required")
        if dtype:
            kw["dataType"] = dtype
        return ids.Property(**kw)
    if kind == "classification":
        return ids.Classification(system=spec_facet[1], cardinality="required")
    raise ValueError(f"facet lạ: {kind}")


def build_ids(du_an: str, discipline: str, stage: str, bat_buoc=None) -> ids.Ids:
    """bat_buoc: tập mã EIR phải TỒN TẠI trong mô hình (minOccurs=1).

    Mặc định rỗng = giữ nguyên minOccurs=0 của chuẩn IDS ('optional': có cấu kiện loại
    này thì phải đạt). Lưu ý đã kiểm chứng: với minOccurs=0, mô hình THIẾU HẲN loại cấu
    kiện đó vẫn được tính ĐẠT. Đặt mã vào bat_buoc để biến thành "bắt buộc phải có".
    Đây là quyết định NGỮ NGHĨA HỢP ĐỒNG — chỉ bật cho những mã mà giai đoạn/bộ môn đó
    chắc chắn phải có cấu kiện (đừng bật hàng loạt, sẽ sinh trượt oan).
    """
    disc = discipline.upper()
    stg = stage.upper()
    title = f"EIR {du_an} — {disc} — {stg}"
    doc = ids.Ids(
        title=title,
        author="BIM Coordinator",
        date=date.today().isoformat(),
        purpose=f"Đặc tả yêu cầu thông tin (EIR) — bộ môn {disc}, giai đoạn {stg}",
        description="Sinh từ QĐ 347/QĐ-BXD. Thước đo chấm bằng máy theo chuẩn buildingSMART IDS.",
    )
    bat_buoc = {str(m).strip().upper() for m in (bat_buoc or [])}
    n = 0
    for code, name, sdisc, smin, entity, facet in EIR:
        if STAGE_ORDER[stg] < STAGE_ORDER[smin]:
            continue
        if sdisc != "ALL" and disc != "ALL" and sdisc != disc:
            continue

        lop = _mo_rong_lop(entity)             # tự thêm lớp con (IfcWallStandardCase...)
        ver = _schema_dung_duoc(lop) or list(IFC_VERSIONS)
        if len(ver) < len(IFC_VERSIONS):
            print(f"  ⚠ {code}: chỉ dùng được trên {', '.join(ver)} "
                  f"(lớp {', '.join(lop)} không có ở schema còn lại) → đã khai đúng trong .ids")
        if not _schema_dung_duoc(lop):
            print(f"  ⛔ {code}: LUẬT CHẾT — {', '.join(lop)} không tồn tại ở bất kỳ schema "
                  f"nào trong {IFC_VERSIONS}, và cũng không có <Lớp>Type để lùi về. Sửa bảng EIR.")

        required = code.upper() in bat_buoc
        spec = ids.Specification(
            name=f"{code} — {name}",
            ifcVersion=ver,
            identifier=code,
            minOccurs=1 if required else 0,
            instructions=(f"Yêu cầu áp dụng từ giai đoạn {smin} (bộ môn {sdisc})."
                          + (" BẮT BUỘC phải có cấu kiện loại này trong mô hình."
                             if required else "")),
        )
        if len(lop) > 1:
            ent = ids.Entity(name=ids.Restriction(
                options={"enumeration": lop}, base="string"))
        else:
            ent = ids.Entity(name=lop[0])
        spec.applicability.append(ent)
        spec.requirements.append(_make_facet(facet))
        doc.specifications.append(spec)
        n += 1
    return doc, n


def main():
    ap = argparse.ArgumentParser(description="Sinh file IDS chuẩn từ EIR (QĐ 347)")
    ap.add_argument("--du-an", default="Du an", help="Tên dự án in vào IDS")
    ap.add_argument("--discipline", default="ALL", choices=["KT", "KC", "MEP", "ALL"])
    ap.add_argument("--stage", default="TKTC", choices=list(STAGE_ORDER))
    ap.add_argument("--all", action="store_true",
                    help="Sinh cả ma trận {KT,KC,MEP} × {TKKT,TKTC} vào thư mục -o")
    ap.add_argument("-o", "--output", required=True,
                    help="File .ids (chế độ đơn) hoặc thư mục (chế độ --all)")
    ap.add_argument("--bat-buoc", default="",
                    help="Danh sách mã EIR BẮT BUỘC phải có cấu kiện trong mô hình "
                         "(minOccurs=1), cách nhau bởi dấu phẩy. VD: EIR-KC-04,EIR-KC-06. "
                         "Mặc định rỗng: spec là 'optional' theo chuẩn IDS — mô hình thiếu "
                         "hẳn loại cấu kiện đó vẫn được tính ĐẠT.")
    args = ap.parse_args()

    bat_buoc = [m for m in (x.strip() for x in args.bat_buoc.split(",")) if m]
    if bat_buoc:
        print(f"Bắt buộc phải tồn tại (minOccurs=1): {', '.join(bat_buoc)}")

    slug = "".join(c if c.isalnum() else "_" for c in args.du_an)

    if args.all:
        outdir = Path(args.output)
        outdir.mkdir(parents=True, exist_ok=True)
        combos = [(d, s) for d in ("KT", "KC", "MEP") for s in ("TKKT", "TKTC")]
        print(f"Sinh {len(combos)} file IDS cho '{args.du_an}' -> {outdir}")
        for d, s in combos:
            doc, n = build_ids(args.du_an, d, s, bat_buoc)
            fp = outdir / f"EIR_{slug}_{d}_{s}.ids"
            doc.to_xml(str(fp))          # to_xml tự validate theo schema IDS
            print(f"  ✅ {fp.name}  ({n} spec)")
    else:
        doc, n = build_ids(args.du_an, args.discipline, args.stage, bat_buoc)
        fp = Path(args.output)
        fp.parent.mkdir(parents=True, exist_ok=True)
        doc.to_xml(str(fp))
        print(f"✅ {fp}  ({n} spec) — đã validate theo schema IDS")


if __name__ == "__main__":
    main()
