"""
ids_dophu.py — Đo ĐỘ PHỦ của một bộ .ids trên một mô hình IFC (dùng chung).
================================================================================
Vì sao cần: chuẩn IDS coi spec mặc định là 'optional' (minOccurs=0). Hệ quả là một
spec KHÔNG tìm thấy cấu kiện nào vẫn được ifctester tính là ĐẠT. Cộng thêm 2 kiểu
mù khác, ta có 3 cách để báo cáo "100% đạt" mà thực chất chưa kiểm gì:

  (1) Spec có, nhưng 0 cấu kiện áp dụng     → ifctester tính ĐẠT (minOccurs=0)
  (2) Cấu kiện có trong mô hình, nhưng KHÔNG spec nào áp dụng lên → không ai kiểm
  (3) Spec nhắm tới lớp IFC không tồn tại trong schema file VÀ cũng không có lớp
      <Lớp>Type tương ứng → luật chết, vĩnh viễn không khớp gì

Module này KHÔNG đổi phán quyết đạt/trượt của ifctester. Nó chỉ trả về dữ kiện để
lớp báo cáo nói rõ "đã kiểm được những gì", thay vì chỉ in % đạt.

HAI CƠ CHẾ CỦA ifctester CẦN BIẾT (đã đọc nguồn 0.8.5, đừng suy đoán lại):
  • facet Entity khớp lớp CHÍNH XÁC, không nhận lớp con (facet.py dùng
    include_subtypes=False). Nên IfcBeam KHÔNG bắt IfcBeamStandardCase.
  • NHƯNG nếu lớp không tồn tại trong schema, ifctester tự lùi sang <Lớp>Type rồi
    lấy các occurrence của type đó (facet.py:205-216). Nhờ vậy IFCPIPESEGMENT vẫn
    khớp đúng 231 ống trên file IFC2X3 (nơi chỉ có IfcFlowSegment + IfcPipeSegmentType).
    => KHÔNG được kết luận "luật chết" chỉ vì lớp thiếu trong schema.

Vì hai cơ chế trên, độ phủ ở đây đo bằng CHÍNH cấu kiện mà ifctester đã áp dụng
(spec.applicable_entities), không so khớp tên lớp — tránh mọi suy đoán sai.

Dùng:
    from ids_dophu import phan_tich, dong_canh_bao
    kq = phan_tich(doc, model)          # gọi SAU doc.validate(model)
    for d in dong_canh_bao(kq): print(d)
"""
from collections import Counter

from ifcopenshell import ifcopenshell_wrapper

# Lớp không tính vào độ phủ: hạ tầng bản vẽ, không phải cấu kiện cần kiểm.
BO_QUA = {
    "IFCGRID", "IFCGRIDAXIS", "IFCANNOTATION", "IFCOPENINGELEMENT",
    "IFCOPENINGSTANDARDCASE", "IFCVIRTUALELEMENT",
}


def _ma_spec(spec) -> str:
    """Mã EIR của spec. ids.open() không luôn nạp identifier → lấy tiền tố tên."""
    if getattr(spec, "identifier", None):
        return str(spec.identifier)
    ten = spec.name or ""
    return ten.split(" — ")[0].strip() or "—"


def _ten_lop_tu_facet(facet) -> tuple[set, bool]:
    """Rút tên lớp IFC từ facet Entity.

    Trả về (tập tên lớp in hoa, đọc_được). đọc_được=False khi tên lớp là Restriction
    dạng pattern/bound — không liệt kê được thành danh sách hữu hạn.
    """
    name = getattr(facet, "name", None)
    if name is None:
        return set(), False
    if isinstance(name, str):
        return {name.strip().upper()}, True
    options = getattr(name, "options", None) or {}
    if "enumeration" in options:
        return {str(v).strip().upper() for v in options["enumeration"]}, True
    return set(), False


def _lop_cua_spec(spec) -> tuple[set, bool]:
    lop, doc_duoc = set(), True
    for facet in spec.applicability:
        if type(facet).__name__ != "Entity":
            continue
        tap, ok = _ten_lop_tu_facet(facet)
        if ok:
            lop |= tap
        else:
            doc_duoc = False
    return lop, doc_duoc


def _schema_cua(model) -> str:
    return (getattr(model, "schema_identifier", None) or model.schema).upper()


def _co_trong_schema(ten_lop: str, schema: str) -> bool:
    try:
        s = ifcopenshell_wrapper.schema_by_name(schema)
    except Exception:
        return True          # không tra được schema → không dám kết luận
    try:
        s.declaration_by_name(ten_lop)
        return True
    except Exception:
        return False


def la_luat_chet(ten_lop: str, schema: str) -> bool:
    """Luật chết: lớp không có trong schema VÀ cũng không có <Lớp>Type để ifctester lùi về."""
    if _co_trong_schema(ten_lop, schema):
        return False
    if ten_lop.upper().endswith("TYPE"):
        return True
    return not _co_trong_schema(ten_lop + "Type", schema)


def phan_tich(doc, model) -> dict:
    """Trả về toàn bộ dữ kiện độ phủ. Gọi SAU doc.validate(model)."""
    schema = _schema_cua(model)

    # Cấu kiện ifctester THỰC SỰ đã áp dụng (đây là sự thật, không suy đoán theo tên lớp)
    da_ap_dung = set()
    for spec in doc.specifications:
        for e in spec.applicable_entities:
            da_ap_dung.add(e.id())

    tat_ca = [e for e in model.by_type("IfcProduct") if e.is_a().upper() not in BO_QUA]
    lop_phu, lop_khong_phu = Counter(), Counter()
    for e in tat_ca:
        (lop_phu if e.id() in da_ap_dung else lop_khong_phu)[e.is_a().upper()] += 1

    spec_khong_ap_dung, spec_luat_chet, spec_lech_schema, spec_khong_doc_duoc = [], [], [], []
    for spec in doc.specifications:
        ma, ten = _ma_spec(spec), spec.name or ""
        lop, doc_duoc = _lop_cua_spec(spec)
        if not doc_duoc:
            spec_khong_doc_duoc.append(ma)
        if not spec.applicable_entities:
            spec_khong_ap_dung.append((ma, ten))
            chet = sorted(l for l in lop if la_luat_chet(l, schema))
            if chet:
                spec_luat_chet.append((ma, ten, chet))
        ver = spec.ifcVersion if isinstance(spec.ifcVersion, (list, tuple)) else [spec.ifcVersion]
        if schema not in [str(v).upper() for v in ver]:
            spec_lech_schema.append((ma, ten, list(ver)))

    ck_phu, ck_tong = sum(lop_phu.values()), len(tat_ca)
    return {
        "schema": schema,
        "lop_phu": dict(lop_phu.most_common()),
        "lop_khong_phu": dict(lop_khong_phu.most_common()),
        "cau_kien_phu": ck_phu,
        "cau_kien_tong": ck_tong,
        "phan_tram_phu": round(ck_phu / ck_tong * 100) if ck_tong else None,
        "spec_khong_ap_dung": spec_khong_ap_dung,
        "spec_luat_chet": spec_luat_chet,
        "spec_lech_schema": spec_lech_schema,
        "spec_khong_doc_duoc": spec_khong_doc_duoc,
        "tong_spec": len(doc.specifications),
    }


def co_canh_bao(kq: dict) -> bool:
    return bool(kq["spec_khong_ap_dung"] or kq["lop_khong_phu"]
                or kq["spec_luat_chet"] or kq["spec_lech_schema"])


def dong_canh_bao(kq: dict) -> list:
    """Các dòng chữ để in màn hình / ghi vào báo cáo."""
    d = []
    n_ka = len(kq["spec_khong_ap_dung"])
    d.append(f"  Độ phủ cấu kiện  : {kq['cau_kien_phu']}/{kq['cau_kien_tong']} "
             f"({kq['phan_tram_phu']}%) — schema {kq['schema']}")
    d.append(f"  Yêu cầu kiểm thật: {kq['tong_spec'] - n_ka}/{kq['tong_spec']}"
             + (f"  (còn {n_ka} yêu cầu KHÔNG có cấu kiện nào để kiểm)" if n_ka else ""))

    if kq["spec_luat_chet"]:
        d.append(f"  ⛔ LUẬT CHẾT — lớp không có trong schema {kq['schema']} và cũng không có "
                 f"<Lớp>Type để lùi về, vĩnh viễn không khớp gì:")
        for ma, ten, chet in kq["spec_luat_chet"]:
            d.append(f"       {ma:12} {', '.join(chet)} — {ten[:46]}")

    if kq["spec_khong_ap_dung"]:
        d.append("  ⚠ KHÔNG KIỂM ĐƯỢC (0 cấu kiện áp dụng — chuẩn IDS vẫn tính là ĐẠT):")
        for ma, ten in kq["spec_khong_ap_dung"]:
            d.append(f"       {ma:12} {ten[:58]}")

    if kq["lop_khong_phu"]:
        tong = sum(kq["lop_khong_phu"].values())
        d.append(f"  ⚠ KHÔNG YÊU CẦU NÀO PHỦ — {tong} cấu kiện trong mô hình không spec nào áp dụng lên:")
        for cls, n in list(kq["lop_khong_phu"].items())[:12]:
            d.append(f"       {cls:30} {n}")
        if len(kq["lop_khong_phu"]) > 12:
            d.append(f"       ... và {len(kq['lop_khong_phu']) - 12} lớp nữa")

    if kq["spec_lech_schema"]:
        d.append(f"  ⚠ Spec không khai schema {kq['schema']} của file này:")
        for ma, ten, ver in kq["spec_lech_schema"]:
            d.append(f"       {ma:12} (khai {', '.join(map(str, ver))}) — {ten[:40]}")

    if kq["spec_khong_doc_duoc"]:
        d.append("  ℹ Không liệt kê được lớp của spec (Restriction dạng pattern): "
                 + ", ".join(map(str, kq["spec_khong_doc_duoc"][:6])))

    if not co_canh_bao(kq):
        d.append("  ✅ Không có cảnh báo độ phủ: mọi yêu cầu đều tìm được cấu kiện, "
                 "mọi cấu kiện đều có yêu cầu phủ.")
    return d
