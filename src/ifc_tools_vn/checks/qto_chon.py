"""Chọn khối lượng chính cho TỪNG cấu kiện IFC — dùng chung cho ifc_qto, model_payment, payment_from_model.

Vì sao có module này (rà soát 03/11/2026 trên 7 IFC Revit thật, xem CHANGELOG ifc-tools-vn 0.2.0):
  L1  Cộng theo (nhóm, TÊN đại lượng) làm rơi cấu kiện chỉ có NetVolume khỏi tổng
      → ở đây mỗi cấu kiện lấy đúng 1 KL chính, rồi người gọi mới cộng theo (nhóm, đơn vị).
  L2  GrossVolume/NetVolume hoà điểm → chọn theo thứ tự thuộc tính trong file (sàn 15 tầng hụt 63 %).
  L4  Revit xuất Net/Gross rác (tấm con của mái IFC4, sàn chia Part) → không quy tắc cố định
      nào đúng, nên dùng TRỌNG TÀI HÌNH HỌC (chon_the_tich).
  L3/L5  Phân nhóm theo cấu kiện cha: tấm con của mái → Mái, chiếu nghỉ + vế thang → Cầu thang.
  L6  Area của cửa IFC4 Revit 2027 không phải diện tích ô cửa → dùng W × H.

Có 2 bản sao GIỐNG HỆT, sửa 1 nơi phải chép sang nơi kia (test chéo kiểm 2 bản bằng nhau):
  gói ifc-tools-vn:      src/ifc_tools_vn/checks/qto_chon.py
  bộ script của tác giả: PHAT_TRIEN_BIM/_chung/qto_chon.py (dùng chung với 2 script thanh toán)
"""
from collections import namedtuple

import ifcopenshell.util.element as ifc_util
import ifcopenshell.util.unit as unit_util

# Ngưỡng (tỷ lệ). Mới thử trên 7 file; có mô hình dầm thép I lệch Gross/Net ~3 % → để thành tham số.
NGUONG_GROSS_NET = 0.005   # Gross ≈ Net trong ngưỡng này → coi như bằng, lấy Net
NGUONG_HINH_HOC  = 0.02    # Net/Gross khớp thể tích hình học trong ngưỡng này
NGUONG_CUA       = 0.05    # Area của cửa lệch W × H quá ngưỡng này → gắn cờ

# Cờ (ghi nguyên văn vào Excel)
CO_NET_SAI    = "Net sai → lấy Gross (Gross khớp hình học)"
CO_QTO_SAI    = "Gross & Net đều lệch hình học → lấy thể tích hình học"
CO_CHUA_KIEM  = "Gross ≠ Net, không đo được hình học → tạm lấy Net (chưa kiểm)"
CO_TAT_KIEM   = "Gross ≠ Net, đã tắt kiểm hình học → tạm lấy Net (chưa kiểm)"
CO_VE_THANG   = "IFC không có KL vế thang, bóc từ Revit"
CO_CUA_AREA   = "Area của Qto lệch W × H quá 5 % → bỏ Area, dùng W × H"
CO_CUA_KHONG_WH = "Cửa không có Width/Height → dùng Area của Qto (chưa kiểm)"
CO_KL_O_CON   = "KL nằm ở cấu kiện con → không cộng ở cấu kiện cha (tránh trùng)"

# Cấu kiện "vỏ chứa": khi đã chia thành cấu kiện con (IsDecomposedBy) thì thể tích nằm ở con.
LOAI_VO_CHUA = {"IfcRoof", "IfcStair", "IfcRamp"}
LOAI_CUA = {"IfcDoor", "IfcWindow"}

# Thứ tự ưu tiên diện tích khi cấu kiện KHÔNG có thể tích (Net = đã trừ lỗ, là số thanh toán).
_UU_TIEN_DIEN_TICH = ["NetSideArea", "GrossSideArea", "NetArea", "GrossArea",
                      "NetFootprintArea", "GrossFootprintArea", "Area"]
_UU_TIEN_CHIEU_DAI = ["Length", "Perimeter"]

# gia_tri: số (đơn vị project cho Qto, m² cho W × H); dvt: "m³"/"m²"/"cái"/đơn vị dài;
# nhan: tên đại lượng hiển thị; nguon: lấy từ đâu (NetVolume / GrossVolume / Hình học / W×H ...)
KL = namedtuple("KL", "gia_tri dvt nhan nguon")


def luoi_kin(verts, faces) -> bool:
    """Lưới kín = mọi cạnh tam giác dùng chung đúng 2 mặt. Lưới hở thì thể tích vô nghĩa."""
    import numpy as np
    e = np.sort(np.concatenate([faces[:, [0, 1]], faces[:, [1, 2]], faces[:, [2, 0]]]), axis=1)
    _, cnt = np.unique(e, axis=0, return_counts=True)
    return bool((cnt == 2).all())


class BoiCanh:
    """Thông tin dùng chung cho 1 file IFC: hệ số đơn vị + bộ dựng hình học (khởi tạo lười)."""

    def __init__(self, ifc_file, kiem_hinh_hoc=True,
                 nguong_gross_net=NGUONG_GROSS_NET, nguong_hinh_hoc=NGUONG_HINH_HOC,
                 nguong_cua=NGUONG_CUA):
        self.f = ifc_file
        self.kiem_hinh_hoc = kiem_hinh_hoc
        self.nguong_gross_net = nguong_gross_net
        self.nguong_hinh_hoc = nguong_hinh_hoc
        self.nguong_cua = nguong_cua
        self.ls = unit_util.calculate_unit_scale(ifc_file, "LENGTHUNIT")   # đơn vị dài project → m
        self.as_ = unit_util.calculate_unit_scale(ifc_file, "AREAUNIT")    # → m²
        self.vs = unit_util.calculate_unit_scale(ifc_file, "VOLUMEUNIT")   # → m³
        self._settings = None
        self.so_lan_do = 0

    def the_tich_hinh_hoc(self, el):
        """Thể tích đo từ hình học (đơn vị thể tích project), None nếu không đo tin được:
        không có Body, là SurfaceModel, lỗi dựng, lưới hở, hoặc V vượt hộp bao."""
        if not el.Representation:
            return None
        if any(r.RepresentationType == "SurfaceModel" for r in el.Representation.Representations):
            return None
        import numpy as np
        import ifcopenshell.geom
        import ifcopenshell.util.shape as shape_util
        if self._settings is None:
            self._settings = ifcopenshell.geom.settings()                 # kết quả theo m
        self.so_lan_do += 1
        try:
            sh = ifcopenshell.geom.create_shape(self._settings, el)      # giữ biến sh: viết gọn 1 dòng thì verts rỗng
        except Exception:
            return None
        geo = sh.geometry
        if not len(geo.verts):
            return None
        verts = np.array(geo.verts).reshape(-1, 3)
        faces = np.array(geo.faces).reshape(-1, 3)
        if not luoi_kin(verts, faces):
            return None
        v = shape_util.get_volume(geo)
        bbox = shape_util.get_x(geo) * shape_util.get_y(geo) * shape_util.get_z(geo)
        if not (0 < v <= bbox * 1.001):
            return None
        return v / self.vs


def doc_qto(el) -> dict:
    """{tên đại lượng: số} từ các bộ Qto_*/BaseQuantities. Trùng tên giữa 2 bộ → giữ bộ gặp trước."""
    try:
        psets = ifc_util.get_psets(el)
    except Exception:
        return {}
    out = {}
    for pname, props in psets.items():
        if not (pname.startswith("Qto_") or pname == "BaseQuantities") or not isinstance(props, dict):
            continue
        for p, x in props.items():
            if p == "id" or isinstance(x, bool) or not isinstance(x, (int, float)):
                continue
            out.setdefault(p, float(x))
    return out


def cau_kien_cha(el):
    for rel in getattr(el, "Decomposes", None) or []:
        return rel.RelatingObject
    return None


def co_cau_kien_con(el) -> bool:
    return any(rel.RelatedObjects for rel in (getattr(el, "IsDecomposedBy", None) or []))


def nhom_theo_cha(el, nhom_mac_dinh: str) -> str:
    """Nhóm KL theo cấu kiện cha: IfcSlab con của IfcRoof → Mái; con của IfcStair → Cầu thang;
    IfcStairFlight → Cầu thang. Còn lại giữ nhóm mặc định theo loại của người gọi."""
    t = el.is_a()
    if t == "IfcStairFlight":
        return "Cầu thang"
    if t == "IfcSlab":
        p = cau_kien_cha(el)
        if p is not None and p.is_a("IfcRoof"):
            return "Mái"
        if p is not None and p.is_a("IfcStair"):
            return "Cầu thang"
    return nhom_mac_dinh


def _gan(a, b, tol) -> bool:
    return a is not None and b is not None and abs(a - b) <= tol * max(abs(a), abs(b), 1e-9)


def chon_the_tich(el, q: dict, bc: BoiCanh):
    """Trọng tài thể tích. Trả về (giá trị, nguồn, cờ) hoặc None nếu cấu kiện không có thể tích.
    1. Chỉ có 1 trong Gross/Net → lấy giá trị đó.
    2. Gross ≈ Net (≤ 0,5 %) → Net.
    3. Lệch → đo hình học: Net khớp → Net; Gross khớp → Gross (cờ Net sai);
       không khớp → hình học (cờ Qto sai).
    4. Không đo được hình học → Net (cờ chưa kiểm)."""
    g, n = q.get("GrossVolume"), q.get("NetVolume")
    if g is None and n is None:
        if "Volume" in q:
            return q["Volume"], "Volume", ""
        for k in sorted(q):
            if "volume" in k.lower():
                return q[k], k, ""
        return None
    if n is None:
        return g, "GrossVolume", ""
    if g is None:
        return n, "NetVolume", ""
    if _gan(g, n, bc.nguong_gross_net):
        return n, "NetVolume", ""
    if not bc.kiem_hinh_hoc:
        return n, "NetVolume", CO_TAT_KIEM
    v = bc.the_tich_hinh_hoc(el)
    if v is None:
        return n, "NetVolume", CO_CHUA_KIEM
    if _gan(n, v, bc.nguong_hinh_hoc):
        return n, "NetVolume (khớp hình học)", ""
    if _gan(g, v, bc.nguong_hinh_hoc):
        return g, "GrossVolume (khớp hình học)", CO_NET_SAI
    return round(v, 4), "Hình học", CO_QTO_SAI


def _lay_theo_uu_tien(q: dict, uu_tien: list, tu_khoa: str):
    for k in uu_tien:
        if k in q:
            return k, q[k]
    for k in sorted(q):                       # tên khác (GrossCeilingArea ...) — thứ tự cố định
        if tu_khoa in k.lower():
            return k, q[k]
    return None


def chon_kl(el, bc: BoiCanh, nhom_mac_dinh: str, dvt_dai=None) -> dict:
    """KL của 1 cấu kiện. Trả về dict:
        nhom  — nhóm theo cha (nhom_theo_cha)
        chinh — KL chính (KL) hoặc None — số mà 3 script cùng cộng
        phu   — list KL phụ (số cái cửa, diện tích chiếu bằng mái ...), chỉ ifc_qto hiển thị
        co    — list cờ
    dvt_dai: ký hiệu đơn vị dài project ("mm"...). None = không nhận Length/Perimeter làm KL chính
    (2 script thanh toán: không thanh toán bê tông theo mét — diễn tập 06/07)."""
    t = el.is_a()
    q = doc_qto(el)
    kq = {"nhom": nhom_theo_cha(el, nhom_mac_dinh), "chinh": None, "phu": [], "co": []}

    # ── Cửa đi / cửa sổ: W × H + số cái (L6) ──
    if t in LOAI_CUA:
        kq["phu"].append(KL(1, "cái", "Số cái", "Đếm"))
        w = q.get("Width") or getattr(el, "OverallWidth", None)
        h = q.get("Height") or getattr(el, "OverallHeight", None)
        a = q.get("Area")
        if w and h:
            wh = round(w * h * bc.ls * bc.ls, 4)
            kq["chinh"] = KL(wh, "m²", "Diện tích W×H", "W×H")
            if a is not None and not _gan(a * bc.as_, wh, bc.nguong_cua):
                kq["co"].append(CO_CUA_AREA)
        elif a is not None:
            kq["chinh"] = KL(round(a * bc.as_, 4), "m²", "Diện tích W×H", "Area (Qto)")
            kq["co"].append(CO_CUA_KHONG_WH)
        return kq

    vo_chua = t in LOAI_VO_CHUA and co_cau_kien_con(el)

    # ── Mái: diện tích nghiêng (GrossArea) + chiếu bằng (ProjectedArea) từ IfcRoof ──
    dt_mai = []
    if t == "IfcRoof":
        nghieng = _lay_theo_uu_tien(q, ["GrossArea", "NetArea", "Area"], "\0")
        if nghieng:
            dt_mai.append(KL(nghieng[1], "m²", "Diện tích mái nghiêng", nghieng[0]))
        if "ProjectedArea" in q:
            dt_mai.append(KL(q["ProjectedArea"], "m²", "Diện tích mái chiếu bằng", "ProjectedArea"))

    # ── Thể tích (trọng tài) ──
    tt = None if vo_chua else chon_the_tich(el, q, bc)
    if tt is not None:
        v, nguon, co = tt
        kq["chinh"] = KL(v, "m³", "Thể tích", nguon)
        if co:
            kq["co"].append(co)
        kq["phu"].extend(dt_mai)
        return kq
    if vo_chua and t != "IfcRoof" and any("volume" in k.lower() for k in q):
        kq["co"].append(CO_KL_O_CON)
    if dt_mai:
        kq["chinh"], kq["phu"] = dt_mai[0], dt_mai[1:]
        return kq
    if vo_chua:
        return kq
    if t == "IfcStairFlight":
        kq["co"].append(CO_VE_THANG)
        return kq

    # ── Không có thể tích: diện tích → (ifc_qto) chiều dài ──
    dt = _lay_theo_uu_tien(q, _UU_TIEN_DIEN_TICH, "area")
    if dt:
        kq["chinh"] = KL(dt[1], "m²", "Diện tích", dt[0])
        return kq
    if dvt_dai is not None:
        cd = _lay_theo_uu_tien(q, _UU_TIEN_CHIEU_DAI, "\0")
        if cd:
            kq["chinh"] = KL(cd[1], dvt_dai, "Chiều dài", cd[0])
    return kq
