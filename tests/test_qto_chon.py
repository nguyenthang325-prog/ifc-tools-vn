"""Kiểm quy tắc chọn KL chính (checks/qto_chon.py) trên IFC tổng hợp dựng trong bộ nhớ.

Không cần file mẫu nên chạy được trên CI. Mỗi test ứng với 1 lỗi đã gặp trên mô hình Revit thật
(L1–L6, xem CHANGELOG 0.2.0).

    python -m unittest tests.test_qto_chon -v
"""
import unittest

import ifcopenshell
import ifcopenshell.api

from ifc_tools_vn.checks import qto_chon as Q


class MoHinh:
    """IFC4 nhỏ: đơn vị mm / m² / m³, 1 context Body."""

    def __init__(self):
        self.f = ifcopenshell.file(schema="IFC4")
        ifcopenshell.api.run("root.create_entity", self.f, ifc_class="IfcProject", name="Thu")
        ifcopenshell.api.run("unit.assign_unit", self.f)          # mặc định: mm, m², m³
        model = ifcopenshell.api.run("context.add_context", self.f, context_type="Model")
        self.body = ifcopenshell.api.run("context.add_context", self.f, context_type="Model",
                                         context_identifier="Body", target_view="MODEL_VIEW", parent=model)

    def cau_kien(self, ifc_class, qto=None, hop=None):
        """hop = (dài, rộng, dày) theo MÉT (api tự đổi sang mm của project) → khối hộp đặc."""
        el = ifcopenshell.api.run("root.create_entity", self.f, ifc_class=ifc_class)
        if hop:
            x, y, z = hop
            rep = ifcopenshell.api.run("geometry.add_slab_representation", self.f, context=self.body, depth=z,
                                       polyline=[(0.0, 0.0), (x, 0.0), (x, y), (0.0, y)])
            ifcopenshell.api.run("geometry.assign_representation", self.f, product=el, representation=rep)
            ifcopenshell.api.run("geometry.edit_object_placement", self.f, product=el)
        if qto:
            q = ifcopenshell.api.run("pset.add_qto", self.f, product=el, name="BaseQuantities")
            ifcopenshell.api.run("pset.edit_qto", self.f, qto=q, properties=qto)
        return el

    def gop(self, cha, *con):
        ifcopenshell.api.run("aggregate.assign_object", self.f, products=list(con), relating_object=cha)

    def chon(self, el, nhom="X", **kw):
        return Q.chon_kl(el, Q.BoiCanh(self.f, **kw), nhom, dvt_dai="mm")


class TestTheTich(unittest.TestCase):
    def setUp(self):
        self.m = MoHinh()

    def test_L1_chi_co_net_van_duoc_tinh(self):
        kq = self.m.chon(self.m.cau_kien("IfcWall", {"NetVolume": 3.0}))
        self.assertEqual((kq["chinh"].gia_tri, kq["chinh"].dvt, kq["chinh"].nguon), (3.0, "m³", "NetVolume"))

    def test_L2_gross_xap_xi_net_lay_net_khong_phu_thuoc_thu_tu(self):
        for qto in ({"GrossVolume": 2.005, "NetVolume": 2.0}, {"NetVolume": 2.0, "GrossVolume": 2.005}):
            kq = self.m.chon(self.m.cau_kien("IfcWall", qto))
            self.assertEqual(kq["chinh"].gia_tri, 2.0)
            self.assertEqual(kq["co"], [])

    def test_trong_tai_net_khop_hinh_hoc(self):            # hộp 1 × 1 × 0,2 m = 0,2 m³
        kq = self.m.chon(self.m.cau_kien("IfcSlab", {"GrossVolume": 0.5, "NetVolume": 0.2}, (1.0, 1.0, 0.2)))
        self.assertEqual(kq["chinh"].gia_tri, 0.2)
        self.assertEqual(kq["co"], [])

    def test_L4_net_rac_lay_gross_khop_hinh_hoc(self):
        kq = self.m.chon(self.m.cau_kien("IfcSlab", {"GrossVolume": 0.2, "NetVolume": 0.05}, (1.0, 1.0, 0.2)))
        self.assertEqual(kq["chinh"].gia_tri, 0.2)
        self.assertEqual(kq["co"], [Q.CO_NET_SAI])

    def test_L4_ca_hai_rac_lay_hinh_hoc(self):
        kq = self.m.chon(self.m.cau_kien("IfcSlab", {"GrossVolume": 0.9, "NetVolume": 0.6}, (1.0, 1.0, 0.2)))
        self.assertAlmostEqual(kq["chinh"].gia_tri, 0.2, places=4)
        self.assertEqual(kq["co"], [Q.CO_QTO_SAI])

    def test_khong_hinh_hoc_lay_net_gan_co_chua_kiem(self):
        kq = self.m.chon(self.m.cau_kien("IfcSlab", {"GrossVolume": 0.9, "NetVolume": 0.6}))
        self.assertEqual((kq["chinh"].gia_tri, kq["co"]), (0.6, [Q.CO_CHUA_KIEM]))

    def test_tat_kiem_hinh_hoc(self):
        el = self.m.cau_kien("IfcSlab", {"GrossVolume": 0.2, "NetVolume": 0.05}, (1.0, 1.0, 0.2))
        kq = self.m.chon(el, kiem_hinh_hoc=False)
        self.assertEqual((kq["chinh"].gia_tri, kq["co"]), (0.05, [Q.CO_TAT_KIEM]))


class TestNhomTheoCha(unittest.TestCase):
    def setUp(self):
        self.m = MoHinh()

    def test_L3_tam_con_cua_mai_vao_nhom_mai_dien_tich_lay_o_ifcroof(self):
        mai = self.m.cau_kien("IfcRoof", {"GrossArea": 81.4, "ProjectedArea": 78.0, "GrossVolume": 99.0})
        tam = self.m.cau_kien("IfcSlab", {"NetVolume": 9.0})
        self.m.gop(mai, tam)
        kq_tam = self.m.chon(tam, "Sàn")
        self.assertEqual((kq_tam["nhom"], kq_tam["chinh"].gia_tri), ("Mái", 9.0))
        kq_mai = self.m.chon(mai, "Mái")
        self.assertEqual((kq_mai["chinh"].gia_tri, kq_mai["chinh"].nhan), (81.4, "Diện tích mái nghiêng"))
        self.assertEqual([k.gia_tri for k in kq_mai["phu"]], [78.0])   # chiếu bằng; thể tích ở tấm con

    def test_L5_cau_thang(self):
        thang = self.m.cau_kien("IfcStair")
        ve = self.m.cau_kien("IfcStairFlight", {"NetVolume": 0.7})
        chieu_nghi = self.m.cau_kien("IfcSlab", {"NetVolume": 0.74})
        self.m.gop(thang, ve, chieu_nghi)
        self.assertEqual(self.m.chon(chieu_nghi, "Sàn")["nhom"], "Cầu thang")
        self.assertEqual(self.m.chon(ve, "Cầu thang")["chinh"].gia_tri, 0.7)
        self.assertIsNone(self.m.chon(thang, "Cầu thang")["chinh"])

    def test_L5_ve_thang_khong_co_the_tich_gan_co(self):    # IFC2X3 Revit
        kq = self.m.chon(self.m.cau_kien("IfcStairFlight", {"Length": 3000.0}), "Cầu thang")
        self.assertIsNone(kq["chinh"])
        self.assertEqual(kq["co"], [Q.CO_VE_THANG])


class TestCua(unittest.TestCase):
    def test_L6_cua_lay_w_x_h_bo_area_sai(self):
        m = MoHinh()
        kq = m.chon(m.cau_kien("IfcDoor", {"Width": 900.0, "Height": 2000.0, "Area": 4.93}), "Cửa")
        self.assertEqual((kq["chinh"].gia_tri, kq["chinh"].dvt), (1.8, "m²"))
        self.assertEqual([(k.gia_tri, k.dvt) for k in kq["phu"]], [(1, "cái")])
        self.assertEqual(kq["co"], [Q.CO_CUA_AREA])

    def test_cua_area_dung_khong_gan_co(self):
        m = MoHinh()
        kq = m.chon(m.cau_kien("IfcWindow", {"Width": 1050.0, "Height": 1350.0, "Area": 1.4175}), "Cửa sổ")
        self.assertEqual(kq["co"], [])


class TestDienTichChieuDai(unittest.TestCase):
    def test_khong_the_tich_lay_net_area_truoc(self):
        m = MoHinh()
        kq = m.chon(m.cau_kien("IfcCovering", {"GrossArea": 12.0, "NetArea": 10.0}))
        self.assertEqual((kq["chinh"].gia_tri, kq["chinh"].nguon), (10.0, "NetArea"))

    def test_thanh_toan_khong_nhan_chieu_dai(self):
        m = MoHinh()
        el = m.cau_kien("IfcRailing", {"Length": 5000.0})
        self.assertEqual(m.chon(el)["chinh"].dvt, "mm")                         # ifc_qto: có
        self.assertIsNone(Q.chon_kl(el, Q.BoiCanh(m.f), "Lan can")["chinh"])     # thanh toán: không


if __name__ == "__main__":
    unittest.main()
