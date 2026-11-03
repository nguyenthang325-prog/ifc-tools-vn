"""Kiểm nhanh trên mô hình mẫu Duplex (buildingSMART, CC BY 4.0).

Chỉ chạy khi có biến môi trường DUPLEX_PATH trỏ tới Duplex_A_20110907.ifc
(xem README mục "Thử với mô hình mẫu"); không có thì bỏ qua.

    DUPLEX_PATH=/duong/dan/Duplex_A_20110907.ifc python -m unittest discover tests -v
"""
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

DUPLEX = os.environ.get("DUPLEX_PATH")
THU_NHA = os.environ.get("THU_NHA_PATH")   # nhà thử IFC4 Revit 2027 (không công khai)


def chay(module, *args, out):
    env = {**os.environ, "IFC_TOOLS_OUTPUT": str(out), "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, "-m", f"ifc_tools_vn.checks.{module}", *args],
                       capture_output=True, text=True, encoding="utf-8", env=env)
    assert r.returncode == 0, f"{module} lỗi:\n{r.stdout}\n{r.stderr}"
    return r.stdout


def tom_tat(xlsx):
    """{(nhóm, đại lượng, ĐVT): (tổng, số cấu kiện)} từ sheet Tom_Tat_BoQ."""
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    rows = {(r[1], r[2], r[3]): (r[4], r[5]) for r in wb["Tom_Tat_BoQ"].iter_rows(min_row=4, values_only=True) if r[1]}
    wb.close()
    return rows


@unittest.skipUnless(DUPLEX, "đặt DUPLEX_PATH để chạy")
class TestDuplex(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_ids(self):
        ids = self.out / "EIR_Duplex_KT_TKTC.ids"
        chay("eir_to_ids", "--du-an", "Duplex", "--discipline", "KT", "--stage", "TKTC", "-o", str(ids), out=self.out)
        log = chay("ids_run", "--ids", str(ids), "--ifc", DUPLEX, "--project", "Duplex", out=self.out)
        self.assertIn("8/10 spec", log)
        self.assertIn("101/245", log)
        xlsx = next(self.out.rglob("IDS_KetQua_Duplex_*.xlsx"))
        wb = openpyxl.load_workbook(xlsx, read_only=True)
        rows = list(wb["Loi_Chi_Tiet"].iter_rows(values_only=True))
        wb.close()  # Windows giữ khoá file nếu không đóng → không dọn được thư mục tạm
        i = next(k for k, r in enumerate(rows) if r and "GUID" in r)
        c = rows[i].index("GUID")
        self.assertEqual(len({r[c] for r in rows[i + 1:] if r and r[c]}), 58)

    def test_lod(self):
        log = chay("lod_checker", "--ifc", DUPLEX, "--stage", "TKTC", "--discipline", "KT", "--project", "Duplex", out=self.out)
        self.assertIn("Tổng elements : 197", log)

    def test_qto_bo_sung(self):
        log = chay("ifc_qto", "--ifc", DUPLEX, "--discipline", "ALL", "--project", "Duplex", "--bo-sung-qto", out=self.out)
        # 0.1.0 ra 216; từ 0.2.0 đếm thêm 2 vế thang IfcStairFlight (trước bị bỏ sót — lỗi L5)
        self.assertIn("Tổng elements     : 218", log)
        self.assertTrue(any(self.out.rglob("Duplex_A_20110907_BoSungQTO.ifc")))
        boq = tom_tat(next(self.out.rglob("IFC_QTO_Duplex_*.xlsx")))
        # "Sàn/Mái" 125,33 m³ của 0.1.0 = Sàn 64,583 + tấm con của mái 60,747 (lỗi L3)
        self.assertEqual(boq[("Sàn", "Thể tích", "m³")], (64.583, 20))
        self.assertEqual(boq[("Mái", "Thể tích", "m³")], (60.747, 1))
        self.assertEqual(boq[("Cầu thang", "Thể tích", "m³")], (0.417, 2))


@unittest.skipUnless(THU_NHA, "đặt THU_NHA_PATH để chạy")
class TestNhaThu(unittest.TestCase):
    """Nhà thử 2 tầng dựng bằng Horizun (Revit 2027, IFC4) — mọi dòng đã đối chiếu tay với Revit."""

    def test_bang_khoi_luong(self):
        with tempfile.TemporaryDirectory() as tmp:
            chay("ifc_qto", "--ifc", THU_NHA, "--project", "ThuNha", out=tmp)
            boq = tom_tat(next(Path(tmp).rglob("IFC_QTO_ThuNha_*.xlsx")))
        mong_doi = {
            ("Tường", "Thể tích", "m³"): (49.521, 9),            # Revit Net (Gross 50,688 là sai — L2)
            ("Sàn", "Thể tích", "m³"): (7.74, 1),                # 60 × 0,15 − lỗ 8,4 × 0,15 (L3)
            ("Mái", "Thể tích", "m³"): (18.323, 2),              # 81,434 × 0,225; Net tấm mái là rác (L4)
            ("Mái", "Diện tích mái nghiêng", "m²"): (81.434, 1),
            ("Mái", "Diện tích mái chiếu bằng", "m²"): (78, 1),
            ("Cầu thang", "Thể tích", "m³"): (2.149, 3),         # 2 vế + chiếu nghỉ (L5)
            ("Cửa", "Diện tích W×H", "m²"): (1.8, 1),            # 900 × 2000, không phải Area 4,93 (L6)
            ("Cửa", "Số cái", "cái"): (1, 1),
            ("Cửa sổ", "Diện tích W×H", "m²"): (2.835, 2),
            ("Cửa sổ", "Số cái", "cái"): (2, 2),
            ("Cột", "Thể tích", "m³"): (1.728, 12),
            ("Dầm", "Thể tích", "m³"): (4.512, 14),
        }
        for k, v in mong_doi.items():
            self.assertEqual(boq.get(k), v, k)


if __name__ == "__main__":
    unittest.main()
