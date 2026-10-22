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


def chay(module, *args, out):
    env = {**os.environ, "IFC_TOOLS_OUTPUT": str(out), "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, "-m", f"ifc_tools_vn.checks.{module}", *args],
                       capture_output=True, text=True, encoding="utf-8", env=env)
    assert r.returncode == 0, f"{module} lỗi:\n{r.stdout}\n{r.stderr}"
    return r.stdout


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
        self.assertIn("Tổng elements     : 216", log)
        self.assertTrue(any(self.out.rglob("Duplex_A_20110907_BoSungQTO.ifc")))


if __name__ == "__main__":
    unittest.main()
