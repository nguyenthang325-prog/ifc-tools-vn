#!/usr/bin/env python3
"""
IFC Tools VN — MCP Server
=========================
Kết nối Claude Desktop (hoặc client MCP khác) với các công cụ kiểm tra mô hình IFC.
5 tools: kiem_tra_toan_dien, kiem_tra_ids, kiem_tra_lod,
         boc_khoi_luong, tong_hop_audit

Kết quả ghi vào biến môi trường IFC_TOOLS_OUTPUT; nếu không đặt thì vào
~/ifc_tools_output/ (thư mục làm việc của client MCP thường không cố định).
"""

import subprocess
import json
import os
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP

OUTPUT = Path(os.environ.get("IFC_TOOLS_OUTPUT") or Path.home() / "ifc_tools_output")
AUDIT  = Path(__file__).parent / "checks"

mcp = FastMCP("IFC Tools VN")


# ── Helpers ─────────────────────────────────────────────────────────────────

def run(script: Path, args: list, timeout: int = 600) -> dict:
    """Chạy script Python, trả về stdout + trạng thái."""
    try:
        r = subprocess.run(
            [sys.executable, str(script)] + args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env={**os.environ, "IFC_TOOLS_OUTPUT": str(OUTPUT)},
            timeout=timeout,
            # Đóng stdin: pipeline spawn tiếp script con, nếu để mặc định chúng
            # thừa kế kênh stdio JSON-RPC của MCP server → deadlock → timeout 600s.
            stdin=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "out": "",
                "err": f"Script chạy quá {timeout}s nên bị dừng: {script.name}"}
    return {"ok": r.returncode == 0, "out": r.stdout, "err": r.stderr}

def ok(data: dict) -> str:
    return json.dumps({"status": "ok", **data}, ensure_ascii=False, indent=2)

def err(msg: str) -> str:
    return json.dumps({"status": "error", "message": msg}, ensure_ascii=False, indent=2)


# ── Tools ────────────────────────────────────────────────────────────────────

@mcp.tool()
def kiem_tra_toan_dien(
    ifc_path: str,
    project: str = "",
    stage: str = "TKKT",
    discipline: str = "ALL",
) -> str:
    """Chạy toàn bộ BIM audit pipeline (IDS + LOD + QTO) và xuất BIM Health Dashboard.

    Args:
        ifc_path:   Đường dẫn đầy đủ tới file IFC
        project:    Tên dự án (tự nhận diện từ tên file nếu để trống)
        stage:      Giai đoạn: TKCS / TKKT / TKTC / AS-BUILT (mặc định: TKKT)
        discipline: Discipline: KT / KC / MEP / ALL (mặc định: ALL)
    """
    r = run(
        AUDIT / "bim_audit_pipeline.py",
        ["--ifc", ifc_path, "--project", project, "--stage", stage, "--discipline", discipline],
    )
    if not r["ok"]:
        return err(r["err"] or "Script thất bại — kiểm tra đường dẫn IFC và ifcopenshell")
    return ok({
        "console": r["out"],
        "ghi_chu": f"File BIM_HealthDashboard_*.xlsx đã xuất vào {OUTPUT}",
    })


@mcp.tool()
def kiem_tra_ids(
    ifc_path: str,
    discipline: str = "ALL",
    stage: str = "TKKT",
    project: str = "",
) -> str:
    """Kiểm tra IFC theo IDS chuẩn buildingSMART (EIR/QD347 -> .ids -> ifctester).

    Tự sinh file .ids từ bộ yêu cầu EIR theo bộ môn + giai đoạn rồi chấm bằng
    ifctester; báo cả độ phủ (bao nhiêu cấu kiện thực sự được kiểm), không chỉ % đạt.

    Args:
        ifc_path:   Đường dẫn đầy đủ tới file IFC
        discipline: KT / KC / MEP / ALL
        stage:      TKCS / TKKT / TKTC / AS-BUILT
        project:    Tên dự án (tự nhận diện nếu để trống)
    """
    proj = project or Path(ifc_path).stem
    slug = "".join(c if c.isalnum() else "_" for c in proj)
    ids_path = OUTPUT / "_ids_cache" / f"EIR_{slug}_{discipline}_{stage}.ids"

    r1 = run(
        AUDIT / "eir_to_ids.py",
        ["--du-an", proj, "--discipline", discipline, "--stage", stage, "-o", str(ids_path)],
    )
    if not r1["ok"]:
        return err(r1["err"] or "Không sinh được file .ids từ EIR")

    r2 = run(
        AUDIT / "ids_run.py",
        ["--ids", str(ids_path), "--ifc", ifc_path, "--project", proj],
    )
    if not r2["ok"]:
        return err(r2["err"] or "Script thất bại")
    return ok({"console": r1["out"] + "\n" + r2["out"]})


@mcp.tool()
def kiem_tra_lod(
    ifc_path: str,
    stage: str = "TKKT",
    discipline: str = "ALL",
    project: str = "",
) -> str:
    """Kiểm tra LOD (Level of Development) 100–500 từng element theo stage và discipline.

    Args:
        ifc_path:   Đường dẫn đầy đủ tới file IFC
        stage:      TKCS / TKKT / TKTC / AS-BUILT
        discipline: KT / KC / MEP / ALL
        project:    Tên dự án (tự nhận diện nếu để trống)
    """
    r = run(
        AUDIT / "lod_checker.py",
        ["--ifc", ifc_path, "--stage", stage, "--discipline", discipline, "--project", project],
    )
    if not r["ok"]:
        return err(r["err"] or "Script thất bại")
    return ok({"console": r["out"]})


@mcp.tool()
def boc_khoi_luong(
    ifc_path: str,
    discipline: str = "ALL",
    stage: str = "TKKT",
    project: str = "",
    bo_sung_qto: bool = False,
) -> str:
    """Bóc khối lượng QTO từ IFC dựa trên BaseQuantities Psets. Xuất IFC_QTO_*.xlsx.

    IFC xuất từ Revit nên bật 'Export base quantities'. Nếu IFC không có
    BaseQuantities (ra 0 giá trị), đặt bo_sung_qto=True: đo khối lượng từ hình học,
    ghi bản sao <tên>_BoSungQTO.ifc vào thư mục kết quả rồi bóc trên bản đó.

    Args:
        ifc_path:    Đường dẫn đầy đủ tới file IFC
        discipline:  KT / KC / MEP / ALL
        stage:       TKCS / TKKT / TKTC / AS-BUILT
        project:     Tên dự án
        bo_sung_qto: True = đo từ hình học khi IFC thiếu BaseQuantities
    """
    args = ["--ifc", ifc_path, "--discipline", discipline, "--stage", stage, "--project", project]
    if bo_sung_qto:
        args.append("--bo-sung-qto")
    r = run(AUDIT / "ifc_qto.py", args)
    if not r["ok"]:
        return err(r["err"] or "Script thất bại")
    return ok({"console": r["out"]})


@mcp.tool()
def tong_hop_audit(
    project: str,
    discipline: str = "ALL",
    stage: str = "TKKT",
    top: int = 10,
) -> str:
    """Tổng hợp kết quả IDS + LOD + QTO từ thư mục kết quả để gửi cho lead discipline.
    Không cần file IFC — đọc từ các file audit đã chạy trước.

    Args:
        project:    Tên dự án (phải khớp với tên đã dùng khi chạy audit trước đó)
        discipline: KT / KC / MEP / ALL
        stage:      TKCS / TKKT / TKTC / AS-BUILT
        top:        Số lỗi ưu tiên hiển thị (mặc định: 10)
    """
    r = run(
        AUDIT / "audit_summary_report.py",
        ["--project", project, "--discipline", discipline, "--stage", stage, "--top", str(top)],
    )
    if not r["ok"]:
        return err(r["err"] or "Script thất bại — đảm bảo đã chạy kiem_tra_ids và kiem_tra_lod trước")
    return ok({"console": r["out"]})


# ── Entry point ──────────────────────────────────────────────────────────────

def main():
    mcp.run()


if __name__ == "__main__":
    main()
