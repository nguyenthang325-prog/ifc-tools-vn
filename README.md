# ifc-tools-vn

Bộ công cụ kiểm tra mô hình IFC theo cách làm Việt Nam: yêu cầu thông tin (IDS), mức độ phát triển (LOD), bóc khối lượng — chạy bằng dòng lệnh hoặc gọi thẳng trong Claude Desktop qua MCP.

*English summary at the end.*

## Vì sao có bộ này

Báo cáo kiểm IDS thường chỉ ghi "% đạt". Nhưng theo chuẩn IDS, một yêu cầu ở dạng mặc định (không khai báo bắt buộc phải có cấu kiện) mà **không áp dụng lên cấu kiện nào vẫn được tính là đạt**. Nên "100% đạt" có thể là 100% của phần chưa kiểm gì.

Bộ này chấm bằng `ifctester` (IfcOpenShell) theo chuẩn buildingSMART IDS 1.0 và báo thêm **độ phủ**: bao nhiêu cấu kiện trong mô hình thực sự có yêu cầu kiểm, yêu cầu nào không khớp cấu kiện nào. Muốn một yêu cầu bắt buộc phải có cấu kiện, dùng `eir_to_ids --bat-buoc`.

Ví dụ trên mô hình mẫu Duplex: 8/10 yêu cầu đạt, nhưng độ phủ chỉ **101/245 cấu kiện (41%)** — 144 cấu kiện không có yêu cầu nào áp dụng.

## Gồm những gì

| Lệnh MCP | Module dòng lệnh | Làm gì | Kết quả |
|---|---|---|---|
| `kiem_tra_ids` | `eir_to_ids` → `ids_run` | Sinh file `.ids` theo bộ môn + giai đoạn, chấm mô hình, đo độ phủ | `IDS_KetQua_*.xlsx`, `.html` (báo cáo gốc của ifctester), `.bcf` |
| `kiem_tra_lod` | `lod_checker` | LOD 100–500 từng cấu kiện theo giai đoạn | `LOD_Audit_*.xlsx` |
| `boc_khoi_luong` | `ifc_qto` | Bóc khối lượng từ BaseQuantities; tuỳ chọn đo từ hình học khi IFC không có | `IFC_QTO_*.xlsx` (+ `<tên>_BoSungQTO.ifc`) |
| `tong_hop_audit` | `audit_summary_report` | Gộp 3 kết quả trên thành 1 bảng gửi trưởng bộ môn | `AuditSummary_*.xlsx` |
| `kiem_tra_toan_dien` | `bim_audit_pipeline` | Chạy IDS → LOD → QTO trong 1 lệnh | `BIM_HealthDashboard_*.xlsx` |

Giai đoạn: `TKCS` · `TKKT` · `TKTC` · `AS-BUILT`. Bộ môn: `KT` · `KC` · `MEP` · `ALL`.

## Cài đặt

Cần Python 3.10 trở lên.

```bash
python -m venv .venv
.venv\Scripts\activate          # macOS/Linux: source .venv/bin/activate
pip install git+https://github.com/nguyenthang325-prog/ifc-tools-vn
```

Thư viện kéo theo: `ifcopenshell`, `ifctester` (LGPL, cài qua pip, không đóng gói trong repo), `openpyxl`, `numpy`, `mcp`.

## Dùng trong Claude Desktop

Mở `claude_desktop_config.json` (Claude Desktop → Settings → Developer → Edit Config), thêm:

```json
{
  "mcpServers": {
    "ifc-tools-vn": {
      "command": "C:\\duong\\dan\\toi\\.venv\\Scripts\\ifc-tools-vn.exe",
      "env": { "IFC_TOOLS_OUTPUT": "C:\\duong\\dan\\luu\\ket_qua" }
    }
  }
}
```

`command` là file `ifc-tools-vn` trong thư mục `Scripts` (Windows) hoặc `bin` (macOS/Linux) của venv vừa tạo. Bỏ `env` thì kết quả ghi vào `~/ifc_tools_output/`. Khởi động lại Claude Desktop rồi hỏi, ví dụ: *"Kiểm tra IDS file C:\...\Duplex_A_20110907.ifc, bộ môn KT, giai đoạn TKTC"*.

## Dùng bằng dòng lệnh

```bash
python -m ifc_tools_vn.checks.bim_audit_pipeline --ifc Duplex_A_20110907.ifc --project Duplex --stage TKTC --discipline KT

python -m ifc_tools_vn.checks.eir_to_ids --du-an Duplex --discipline KT --stage TKTC -o EIR_Duplex_KT_TKTC.ids
python -m ifc_tools_vn.checks.ids_run --ids EIR_Duplex_KT_TKTC.ids --ifc Duplex_A_20110907.ifc --project Duplex

python -m ifc_tools_vn.checks.lod_checker --ifc Duplex_A_20110907.ifc --stage TKTC --discipline KT
python -m ifc_tools_vn.checks.ifc_qto --ifc Duplex_A_20110907.ifc --discipline ALL --project Duplex --bo-sung-qto
python -m ifc_tools_vn.checks.audit_summary_report --project Duplex --discipline KT --stage TKTC
```

Thêm `-h` sau mỗi lệnh để xem đủ tuỳ chọn. Kết quả ghi vào `./ifc_tools_output/<ngày>/`, hoặc thư mục trong biến `IFC_TOOLS_OUTPUT`.

**Dùng file `.ids` của dự án mình:** gọi thẳng `ids_run --ids <file của bạn>.ids`. Lệnh MCP `kiem_tra_ids` luôn sinh `.ids` từ bộ yêu cầu mẫu bên dưới.

## Bộ yêu cầu mẫu

`eir_to_ids` sinh `.ids` từ một bộ 19 yêu cầu **do tác giả soạn**, tham khảo bảng thông tin theo giai đoạn của QĐ 347/QĐ-BXD và cách Revit xuất IFC2X3. Đây **không** phải yêu cầu chính thức của văn bản nào — ví dụ "tường phải khai FireRating từ TKTC" là yêu cầu tác giả đặt ra. Bảng đầy đủ và lý do từng lựa chọn: [`examples/EIR_mau/README_EIR.md`](examples/EIR_mau/README_EIR.md). Dự án thật nên dùng `.ids` theo EIR của chủ đầu tư.

## Thử với mô hình mẫu

Tải mô hình Duplex của buildingSMART (CC BY 4.0): [Duplex_A_20110907.ifc](https://github.com/buildingsmart-community/Community-Sample-Test-Files/tree/main/IFC%202.3.0.1%20(IFC%202x3)/Duplex%20Apartment). Chạy lệnh `bim_audit_pipeline` ở trên, kết quả đúng phải là:

| Kiểm | Kết quả |
|---|---|
| IDS (KT, TKTC) | 8/10 yêu cầu đạt · độ phủ 101/245 · 58 cấu kiện trượt (57 tường thiếu FireRating + công trình thiếu tên) |
| LOD (KT, TKTC) | 197 cấu kiện, 0% đạt LOD 350 (104 thiếu vật liệu, 90 thiếu FireRating, 3 không có hình học) |
| Khối lượng `--bo-sung-qto` | 216 cấu kiện |

Kiểm tự động: `DUPLEX_PATH=<đường dẫn file> python -m unittest discover tests -v` (cần bản clone repo).

## Giới hạn đã biết

- Đã chạy thử trên Windows 11 với Python 3.10 và 3.14, ifcopenshell 0.8.5 và 0.9.0. CI cấu hình chạy trên Windows và Ubuntu, Python 3.12.
- Viết cho IFC2X3 xuất từ Revit. Định dạng khác vẫn chạy nhưng ít được thử.
- `--bo-sung-qto` chưa đo được lan can, cầu thang gộp, nội thất, MEP. Thứ tự dòng trong bảng có thể đổi giữa các lần chạy (engine hình học chạy nhiều luồng); nội dung không đổi.
- Điểm IDS cao **không** có nghĩa mô hình dùng được để thanh toán: IDS không kiểm phân loại đúng/sai hay hình học.
- Chưa có công cụ báo cáo xung đột (Navisworks) — dự kiến bản sau.

## Giấy phép

MIT — xem [LICENSE](LICENSE). Tác giả: Nguyễn Đắc Thắng.

---

## English summary

**ifc-tools-vn** checks IFC models against Vietnamese practice: information requirements (buildingSMART IDS 1.0, checked with IfcOpenShell's `ifctester`), LOD per project stage, and quantity take-off — from the command line or as an MCP server for Claude Desktop.

Its main point is **coverage**: under the IDS standard, an optional (default) specification that matches no element still passes, so "100% compliant" can mean nothing was checked. Reports here state how many elements were actually covered (e.g. Duplex sample: 8/10 specs pass, but only 101/245 elements covered).

Install with `pip install git+https://github.com/nguyenthang325-prog/ifc-tools-vn`; run modules as `python -m ifc_tools_vn.checks.<name> -h`, or register the `ifc-tools-vn` executable as an MCP server (see `docs/claude_desktop_config.example.json`). The bundled requirement set is the author's own sample loosely based on Decision 347/QĐ-BXD, not an official specification. Output and messages are in Vietnamese. MIT licensed.
