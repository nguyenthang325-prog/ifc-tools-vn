# EIR — Yêu cầu trao đổi thông tin (bản người đọc)

Bản này là **mặt người đọc** của các file `.ids` trong cùng thư mục (mặt máy đọc). Hai bản
sinh từ cùng một nguồn (`eir_to_ids.py`) nên luôn khớp nhau. Đưa bản này vào hồ sơ mời thầu /
phụ lục hợp đồng; đưa file `.ids` cho bên thi công tự kiểm trước khi nộp.

- **Nguồn:** QĐ 347/QĐ-BXD (bảng thông tin theo giai đoạn) + thực tế Revit xuất IFC2X3.
- **Chuẩn:** buildingSMART IDS 1.0 · engine chấm: `ifctester` (IfcOpenShell).
- **Nguyên tắc:** mọi yêu cầu là **bắt buộc** ở giai đoạn ghi trong cột "Từ giai đoạn". Không có
  mức "khuyến nghị"; cần nới thì dời sang giai đoạn sau.

## Giai đoạn

`TKCS` (thiết kế cơ sở) → `TKKT` (kỹ thuật) → `TKTC` (thi công) → `AS-BUILT` (hoàn công). File
`.ids` giai đoạn sau bao gồm toàn bộ yêu cầu của giai đoạn trước.

## Bảng yêu cầu

| Mã | Cấu kiện (IFC) | Yêu cầu thông tin | Bộ môn | Từ giai đoạn |
|---|---|---|---|---|
| EIR-PRJ-01 | IfcProject | Có Tên dự án | Chung | TKCS |
| EIR-PRJ-02 | IfcBuilding | Có Tên công trình | Chung | TKCS |
| EIR-PRJ-03 | IfcSite | Có Tên địa điểm | Chung | TKKT |
| EIR-PRJ-04 | IfcBuildingStorey | Tầng có Tên | Chung | TKCS |
| EIR-KT-01 | IfcWall(+StandardCase) | Tường có Tên | KT | TKKT |
| EIR-KT-02 | IfcWall(+StandardCase) | Tường gán Vật liệu | KT | TKKT |
| EIR-KT-03 | IfcWall(+StandardCase) | Khai IsExternal (trong/ngoài) | KT | TKKT |
| EIR-KT-04 | IfcWall(+StandardCase) | Khai FireRating (chống cháy) | KT | TKTC |
| EIR-KT-05 | IfcDoor | Cửa đi khai FireRating | KT | TKTC |
| EIR-KT-06 | IfcWindow | Cửa sổ khai IsExternal | KT | TKKT |
| EIR-KC-01 | IfcBeam | Dầm có Tên | KC | TKKT |
| EIR-KC-02 | IfcBeam | Dầm gán Vật liệu | KC | TKKT |
| EIR-KC-03 | IfcBeam | Khai LoadBearing (chịu lực) | KC | TKTC |
| EIR-KC-04 | IfcColumn | Cột gán Vật liệu | KC | TKKT |
| EIR-KC-05 | IfcColumn | Khai LoadBearing | KC | TKTC |
| EIR-KC-06 | IfcSlab | Sàn gán Vật liệu | KC | TKKT |
| EIR-KC-07 | IfcFooting | Móng gán Vật liệu | KC | TKKT |
| EIR-MEP-01 | IfcPipeSegment | Ống có Tên | MEP | TKKT |
| EIR-MEP-02 | IfcDuctSegment | Ống gió có Tên | MEP | TKKT |

## Cố ý KHÔNG đưa vào (tránh báo động giả)

- **FireRating trên dầm/cột/sàn kết cấu** — FireRating là thuộc tính tường/cửa/sổ; đòi trên kết cấu
  từng khiến mọi mô hình KC bị chấm LOD ~0%.
- **Cốt thép / khối lượng** — thuộc khâu bóc QTO, không phải kiểm tra thông tin (IDS).
- **Kiểm "vật liệu ở instance"** — Revit gán vật liệu ở Type; engine chuẩn tự phân giải qua Type.

## Điều IDS KHÔNG kiểm được

Phân loại đúng/sai (cọc bị mô hình thành IfcSlab), cấu trúc tầng, hình học — những cái đó thuộc khâu
audit phân loại / QTO. Điểm IDS cao **không** đồng nghĩa mô hình dùng được cho thanh toán.
