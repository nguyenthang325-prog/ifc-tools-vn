# Thay đổi

## 0.2.0 — 2026-11-03

Sửa cách bóc khối lượng (`ifc_qto`). **Số trong `Tom_Tat_BoQ` sẽ đổi**, có nhóm đổi mạnh. Báo cáo QTO hoặc HealthDashboard đã xuất bằng 0.1.0 từ IFC Revit nên xuất lại.

Lỗi tìm thấy khi đối chiếu 7 mô hình Revit thật với thể tích đo từ hình học:

- **L1** Bảng tóm tắt cộng theo *tên* đại lượng, nên cấu kiện thiếu đúng đại lượng được chọn (có `NetVolume`, không có `GrossVolume`) bị rơi khỏi tổng. Có mô hình tường chỉ còn 3,8 m³ thay vì 56,6 m³. Nay mỗi cấu kiện lấy đúng 1 khối lượng chính rồi mới cộng.
- **L2** `GrossVolume` và `NetVolume` cùng điểm ưu tiên, nên số được chọn tuỳ thứ tự thuộc tính trong file. Có mô hình sàn hụt 63 %, có mô hình dư 56 %.
- **L4** Revit có lúc xuất `NetVolume`/`GrossVolume` sai: tấm con của mái (IFC4, Revit 2027) và sàn chia Part. Không quy tắc cố định nào đúng cho mọi file, nên nay dùng **trọng tài hình học**: Gross ≈ Net (≤ 0,5 %) thì lấy Net; nếu lệch thì đo thể tích hình học và lấy số khớp (≤ 2 %); không số nào khớp thì lấy hình học. Kết quả có gắn cờ.
- **L3** Tấm con (`IfcSlab`) của `IfcRoof` bị cộng vào "Sàn/Mái". Nay vào nhóm **Mái**. Diện tích mái lấy từ `IfcRoof`: `GrossArea` (mặt nghiêng) và `ProjectedArea` (chiếu bằng).
- **L5** Vế thang `IfcStairFlight` không được bóc, còn chiếu nghỉ bị cộng vào Sàn. Nay cả hai vào nhóm **Cầu thang**. IFC2X3 Revit không xuất thể tích vế thang, nên gắn cờ để bóc từ Revit.
- **L6** `Area` của cửa đi/cửa sổ trong IFC4 Revit 2027 không phải diện tích ô cửa (cửa 900 × 2000 ra 4,93 m²). Nay dùng **số cái + W × H**.

Thay đổi đầu ra:

- `Tom_Tat_BoQ`: mỗi nhóm có thể có nhiều dòng (theo đại lượng và đơn vị). 6 cột đầu giữ nguyên vị trí, thêm cột 7 "Nguồn / cờ".
- Sheet mới `KL_Chinh`: khối lượng chính của từng cấu kiện, nguồn và cờ. Cấu kiện có cờ xếp đầu.
- `Chi_Tiet` thêm cột "Nhóm theo cha". `Tong_Hop` ghi rõ là số thô, không dùng làm BoQ.
- Tuỳ chọn mới `--khong-kiem-hinh-hoc`: bỏ bước đo hình học, để chạy nhanh với file rất lớn. File 15 tầng (16 MB) mất khoảng 35 s thay vì 8 s.
- Duplex: `Tổng elements` từ 216 lên 218 (2 vế thang).
- Module mới `checks/qto_chon.py`, kèm test `tests/test_qto_chon.py` dựng IFC trong bộ nhớ (chạy trên CI, không cần file mẫu).

## 0.1.0 — 2026-10-22

Bản đầu tiên.

- 5 lệnh MCP: `kiem_tra_toan_dien`, `kiem_tra_ids`, `kiem_tra_lod`, `boc_khoi_luong`, `tong_hop_audit`.
- Kiểm IDS bằng `ifctester` kèm **độ phủ**: báo bao nhiêu cấu kiện thực sự được kiểm, không chỉ % đạt.
- Xuất BCF mở được trong BIMcollab Zoom.
- `ifc_qto --bo-sung-qto`: đo khối lượng từ hình học khi IFC không có BaseQuantities, ghi vào bản sao.
- Đường dẫn kết quả cấu hình bằng biến `IFC_TOOLS_OUTPUT`.
