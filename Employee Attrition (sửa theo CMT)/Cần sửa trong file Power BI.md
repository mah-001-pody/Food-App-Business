# Các chỗ cần sửa trong file Power BI (Employee Attrition) để khớp báo cáo

Làm trong Power BI Desktop.

## 1. Tham số chi phí thay thế (bắt buộc: để dashboard ra được 2,7 triệu USD)

1. Chọn bảng **Replacement Cost** → sửa biểu thức bảng thành:
   `Replacement Cost = GENERATESERIES(0.1, 1.5, 0.05)`
2. Sửa measure tỷ lệ chi phí thay thế (measure đang đọc giá trị từ slicer) để mặc định là 20%:
   `Tỷ lệ chi phí thay thế = SELECTEDVALUE('Replacement Cost'[Replacement Cost], 0.2)`
3. Trên slicer tham số ở trang 6: bỏ chọn 50% (để trống thì measure dùng 20%) hoặc chọn sẵn 20% rồi lưu file.
4. Kiểm tra 5 thẻ KPI trang 6: "Chi phí thay thế ước tính" phải ra khoảng **2,7 triệu USD** (13,6 triệu × 20%). Nếu tiêu đề thẻ hay ô chữ có ghi "50%" thì đổi thành "20% (mức cơ sở)".

## 2. Tiêu đề trang (khớp Bảng 4.1 của báo cáo)

| Trang | Tiêu đề cũ | Tiêu đề mới |
|---|---|---|
| 1 | Cứ 6 nhân viên có 1 người rời đi - từ tuyến đầu | Cứ 6 nhân viên có 1 người rời đi - tập trung ở các vị trí cấp thấp |
| 3 | Làm thêm giờ là yếu tố nguy hiểm nhất | Làm thêm giờ gắn với chênh lệch tỷ lệ nghỉ lớn nhất |
| 5 | Đổi quản lý làm tỷ lệ nghỉ tăng gần gấp đôi | Vừa đổi quản lý: tỷ lệ nghỉ 23,0% so với 12,4% |

## 3. Nhãn nhóm thâm niên

Trong Power Query (cột Tenure_Group và bảng Dim_Tenure): đổi nhãn **"20+"** thành **"21+"**, vì nhóm "11-20" đã gồm cả 20 năm.

## 4. Không bắt buộc

Trên dashboard, nhãn thang cân bằng CV-CS đang để Bad/Good/Better/Best (theo mô tả của bộ dữ liệu). Báo cáo giờ ghi "Mức 1-4"; có thể đổi trên dashboard cho thống nhất.
