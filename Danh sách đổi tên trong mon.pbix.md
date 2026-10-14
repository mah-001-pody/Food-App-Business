# Các chỗ cần sửa trong mon.pbix (Food App) để khớp báo cáo

Làm trong Power BI Desktop. Khi đổi tên measure, Power BI tự cập nhật các visual đang dùng measure đó; tooltip và tiêu đề visual gõ tay thì phải sửa riêng.

## 1. Sửa measure thu nhập khách chấp nhận (tính theo khách, không theo lượt)

Đổi tên `Thu nhập TB người nhận CD` thành `Thu nhập TB khách chấp nhận CD` và thay công thức:

```
Thu nhập TB khách chấp nhận CD =
VAR _kh = FILTER(VALUES(Customers[CustomerID]), CALCULATE(MAX(Campaign_Response[Accepted])) = 1)
RETURN CALCULATE(AVERAGE(Customers[Income]), _kh)
```

Kiểm tra: biểu đồ theo chiến dịch vẫn ra 78,8K / 71,5K / 51,0K / 68,2K / 82,4K. Đặt measure vào một thẻ không lọc thì phải ra 65.183 (công thức cũ ra 69.868).
Làm tương tự với `Số con TB người nhận CD` và `Chi tiêu TB người nhận CD` (thay `AVERAGE(Customers[Income])` bằng `AVERAGE(Customers[Total_Children])` và `AVERAGE(Customers[Total_Spend])`).

## 2. Đổi tên measure, nhãn, tiêu đề

| Vị trí | Tên hiện tại | Tên mới |
|---|---|---|
| Tên trang 4 + textbox tiêu đề | 4. Hiệu quả chiến dịch / Hiệu quả chiến dịch | 4. Phản hồi chiến dịch / Phản hồi chiến dịch |
| Phụ đề trang 4 | Tỷ lệ chấp nhận, phản hồi và độ phủ chiến dịch marketing | Tỷ lệ chấp nhận và đặc điểm khách chấp nhận chiến dịch |
| Tiêu đề visual trang 4 | Thu nhập TB của người nhận từng chiến dịch | Thu nhập TB của khách chấp nhận từng chiến dịch |
| Nhãn trục / tiêu đề visual trang 4 | Số chiến dịch đã nhận | Số chiến dịch đã chấp nhận |
| Measure | Doanh thu rủi ro (At Risk) | Chi tiêu lịch sử nhóm At Risk |
| Measure | % doanh thu rủi ro | % chi tiêu lịch sử nhóm At Risk |
| Measure | Doanh thu giữ lại được | Chi tiêu lịch sử theo kịch bản win-back |
| Measure | % tổng doanh thu giữ lại | % tổng chi tiêu theo kịch bản win-back |
| Tiêu đề slicer trang 5 | Giả định tỷ lệ win-back (mặc định 10%) | Giả định tỷ lệ win-back (khi chưa chọn: 10%) |
| Dim_Segment (Power Query) + giá trị RFM_Segment trong Customers | New / Promising | Promising (sửa cả Recommended_Action: bỏ "kiểu Chiến dịch 3") |
| Nhãn Marital_Status | Married/Partner | Married |

## 3. Kiểm tra trên giao diện

- **Slicer phân khúc RFM:** chọn Champions. Các thẻ At Risk phải để trống, không giữ số At Risk; slicer lọc qua Dim_Segment nên measure hiện tại đã đúng. Có thể ghi chú dưới thẻ: "chỉ có giá trị khi chọn Tất cả hoặc At Risk".
- **Slicer win-back:** để trống (không chọn sẵn). Thẻ win-back phải ra 44,6K, là giá trị dự phòng 10%.
- **Đối chiếu số liệu:** thử vài tổ hợp slicer và so với file Excel sạch (05_FoodAppBusiness_Clean.xlsx): số khách, tổng chi tiêu, AOV.

## 4. Xuất hình thật thay cho hình dựng lại

Sau khi sửa, xuất 5 trang dashboard (File → Export → PDF, hoặc chụp màn hình ở độ phân giải cao) và chụp Model view. Thay vào Hình 3.1 và Hình 4.1–4.5 trong Word, rồi sửa dòng nguồn thành "Nguồn: Nhóm tác giả, xuất từ file mon.pbix".
