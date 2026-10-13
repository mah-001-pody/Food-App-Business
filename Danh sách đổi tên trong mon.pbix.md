# Các tên cần đổi trong mon.pbix (làm bằng Power BI Desktop)

Báo cáo Word và các hình dashboard (Hình 4.1-4.5) đã dùng tên mới. Đổi tên measure trong Power BI Desktop sẽ tự cập nhật các visual đang dùng measure đó.

| Vị trí | Tên hiện tại | Tên mới |
|---|---|---|
| Tên trang 4 + textbox tiêu đề | 4. Hiệu quả chiến dịch / Hiệu quả chiến dịch | 4. Phản hồi chiến dịch / Phản hồi chiến dịch |
| Phụ đề trang 4 | Tỷ lệ chấp nhận, phản hồi và độ phủ chiến dịch marketing | Tỷ lệ chấp nhận và đặc điểm khách chấp nhận chiến dịch |
| Measure | Thu nhập TB người nhận CD | Thu nhập TB khách chấp nhận CD |
| Measure | Số con TB người nhận CD | Số con TB khách chấp nhận CD |
| Measure | Chi tiêu TB người nhận CD | Chi tiêu TB khách chấp nhận CD |
| Tiêu đề visual trang 4 | Thu nhập TB của người nhận từng chiến dịch | Thu nhập TB của khách chấp nhận từng chiến dịch |
| Tên hiển thị trục / tiêu đề visual trang 4 | Số chiến dịch đã nhận | Số chiến dịch đã chấp nhận |
| Measure | Doanh thu giữ lại được | Chi tiêu lịch sử theo kịch bản win-back |
| Measure | % tổng doanh thu giữ lại | % tổng chi tiêu theo kịch bản win-back |
| Tiêu đề slicer trang 5 | Giả định tỷ lệ win-back (mặc định 10%) | Giả định tỷ lệ win-back (khi chưa chọn: 10%) |
| Dim_Segment (Power Query) + giá trị RFM_Segment trong Customers | New / Promising | Promising (sửa cả Recommended_Action: bỏ "kiểu Chiến dịch 3") |
| Nhãn Marital_Status | Married/Partner | Married |

Sau khi đổi, nên chụp lại 5 trang dashboard và Model view thật từ Power BI Desktop để thay Hình 3.1 và 4.1-4.5 (hiện là hình dựng lại).
