# Bước 1: Tiền xử lý và làm sạch dữ liệu (tiêu chí 2, 15%)

> Tài liệu này dùng để viết phần **"Data Pre-processing & Cleaning"** trong báo cáo cuối kỳ VJP205.
> Code: `scripts/01_data_cleaning.py`. Dữ liệu sạch: `data/processed/FoodApp_Clean.xlsx` và `.csv`.

## 1. Đánh giá tổng quan chất lượng dữ liệu gốc

| Hạng mục | Kết quả | Đánh giá |
|---|---|---|
| Kích thước | 2.205 dòng × 26 cột, toàn bộ là số nguyên | Đủ lớn cho thống kê, phân cụm và hồi quy logistic |
| Giá trị thiếu | 0 | Bộ dữ liệu đã được xử lý sơ bộ (đây là phiên bản rút gọn của bộ iFood Marketing) |
| Mã khách hàng | **Không có** | Không truy vết được từng khách, không phân biệt được bản ghi trùng thật hay trùng giả → phải tạo `CustomerID` |
| Bản ghi trùng lặp | **204 dòng (9,3%)** trùng hoàn toàn 26/26 cột | Vấn đề chất lượng **lớn nhất**. Nếu giữ lại, các chỉ số sẽ bị thổi phồng và mô hình bị lệch |
| Mâu thuẫn logic | 9 dòng | Lỗi nhập liệu (chi tiết ở mục 2) |
| Biến dư thừa | `Single` = 1 − `Married`; `Total_Campaigns` = tổng Cmp1…Cmp5 | Gây đa cộng tuyến, có thể rò rỉ dữ liệu (leakage) khi mô hình hoá |
| Phân phối | Chi tiêu lệch phải mạnh (skew 1,2 → 2,1) | Có nhiều ngoại lai nhưng **đó là hành vi thật** |
| Mất cân bằng nhãn | Complain chỉ 0,95%; Cmp2 chỉ 1,3% | Cần lưu ý khi kiểm định và xây mô hình |

**Lưu ý về ngữ nghĩa:** cột `Monthly Income` có giá trị từ 1.730 đến 113.734, trung vị khoảng 51.000. Mức này giống **thu nhập năm của hộ gia đình** trong bộ iFood gốc hơn là thu nhập tháng. Báo cáo nên nêu rõ giả định này. Mình đặt tên trung tính là `Income`.

## 2. Các bước làm sạch (Cleaning Log)

| Bước | Quy tắc | Số dòng | Lý do theo góc nhìn kinh doanh |
|---|---|---|---|
| R1 | Kiểm tra missing, miền giá trị nhị phân, `Married + Single = 1`, `Total = ΣCmp`, giá trị âm | 0 lỗi | Xác nhận cấu trúc dữ liệu nhất quán |
| R2 | **Loại bản ghi trùng hoàn toàn** (giữ bản đầu tiên) | −204 | Hai khách khác nhau gần như không thể trùng cả 26 biến liên tục (thu nhập, số ngày hoạt động, chi tiêu 5 nhóm hàng…) |
| R3 | Tổng số lần mua = 0 nhưng chi tiêu > 0 | −4 | Không thể chi tiền mà không mua lần nào |
| R4 | Số lần mua có giảm giá > tổng số lần mua | −1 | Tập con không thể lớn hơn tập cha (15 lần giảm giá / 1 lần mua) |
| R5 | Tổng chi tiêu > 50% thu nhập | −1 | Thu nhập 2.447 nhưng chi 1.729. Bản ghi cao kế tiếp chỉ khoảng 3%, khoảng cách quá lớn nên nghi sai nhập liệu thu nhập |
| R6 | > 20 lần mua qua web, ≤ 1 lượt truy cập/tháng và chi tiêu rất thấp | −3 | 23–27 đơn web nhưng gần như không truy cập web, chi tiêu chỉ 38–274 → nghi bot hoặc lỗi hệ thống |
| R7 | Ngoại lai IQR ở chi tiêu và kênh mua | **Giữ lại** (488 khách) | Đây là nhóm khách chi tiêu cao, chính là nhóm doanh nghiệp quan tâm nhất. Không xoá, không cắt ngưỡng (cap); khi chạy mô hình thì dùng `log(1+x)` |
| R8 | Bỏ cột `Single` | −1 cột | Tránh bẫy biến giả (dummy trap) và đa cộng tuyến |
| R9 | Feature engineering | +33 biến | Xem mục 3 |
| R10 | Kiểm tra cuối (assert) | ✓ | 0 missing, 0 trùng lặp, 0 mâu thuẫn logic |

**Kết quả: từ 2.205 dòng còn 1.992 dòng (giữ lại 90,3%); từ 26 cột thành 60 cột.**
Toàn bộ dòng bị loại được lưu ở sheet `Removed_Rows`, kèm quy tắc, lý do và số dòng gốc, để có thể kiểm tra lại.

> **Vì sao ngoại lai không bị xoá?** Đây là lỗi phổ biến của sinh viên. Theo nguyên tắc Pareto, **top 20% khách hàng tạo ra 53,4% tổng chi tiêu**. Xoá ngoại lai đồng nghĩa với xoá chính nhóm khách giá trị nhất. Chỉ những gì **mâu thuẫn logic** mới bị loại; những gì chỉ **cực đoan về thống kê** thì vẫn giữ.

## 3. Feature engineering (33 biến mới)

| Nhóm | Biến | Mục đích phân tích |
|---|---|---|
| Nhân khẩu học | `Age_Group`, `Income_Group`, `Marital_Status`, `Education`, `Total_Children`, `Has_Children`, `Family_Size`, `Tenure_Years` | Vẽ chân dung khách hàng, dùng làm slicer trong Power BI |
| Chi tiêu | `Total_Spend`, `Share_Wines…Share_Sweets`, `Spend_to_Income_Pct`, `Value_Tier` | Đo giá trị khách hàng, cơ cấu giỏ hàng |
| Hành vi mua | `Total_Purchases`, `AOV`, `Deal_Ratio`, `Share_Web/Catalog/Store`, `Preferred_Channel`, `Recency_Group` | Phân tích kênh bán và độ nhạy cảm với giá |
| Marketing | `Campaign_Responder` | Biến mục tiêu cho hồi quy logistic |
| RFM | `R_Score`, `F_Score`, `M_Score`, `FM_Score`, `RFM_Score`, `RFM_Segment` | Phân khúc khách hàng, xác định nhóm có nguy cơ rời bỏ |
| Mô hình | `Log_Income`, `Log_Total_Spend` | Giảm lệch phải cho hồi quy |

**Quy tắc phân khúc RFM**: chấm điểm R/F/M từ 1–5 theo ngũ phân vị (quintile); điểm R đảo chiều, mua càng gần đây điểm càng cao.

| Phân khúc | Điều kiện | Số khách |
|---|---|---|
| Champions | R ≥ 4 và FM ≥ 4 | 290 |
| Loyal Customers | R ≥ 3 và FM ≥ 3 | 417 |
| New / Promising | R ≥ 4 và FM < 3 | 321 |
| Need Attention | R = 3 và FM < 3 | 167 |
| **At Risk** | R ≤ 2 và FM ≥ 3 | **480** |
| Hibernating | còn lại | 317 |

## 4. Một số tín hiệu ban đầu (gợi ý giả thuyết cho bước sau)

- **Rượu vang chiếm 54,4%** và **thịt chiếm 29,4%** tổng chi tiêu. Doanh thu phụ thuộc lớn vào 2 nhóm hàng có biên lợi nhuận cao.
- **Cửa hàng là kênh chính** của 66% khách. Kênh web chiếm 27%, catalog chỉ 7%.
- **Chỉ 21% khách** nhận ít nhất 1 chiến dịch. Chiến dịch 2 gần như thất bại (1,3%).
- Tương quan với tổng chi tiêu:
  - Thu nhập: **+0,83**
  - Tỷ lệ mua hàng giảm giá: **−0,63**
  - Số lượt truy cập web: **−0,50**
  - Số con: **−0,32**

  → Gợi ý giả thuyết: *nhóm khách săn giảm giá, có con nhỏ thì chi ít hơn; truy cập web nhiều nhưng không chuyển đổi thành đơn*.
- **480 khách "At Risk"** (24%) là nhóm từng chi tiêu tốt nhưng lâu không quay lại. Đây là cơ hội cho chiến dịch win-back (kéo khách quay lại).

## 5. Cấu trúc file đầu ra `FoodApp_Clean.xlsx`

| Sheet | Nội dung |
|---|---|
| `Data_Clean` | 1.992 khách × 60 cột, dùng để nạp vào Power BI và mô hình |
| `Data_Dictionary` | Mô tả từng biến, kiểu dữ liệu, biến gốc hay biến tạo mới |
| `Cleaning_Log` | Nhật ký từng bước, số dòng bị ảnh hưởng, lý do |
| `Removed_Rows` | 213 dòng bị loại, kèm quy tắc tương ứng |
| `Outlier_Check` | Ngưỡng IQR, số ngoại lai và độ lệch của từng biến |
| `Quality_Before_After` | Thống kê mô tả trước và sau khi làm sạch |
