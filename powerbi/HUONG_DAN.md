# Hướng dẫn mở dashboard Power BI (`.pbip`) và lưu thành `.pbix`

## 0. Yêu cầu
- **Windows** có cài **Power BI Desktop bản mới** (2025 trở lên). Cài từ Microsoft Store để luôn được cập nhật.
- Nếu Power BI báo không mở được `.pbip`, vào **File → Options and settings → Options → Preview features**, bật **Power BI Project (.pbip) save option** và **Store reports using enhanced metadata format (PBIR)**, rồi khởi động lại Power BI.

## 1. Chuẩn bị thư mục (cách nhanh nhất)
Giải nén `FoodApp_Dashboard.zip` vào **ổ C:\\**. Sau khi giải nén, bạn sẽ có:

```
C:\FoodApp\
├── FoodApp_Clean.xlsx               ← dữ liệu đã làm sạch
├── FoodApp_Dashboard.pbip           ← BẤM ĐÚP FILE NÀY
├── FoodApp_Dashboard.SemanticModel\ ← mô hình dữ liệu (Power Query + DAX)
└── FoodApp_Dashboard.Report\        ← 7 trang báo cáo
```

> Đường dẫn mặc định của tham số `DataFile` là `C:\FoodApp\FoodApp_Clean.xlsx`. Giải nén đúng chỗ như trên thì **không cần sửa gì**.

## 2. Mở và nạp dữ liệu
1. Bấm đúp **`FoodApp_Dashboard.pbip`**.
2. Lần đầu mở, các biểu đồ sẽ **trống**. Đây là bình thường: file `.pbip` chỉ chứa cấu trúc, chưa chứa dữ liệu.
3. Bấm **Home → Refresh**. Chờ khoảng 10–30 giây.
4. *(Chỉ khi bạn để file Excel ở chỗ khác)* Vào **Home → Transform data ▾ → Edit parameters**, sửa `DataFile` thành đường dẫn đầy đủ tới `FoodApp_Clean.xlsx`, bấm OK rồi **Apply changes**.

## 3. Lưu thành file `.pbix` để nộp bài
**File → Save as**, chọn kiểu **Power BI file (\*.pbix)**, đặt tên theo quy định, ví dụ `MaLopBA_SoNhom.pbix`.

## 4. Kiểm tra nhanh: các con số phải khớp

| Trang | Visual | Giá trị đúng |
|---|---|---|
| 1 | Số khách hàng | **1.992** |
| 1 | Tổng chi tiêu | **1.125.815** |
| 1 | Chi tiêu TB/khách | **565** |
| 1 | AOV (giá trị TB/đơn) | **44,9** |
| 1 | Tỷ lệ phản hồi CD | **21,0%** |
| 1 | Top 20% khách chiếm % doanh thu | **53,4%** |
| 1 | Pareto, cột "Top 10%" | **31,0%** |
| 4 | Tỷ lệ chấp nhận CD1 → CD5 | 6,6% · **1,3%** · 7,4% · 7,6% · 7,3% |
| 5 | Doanh thu rủi ro (At Risk) | **446.346** (39,6%) |
| 5 | Khách VIP đang At Risk | **166** |
| 5 | Doanh thu giữ lại được (mặc định 10%) | **44.635** |
| 6 | Dòng gốc → bị loại → sạch | **2.205 → 213 → 1.992** (giữ lại 90,3%) |

Số khớp tức là mô hình và DAX đang chạy đúng.

## 5. Cấu trúc dashboard (7 trang)

| Trang | Câu hỏi kinh doanh | Visual chính |
|---|---|---|
| 1. Tổng quan | Doanh thu tập trung ở đâu? | 6 thẻ KPI, Pareto thập phân vị, donut nhóm hàng, treemap RFM |
| 2. Chân dung khách hàng | Ai chi tiêu nhiều, ai chi tiêu ít? | Chi tiêu theo thu nhập/số con/tuổi, scatter thu nhập × chi tiêu |
| 3. Sản phẩm & kênh | Mua gì, mua qua đâu? | Cơ cấu giỏ hàng 100%, AOV theo kênh, bảng truy cập web |
| 4. Hiệu quả chiến dịch | Chiến dịch nào hiệu quả, với ai? | Tỷ lệ chấp nhận, thu nhập người nhận, phủ sóng theo nhóm giá trị |
| 5. RFM & giữ chân | Ai sắp rời đi, đáng bao nhiêu tiền? | KPI rủi ro, slicer what-if win-back, bubble chart phân khúc, bảng hành động |
| 6. Chất lượng dữ liệu | Vì sao các con số đáng tin? | Phễu dữ liệu, Cleaning Log, giải thích về ngoại lai |
| Chi tiết khách hàng *(ẩn)* | Danh sách khách theo phân khúc | Trang drill-through: ở trang 5, chuột phải vào một phân khúc → **Drill through** |

- **Slicer đồng bộ** giữa trang 1–5: nhóm thu nhập, nhóm tuổi, số con, phân khúc RFM.
- **Mỗi trang có ô "SO WHAT?"** ở cuối, ghi kết luận và hành động đề xuất.

## 6. Mô hình dữ liệu (star schema)

```
Dim_IncomeGroup ─┐
Dim_AgeGroup ────┼──► Customers (1 dòng/khách) ◄── Spend_By_Category    (unpivot 5 nhóm hàng)
Dim_Segment ─────┘          ▲                 ◄── Purchases_By_Channel (unpivot 3 kênh)
                            └─────────────────── Campaign_Response    (unpivot 5 chiến dịch)

Bảng độc lập: _Measures (40 measure DAX) · Winback Rate (what-if) · Cleaning_Log · Data_Funnel
```

- **Power Query** (Home → Transform data): đọc Excel, ép kiểu, tạo nhãn (số con, nhóm truy cập web, săn giảm giá), xếp hạng chi tiêu theo thập phân vị, unpivot các bảng nhóm hàng, kênh, chiến dịch. Đây là bằng chứng cho phần tiền xử lý dữ liệu (tiêu chí 2).
- **DAX**: 40 measure trong bảng `_Measures`, chia thành 6 thư mục theo đúng 6 trang.

## 7. Nếu gặp lỗi
| Hiện tượng | Cách xử lý |
|---|---|
| Không mở được `.pbip` / báo lỗi định dạng | Cập nhật Power BI Desktop; bật Preview features ở mục 0 |
| Refresh báo *"Could not find file"* | Sửa tham số `DataFile` (mục 2, bước 4) |
| Một biểu đồ báo lỗi hoặc hiển thị lệch | Chụp màn hình gửi lại để mình sửa; hoặc xoá visual đó rồi kéo lại trường tương ứng |
| Ô so-what bị tràn chữ | Kéo giãn ô, hoặc giảm cỡ chữ (chọn ô → Format) |

## 8. Tạo lại toàn bộ từ đầu (tuỳ chọn)
```bash
python scripts/01_data_cleaning.py   # làm sạch → data/processed/FoodApp_Clean.xlsx
python scripts/02_build_pbip.py      # sinh lại thư mục powerbi/
```
