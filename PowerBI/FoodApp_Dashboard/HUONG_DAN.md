# Food App Dashboard - bản PBIP (theme "Fresh Market")

## Cách mở trên Windows

1. Giải nén cả thư mục `FoodApp_Dashboard` (giữ nguyên cấu trúc, **không** tách file `.pbip` ra riêng).
2. Mở bằng **Power BI Desktop bản mới nhất** (bản Microsoft Store hoặc bản tải về từ tháng 7/2026 trở đi):
   `File → Open → FoodApp_Dashboard.pbip`.
3. Lần mở đầu tiên mô hình chưa có dữ liệu → bấm **Home → Refresh**.
   Toàn bộ dữ liệu đã được nhúng sẵn trong Power Query, nên **không cần file Excel và không cần mạng**.
4. Muốn nộp file `.pbix`: `File → Save as → Power BI file (.pbix)`.

## Cấu trúc thư mục

```
FoodApp_Dashboard/
├── FoodApp_Dashboard.pbip              ← mở file này
├── FoodApp_Dashboard.Report/           ← báo cáo (định dạng PBIR)
│   ├── definition/pages/...            ← 7 trang, mỗi visual 1 file JSON
│   └── StaticResources/RegisteredResources/FoodAppTheme.json   ← theme
└── FoodApp_Dashboard.SemanticModel/    ← mô hình dữ liệu (TMDL)
    └── definition/tables/*.tmdl        ← 11 bảng, 40 measure DAX, 6 quan hệ
```

## Thay đổi so với bản gốc

**Bảng màu "Fresh Market"** - tông ấm, hợp với chủ đề thực phẩm / ăn uống:

| Vai trò | Màu | Mã |
|---|---|---|
| Màu chính (cột, KPI, tiêu đề) | Xanh húng quế | `#2E6B4F` / `#1F4D3A` |
| Màu nhấn (điểm cần chú ý, đường lũy kế) | Đỏ cà chua | `#E4572E` |
| Phụ 1 | Vàng nghệ | `#F2A541` |
| Phụ 2 | Đỏ rượu vang | `#7B2D43` |
| Phụ 3 | Xanh biển | `#3D7EA6` |
| Nền trang | Kem | `#F7F3EC` |
| Chữ | Nâu espresso | `#2B2118` |

**Mỗi nhóm luôn giữ một màu cố định trên mọi trang:**
- Nhóm hàng: Rượu vang = đỏ vang · Thịt = đỏ cà chua · Cá = xanh biển · Trái cây = xanh lá · Đồ ngọt = vàng nghệ
- Phân khúc RFM: Champions = xanh húng quế · Loyal = xanh lá · New/Promising = xanh biển · Need Attention = vàng · **At Risk = đỏ cà chua** · Hibernating = xám
- Kênh: Store = xanh húng quế · Web = xanh biển · Catalog = vàng nghệ
- Biểu đồ "Tỷ lệ chấp nhận theo chiến dịch": **Chiến dịch 2 (thất bại) tô đỏ**, các chiến dịch còn lại tô xanh.

**Phông chữ:** Segoe UI (tiêu đề dùng Segoe UI Semibold), số KPI dùng DIN. Segoe UI hiển thị tiếng Việt có dấu đầy đủ trên Windows.

**Trình bày:**
- Thẻ visual bo góc 10px, viền màu be nhạt, đổ bóng nhẹ.
- Header bảng nền xanh, chữ trắng, các dòng xen kẽ màu kem.
- Mỗi trang có thanh nhấn màu đỏ cà chua bên trái tiêu đề.
- Hộp "SO WHAT?" nền kem, viền vàng nghệ.
- Căn lưới đều 10px: 4 slicer thẳng hàng với lề phải, các hàng thẻ KPI chia đều khoảng cách.

**Không thay đổi:** dữ liệu, measure DAX, quan hệ, nội dung chữ, trang drill-through, slicer đồng bộ.

## Muốn đổi màu sau này?

Sửa file `FoodApp_Dashboard.Report/StaticResources/RegisteredResources/FoodAppTheme.json`, hoặc trong Power BI Desktop: `View → Themes → Customize current theme`.

## Bản 2 - căn chỉnh lại bố cục (từ DB.2)

- **Lưới thống nhất:** lề 20px, khoảng cách giữa các ô đều 12px trên cả 7 trang.
- **Hàng tiêu đề + slicer:** slicer cao hơn (62px), có lề trong, cách phần nội dung bên dưới 14px nên ô chọn không còn sát thẻ KPI hay biểu đồ.
- **Tiêu đề trang:** 16pt, luôn nằm trên 1 dòng. Tiêu đề trang 3 rút gọn thành "Catalog bị bỏ quên, web nhiều lượt xem nhưng ít đơn" (bản cũ quá dài, bị xuống dòng và che mất dòng phụ đề).
- **Trang 3 - Truy cập nhiều ≠ mua nhiều:** bảng 2 dòng (thừa khoảng trắng) đổi thành ma trận: 6 chỉ số theo hàng × 2 nhóm truy cập theo cột, lấp đầy khung.
- **Trang 4 - Phủ sóng chiến dịch theo nhóm giá trị:** đổi thành ma trận 5 chỉ số × High/Mid/Low (thêm dòng "Chi tiêu TB/khách" để thấy rõ nhóm High là nhóm giá trị nhất).
- **Trang 5:** 2 bảng phân khúc rộng hơn (716px) để chữ "Hành động đề xuất" không bị tràn.
- **Trang 6:** bảng Cleaning Log rộng hơn, cột trái gọn lại.
- **Biểu đồ cột/thanh:** bỏ tiêu đề trục và trục giá trị thừa (đã có nhãn số trên cột), nhìn thoáng hơn.

## Bản Final - thay bảng bằng biểu đồ (từ Final.1)

Giữ nguyên theme, bố cục và chỉnh sửa của nhóm trong Final.1 (5 trang; trang Chất lượng dữ liệu và trang drill-through đã bỏ). Chỉ thay 4 bảng, mỗi biểu đồ mới nằm đúng trong khung của bảng cũ:

| Trang | Bảng cũ | Biểu đồ mới |
|---|---|---|
| 3 | Truy cập nhiều ≠ mua nhiều (ma trận 6 chỉ số) | 3 biểu đồ cột nhỏ kể 1 câu chuyện: **truy cập nhiều gấp ~2 lần → chi tiêu chỉ ~40% → vì săn hàng giảm giá**. Nhóm truy cập cao = đỏ cà chua, thấp = xanh |
| 4 | Phủ sóng chiến dịch theo nhóm giá trị | Thanh 100% chồng High / Mid / Low: **Đã từng phản hồi** (xanh) vs **Chưa từng phản hồi** (đỏ) - thấy ngay ~1/2 khách VIP chưa từng phản hồi |
| 5 | So sánh 6 phân khúc RFM | Cột **% khách hàng vs % doanh thu** theo phân khúc (At Risk 24% khách → 40% doanh thu). Chi tiêu TB, recency, tỷ lệ phản hồi hiện khi rê chuột (tooltip) |
| 5 | Hành động đề xuất theo phân khúc | 6 thẻ hành động, viền màu theo phân khúc (cùng màu với biểu đồ bong bóng) |

Tiêu đề bảng phân khúc cũ nhắc "chuột phải → Drill through" nhưng trang drill-through đã bị xoá, nên tiêu đề mới bỏ hướng dẫn đó.
