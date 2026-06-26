# Project Brief: Data Cleaner — thêm tab Pivot Bệnh Nhân

## Vấn đề
Hai công cụ pivot bệnh nhân hiện là app desktop Python (tkinter), phải cài Python + chạy local. Người dùng muốn dùng ngay trên web cùng chỗ với tool làm sạch CSV hiện có.

## Giải pháp
Mở rộng web app hiện tại thành "Data Cleaner" với tab bar ngang: giữ tab Wholesale CSV cũ và thêm tab "Pivot Bệnh Nhân" có 2 chế độ pivot, port logic Python (pandas + openpyxl) sang TypeScript, output Excel có format đẹp.

## Đối tượng & Mục tiêu
- Người dùng: nội bộ, xử lý file Excel/CSV danh sách bệnh nhân
- Success metrics: kết quả pivot + format Excel khớp với output của 2 file Python gốc; tab cũ không bị ảnh hưởng

## Phạm vi đã thống nhất

### 🎯 MVP (in-scope)
- [ ] Đổi tên web "Wholesale CSV Cleaner" → "Data Cleaner" (title, metadata, header) — Dễ
- [ ] Tab bar ngang trên cùng: Tab 1 "Wholesale CSV" (giữ nguyên), Tab 2 "Pivot Bệnh Nhân" — Dễ
- [ ] Tách `page.tsx` thành component riêng cho từng tab (WholesaleTab, PivotTab) — Dễ
- [ ] Tab Pivot: segmented control 2 mode ("Theo bệnh nhân" / "Theo nguồn + bệnh nhân") — Dễ
- [ ] Dropzone nhận `.xlsx`, `.xls`, `.csv` — Dễ
- [ ] API mới `/api/pivot` (tab cũ giữ `/api/clean`) — TB
- [ ] Mode 1: group theo `Tên Bệnh nhân` → tổng `Tiền sau miễn giảm` + dòng TỔNG CỘNG — TB
- [ ] Mode 2: group `Nguồn`+`Tên Bệnh nhân` → subtotal mỗi nguồn + grand total; thêm sheet 2 "Tổng hợp Nguồn" — TB
- [ ] Output Excel có format (header xanh đậm/chữ trắng, money `#,##0` căn phải, freeze panes, mode 2 tô màu xen kẽ theo nhóm nguồn + subtotal bold + grand total nổi bật) — TB
- [ ] Kết quả hiển thị: số bệnh nhân unique (+ số nguồn ở mode 2) + nút download + reset — Dễ

### ⭐ Nice-to-have
- [ ] Khi thiếu cột, báo lỗi kèm danh sách các cột tìm thấy trong file
- [ ] Preview vài dòng trước khi tải

### ⏳ Out-of-scope / Tương lai
- Chọn tên cột tùy ý (không hardcode)
- Nhiều sheet input

## Quyết định kỹ thuật
- **Kiến trúc**: Tab bar ở client (`page.tsx` quản state tab), 2 component tab tách riêng, 2 API route độc lập (`/api/clean` cũ, `/api/pivot` mới).
- **Thư viện**: dùng **`exceljs`** cho cả đọc + ghi Excel (.xlsx/.xls), **papaparse** cho CSV. Bỏ `xlsx` ở route pivot. (Tab cũ `/api/clean` giữ nguyên `xlsx` để không regress.)
- **Match tên cột linh hoạt**: `normalize("NFC")` + `toLocaleLowerCase("vi")` + strip NBSP/zero-width + strip prefix `Range[...]`, so khớp candidate list. Cột cần: `Tên Bệnh nhân`, `Nguồn`, `Tiền sau miễn giảm`.
- **Parse tiền VN** (rủi ro cao): normalizer riêng xử lý "1.000.000" (chấm ngăn nghìn) và "...,50" (phẩy thập phân); cell đã là number thì giữ; làm tròn số nguyên VND khi cộng tổng.
- **Sort**: `Intl.Collator("vi")` cho tên/nguồn; chèn dòng subtotal/total SAU khi sort.
- **Nguồn rỗng** → gán "(Không rõ nguồn)" như bản Python.

## Rủi ro & Mitigation
- **Parse tiền sai** (dấu chấm ngăn nghìn) → normalizer chuyên dụng + unit-test vài mẫu trước khi tin kết quả.
- **Unicode NFD từ openpyxl** → luôn NFC hóa trước khi match cột.
- **Regress tab cũ** → không đụng `/api/clean` và logic Wholesale; chỉ refactor tách component, verify build.
- **`exceljs` đọc .xls cũ (binary)** → nếu exceljs không đọc được .xls, fallback thông báo yêu cầu .xlsx/.csv.

## Tiêu chí thành công (Acceptance)
- Mode 1 & Mode 2 cho ra số liệu + cấu trúc sheet khớp output Python gốc.
- File Excel tải về mở được, có format (màu/bold/money/freeze) đúng.
- Tab Wholesale CSV hoạt động y như trước.
- Tiền dạng "1.000.000" được cộng đúng (không thành 1).
- `npm run build` pass, không lỗi TypeScript.

## Constraints
- Giữ kiến trúc Next.js App Router hiện tại, deploy không cần Python runtime.
- Không thay đổi hành vi tab Wholesale CSV.

## Next: cook
