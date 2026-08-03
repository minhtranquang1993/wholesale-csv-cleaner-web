# Decision Summary — Gate 2

## Quyết định đã chốt

| # | Quyết định | Chốt gì | Lý do |
|---|---|---|---|
| 1 | Wholesale xử lý ở đâu | **Chuyển sang browser** (`app/lib/clean.ts`). `app/api/clean/route.ts` **giữ lại trên disk, không xóa** | Progress % thật cần xử lý client; bỏ luôn giới hạn body 4.5MB của Vercel; data không rời máy |
| 2 | Dòng trùng khi gộp file | **Không dedup**. Hiện bảng từng file: số dòng + dòng header + tổng dòng sau gộp | Dedup theo biên lai có thể xoá dòng thật (1 biên lai 2 dòng cùng dịch vụ) → mất tiền âm thầm |
| 3 | Detect header sai | **Hiện dòng detect được + cho sửa tay** (input 1-based mỗi file) | Detect không bao giờ đúng 100%; phải có đường thoát không cần sửa file Excel |
| 4 | Auto-detect header áp tab nào | **Chỉ Pivot Bệnh Nhân** | Đúng phạm vi user yêu cầu; Wholesale anchor khác hẳn, viết detect riêng = tăng surface lỗi cho tab đang ổn |
| 5 | Multi-file áp tab nào | **Chỉ Pivot Bệnh Nhân** | Đúng phạm vi user yêu cầu; Wholesale multi-file phát sinh câu hỏi sort/cột lệch ngoài yêu cầu |
| 6 | Output khi gộp nhiều file | **1 file pivot gộp duy nhất**, giữ đúng 2 sheet như hiện tại (mode 2) | Đúng ý "gộp data"; không thêm sheet mới → không đụng format output đã verify |
| 7 | Lưới verify | **Commit `scripts/verify-pivot.ts`**, chạy bằng `node` type-stripping, không thêm dependency | Có lưới chặn regression cho lần sửa sau |

## Giữ nguyên (KHÔNG đổi lần này)

- 3 rule làm sạch Wholesale + `BASE_PREFIX` + danh sách cột thời gian → port y nguyên, không "cải thiện"
- Format Excel output cả 2 mode pivot: màu header `1F4E79`, `GROUP_COLORS`, money `#,##0`, freeze panes, subtotal bold, grand total — giữ nguyên về **cấu trúc + style** (tên/thứ tự sheet, giá trị dòng header, `numFmt`/fill/font ở ô có style, freeze pane, số dòng). **Không** so byte: `new ExcelJS.Workbook()` stamp `created`/`modified` bằng `new Date()` vào `docProps/core.xml` → 2 lần build cùng workbook đã khác 62 byte (đã verify). Muốn so byte thì phải pin `wb.created`/`wb.modified` trước. (Byte-identical đã verify **chỉ áp cho** SheetJS `type:"array"` vs `"buffer"` ở Wholesale — khác path, đừng lẫn.)
- **Rounding round-then-sum** (`pivot.ts:198-203`): làm tròn từng bệnh nhân TRƯỚC rồi cộng. Không đổi thành sum-then-round dù merge total lệch 1đ so với 2× single
- `app/api/pivot/route.ts` — dead code từ commit `eb18615`, chỉ báo lại, không xóa
- `app/api/clean/route.ts` — thành dead code sau change này, user chốt giữ
- Tab Wholesale: vẫn 1 file, vẫn header dòng 1, vẫn `.csv` only
- Không dedup, không cho chọn tên cột tùy ý, không preview data

## Acceptance Criteria (coi như XONG khi:)

### Mục tiêu cải thiện đạt
- [ ] Upload file mẫu gốc (header dòng 7, không sửa tay): mode 1 → **156 bn / 478,735,371**; mode 2 → **20 nguồn / grand total 478,735,371**
- [ ] Add cùng file 2 lần: **4414 dòng / 156 bn / 957,470,743** (KHÔNG phải 742 — round-then-sum, xem brief)
- [ ] **Gate phụ (chỉ 2 input đã verify)**: patient-mode TOTAL == source-mode GRAND TOTAL trên file mẫu đơn lẻ và file mẫu add 2 lần. **KHÔNG** phải invariant chung: `buildByPatient` round per-bệnh-nhân, `buildBySource` round per-(nguồn, bệnh nhân) → gộp nhiều tháng thật có thể lệch tới ~1đ cho **mỗi** bệnh nhân có tiền lẻ trải trên nhiều nguồn, **cả hai chiều** (đã verify lệch ±50 với 50 bệnh nhân chia lẻ). Lệch này là bình thường, không phải bug
- [ ] 2 file lệch cách viết tên cột (`Tên Bệnh nhân` vs `Tên bệnh nhân`): không mất tiền
- [ ] Progress bar cả 2 tab: **không đứng yên quá 1s ở bất kỳ mốc nào**; phase `XLSX.write` (66% cost) dùng indeterminate
- [ ] Wholesale: bỏ được giới hạn 4.5MB (không còn POST `/api/clean`)

### Baseline behavior không tụt (chặn regression)
- [ ] File header dòng 1 vẫn pivot đúng
- [ ] Wholesale: 3 số thống kê (`first_source_url` / `first_user_source` / `first_user_medium`) khớp **chính xác** logic route cũ trên cùng input
- [ ] Format Excel output 2 mode pivot không đổi (mở được, đủ màu/bold/money/freeze)
- [ ] `npm run build` pass, không lỗi TypeScript

### Edge case phải xử lý
- [ ] File không detect được header (anchor score 0): hiện "nhập dòng header thủ công", KHÔNG throw lỗi rác `cột hiện có: BỆNH VIỆN MẮT...`
- [ ] Override dòng header = 0 / rỗng / 99999: không crash
- [ ] Cột trùng `normKey` bị bỏ: hiện cảnh báo vàng, không âm thầm
- [ ] `matrix` build dùng `eachCell` không phải `getCell` (127ms vs 5408ms, đã verify identical)
