# Improvement Brief: Pivot Bệnh Nhân (header rác + multi-file) & Progress bar 2 tab

## Hiện trạng (đã verify trên code + file thật)

File mẫu: `~/Downloads/export_data_3f38beab-35a7-4eaf-9710-4ab6a297f618.xlsx`
(2214 dòng sheet, **header thật ở dòng 7**, dòng 1–6 là metadata bệnh viện + tiêu đề báo cáo + 1 dòng trống)

| # | Vấn đề | Vị trí | Hành vi hiện tại |
|---|---|---|---|
| 1 | Header hardcode row 1 | `app/lib/pivot.ts:110` `readXlsx` đọc `ws.getRow(1)`; `readCsv` dùng `Papa.parse(header:true)` | Pivot **fail hoàn toàn** với file export thật → throw `Không tìm thấy cột 'Tên Bệnh nhân'. Các cột hiện có: BỆNH VIỆN MẮT QUỐC TẾ DND - SÀI GÒN, 1` |
| 2 | Chỉ 1 file | `app/components/PivotTab.tsx:29` `useState<File \| null>(null)` | Muốn gộp 3 tháng phải chạy 3 lần rồi cộng tay |
| 3 | Không có progress | `PivotTab.tsx:189`, `WholesaleTab.tsx:165` | Chỉ đổi chữ nút thành "Đang xử lý..."; file 2000+ dòng treo UI im lặng |
| 4 | Wholesale chạy server | `app/api/clean/route.ts` | Không đo được progress thật; dính giới hạn body 4.5MB của Vercel (Pivot đã chuyển browser ở commit `eb18615`) |

### Baseline số học (ground truth, tính bằng Python trực tiếp từ XML của file mẫu)
```
data rows (từ dòng 8):            2207   (0 dòng thiếu tên bệnh nhân)
unique patients:                   156
TOTAL mode "Theo bệnh nhân":       478,735,371
unique sources:                     20
GRAND TOTAL mode "Theo nguồn":     478,735,371   (phải bằng mode 1)
patients trong mode 2 (union):     156           (phải bằng 156)
```
Top nguồn: `Ngoại viện 01` 116,700,536 (28 bn) · `Ngoại viện 02` 94,702,715 (22 bn) ·
`Marketing HCM` 57,043,400 (37 bn) · `BN Nội Viện` 48,329,450 (13 bn) ·
`(Không rõ nguồn)` 16,223,400 (13 bn) · `CTV01` 0 (2 bn) · `CTV03` 0 (1 bn)

Baseline gộp 2 file (upload cùng file 2 lần): **4414 rows, 156 patients, 957,470,743**.

> ⚠️ Con số này **KHÔNG** phải `2 × 478,735,371 = 957,470,742`. Lệch 1đ là **đúng**, không phải bug:
> `buildByPatient` cố ý làm tròn **sau khi cộng theo từng bệnh nhân** (pivot.ts:198-203) để dòng TỔNG CỘNG
> bằng đúng tổng các dòng hiển thị. 24/156 bệnh nhân có tổng lẻ, nên `round(2v) ≠ 2·round(v)`:
> `NGUYỄN VĂN VY` 7622149.5 → round 7622150 nhưng round(2v)=15244299 (−1);
> `HÙYNH HỮU PHƯỚC` 31044250.45 → round(2v)=62088501 (+1); `VÕ TÂM` 6073589.45 (+1). Net +1.
> **Tuyệt đối không "sửa" rounding để ra 742** — sẽ phá invariant TỔNG CỘNG = tổng dòng hiển thị.
>
> Gate phụ (**không** phải invariant toán học): trên **đúng 2 input đã verify** — file mẫu đơn lẻ và
> file mẫu add 2 lần — patient-mode TOTAL == source-mode GRAND TOTAL (478,735,371 và 957,470,743).
>
> ⚠️ **KHÔNG** được coi đây là invariant "đúng mọi input". Hai mode làm tròn ở **granularity khác nhau**:
> `buildByPatient` round 1 lần/bệnh nhân (pivot.ts:200-203), còn `buildBySource` round 1 lần/cặp
> (nguồn, bệnh nhân) rồi cộng subtotal đã round (pivot.ts:289-300). Hai bên chỉ bằng nhau khi mỗi
> bệnh nhân chỉ nằm trong 1 nguồn, hoặc phần lẻ tình cờ triệt tiêu. File mẫu bằng nhau là **may**:
> có 6 bệnh nhân trải trên >1 nguồn (`NGUYỄN ĐỨC HÙNG`, `TRẦN THỊ TÙNG`, `TRẦN TRƯƠNG TẤN TÀI`,
> `PHẠM ĐỨC MINH AN`, `MAI NGỌC LINH`, `PHAN XUÂN THI`) và **không ai trong số đó có tổng lẻ**.
> Đã repro phá được (**1 ví dụ minh hoạ, KHÔNG phải cận trên**): thêm 1 dòng tháng-B cho
> `NGUYỄN VĂN VY` (7622149.5, đang ở `Ngoại viện 02`) dưới nguồn `Marketing HCM` số 1000000.5
> → patient 479,735,371 vs source 479,735,372, lệch 1.
> Đây đúng là use case gộp nhiều tháng ở yêu cầu #2.
>
> **Cận trên của độ lệch là cấu trúc, không phải con số cố định**: lệch tối đa ~1đ cho **mỗi**
> cặp (nguồn, bệnh nhân) có tổng lẻ, **cả hai chiều**, nên nó **lớn dần theo data**. Đã verify:
> 50 bệnh nhân mỗi người chia 0.5/0.5 trên 2 nguồn → patient 50 vs source 100, **lệch +50**;
> chia 0.4/0.4 → patient 50 vs source 0, **lệch −50**. File mẫu đã có 24 bệnh nhân tổng lẻ và
> 6 bệnh nhân đa nguồn, nên gộp 3 tháng thật lệch vài chục đồng là bình thường.
> **Lệch khi gộp nhiều tháng KHÔNG phải bug của `mergeTables`** — đừng "sửa" rounding trong
> `buildBySource` vì sẽ phá tính chất `Subtotal [nguồn]` = tổng dòng hiển thị và tính khớp giữa
> sheet 1 / sheet 2 (`sourceSubtotals` tồn tại chính để bảo đảm điều đó).

### Test/metric có sẵn
Không có test framework, không có script verify. → phải tự dựng lưới verify trước khi tin kết quả.

## Mục tiêu cải thiện (đo được)

1. Upload thẳng file export gốc (header dòng 7) → pivot ra đúng 3 số baseline trên, **không phải sửa file Excel bằng tay**.
2. Thêm N file vào tab Pivot → gộp rồi pivot; tổng = tổng từng file cộng lại.
3. Cả 2 tab có progress bar % + nhãn phase, không đứng im.
4. Wholesale CSV: bỏ giới hạn upload 4.5MB, kết quả 3 số thống kê y hệt bản server.

## Quyết định đã chốt với user

| Câu | Chốt |
|---|---|
| Wholesale xử lý ở đâu | **Chuyển sang browser** (port logic sang `app/lib/clean.ts`); `app/api/clean/route.ts` **giữ lại, không xóa** |
| Dòng trùng khi gộp file | **Không dedup**. UI hiện bảng: từng file bao nhiêu dòng + header dòng nào + tổng dòng sau gộp → user tự kiểm |
| Detect header sai thì sao | **Hiện dòng detect được + cho sửa tay** (number input 1-based cho từng file) |

## Hướng đề xuất

### A. Header detection (`app/lib/pivot.ts`)
Refactor đọc file thành 2 bước: **raw matrix** → **detect header** → **table**.

- `readRawXlsx(buffer) → unknown[][]` — giữ đúng chỉ số dòng như Excel (index 0 = row 1), không skip dòng rỗng.
  **BẮT BUỘC** dùng `row.eachCell({includeEmpty:false})` ghi vào `new Array(ws.columnCount).fill(null)`.
  **KHÔNG** dùng `row.getCell(c)` trong vòng lặp: đã đo trên file mẫu — `getCell` **5408ms** vs
  `eachCell` **127ms** (42x), matrix ra **giống hệt nhau** (diff = 0, row 7 vẫn đủ 44 ô).
  `getCell` chậm vì nó materialize cả ô trống.
  Duyệt tới `ws.rowCount` (2214) chứ không phải `actualRowCount` (2213) — cần index tuyệt đối để
  dòng header người dùng nhập tay khớp với số dòng thấy trong Excel.
  Yield mỗi 500 dòng **bên trong** vòng lặp (không chỉ giữa các phase) để progress bar repaint được.
- `readRawCsv(text) → unknown[][]`: `Papa.parse(text, {header:false, skipEmptyLines:false})`.
  Đã verify index khớp tuyệt đối (dòng junk thứ 7 → idx 6) và `"2,500"` trong dấu ngoặc kép vẫn nguyên ô.
- `detectHeaderRow(matrix)`: quét 30 dòng đầu, chấm điểm 2 tầng:
  1. **Primary** = số cột anchor khớp (`Tên Bệnh nhân`, `Tiền sau miễn giảm`, `Nguồn` qua `normKey` sẵn có)
  2. **Tiebreak** = số **giá trị non-empty phân biệt** trong dòng. Đã đo trên file mẫu, tầng này decisive:
     row 1 = 2 · row 2-5 = 1 · row 6 = 0 · **row 7 = 44**. Merged cell làm dòng metadata collapse còn 1-2 giá trị.
  Cần tầng 2 vì `COL_TIEN_CANDIDATES` chỉ có **1** cách viết (pivot.ts:6) — file ghi
  `Tiền sau miễn giảm (VNĐ)` là mất anchor, chỉ còn điểm 1, dễ tie với dòng tiêu đề chứa chữ trùng.
  **Nếu điểm anchor cao nhất = 0**: KHÔNG mặc định về dòng 1 (sẽ throw đúng cái lỗi rác đang cần fix).
  Thay vào đó set `needsManualHeader` + pre-fill dòng có structural score cao nhất, UI hiện
  "Không tự nhận được dòng header — vui lòng nhập dòng header thủ công".
- `tableFromMatrix(matrix, headerIndex)`: cột = ô header không rỗng; rows từ `headerIndex+1`, bỏ dòng rỗng hoàn toàn.
  Cột trùng `normKey` → giữ cái đầu **và trả về `droppedColumns[]`** để UI hiện cảnh báo.
  (File mẫu: 44 header, 44 normKey phân biệt, 0 collision — nên case này chưa được test bởi data thật,
  càng cần cảnh báo hiện ra thay vì âm thầm. Nếu cột giữ lại là cột rỗng, `parseMoney('')` → 0 và
  `Nguồn` rỗng → `(Không rõ nguồn)`: pivot **chạy thành công với số tiền sai**, bảng row-count không phát hiện được.)
- **Override dòng header (contract rõ ràng)**: input 1-based, clamp `[1, matrix.length]`,
  rỗng/không hợp lệ → fallback về dòng auto-detect. Đổi dòng header **không đọc lại file**
  (matrix đã giữ trong RAM) → dựng lại table ngay.
- `parseFile()` giữ nguyên signature (wrapper: readRaw → detect → tableFromMatrix) ⇒ `/api/pivot/route.ts`
  không cần sửa mà cũng auto-detect được.
- `unwrapCell` thêm nhánh `richText` — **phòng ngừa**, không phải nguyên nhân đã quan sát:
  đã kiểm file mẫu, dòng 1-5 exceljs trả về **plain string**, không phải richText.

Trade-off: giữ nguyên matrix trong RAM (2214×44 ô/file) để đổi dòng header không cần đọc lại file.
Chấp nhận được cho tool nội bộ vài file.

### B. Gộp nhiều file
`mergeTables(tables[])`: union cột theo `normKey`, lấy cách viết đầu tiên làm canonical, **remap key từng row về canonical**.
> Bắt buộc phải remap: nếu file A viết `Tên Bệnh nhân` và file B viết `Tên bệnh nhân`, merge thô sẽ khiến `findColumn` chọn 1 tên, rows của file kia trả `undefined` → cộng thành 0 mà **không báo lỗi**.

Validate cột bắt buộc **theo từng file** lúc submit (theo mode đang chọn) → báo rõ file nào thiếu cột nào, thay vì lỗi chung sau khi đã gộp.

### C. Progress bar
`app/components/ProgressBar.tsx` dùng chung, hỗ trợ 2 state: **determinate** (%) và
**indeterminate** (sọc chạy, cho phase blocking không chia nhỏ được).
Progress theo **cost profile đã đo thật**, không chia đều cho đẹp.

**Pivot** — đo lại với `eachCell` (bản đã bắt buộc): exceljs `load` **246ms (48.2%)** +
matrix build **225ms (44.1%)** + tableFromMatrix 6ms (1.2%) + aggregate **1ms (0.2%)** + writeBuffer 32ms (6.3%).
⇒ ~92% cost nằm ở **đọc file lúc add**, không phải lúc bấm chạy. Chia rõ 2 phase:
- **Lúc add file (nơi chứa gần hết thời gian, ~459ms/file)**: `Đang đọc {tên file} (i/N)` →
  bar 0–100% cho từng file, yield mỗi 500 dòng trong matrix loop.
  `wb.xlsx.load()` (246ms, slice lớn nhất) là **1 call async opaque, không có progress hook** →
  set label **trước** khi gọi; yield mỗi 500 dòng chỉ cứu được phase matrix build.
- **Lúc bấm chạy (matrix đã ở RAM nên KHÔNG đọc lại file)**: chỉ còn tableFromMatrix 7ms (26.9%) +
  aggregate ~0ms + write 19ms (73.1%) = **~26ms** ở N=1. Band: tableFromMatrix 0–30% → gộp 30–40% →
  pivot 40–50% → ghi XLSX 50–100%.
  ⚠️ Band ở phase này là **nominal, không phải cost-derived** — toàn bộ phase chạy ≤226ms kể cả
  10 file, nên chia thế nào cũng gần như tức thì. Cụ thể `mergeTables` thay đổi mạnh theo N
  (N=1 8ms/20.5% · N=3 24ms/**49.0%** · N=5 29ms/31.2% · N=10 82ms/36.3%) — ở N=3 nó là phase
  đắt nhất, còn write tụt còn 12.2%. Không chỉnh band theo N vì không đáng; **không bịa filler**
  để bar trông "chạy đẹp". Progress có ý nghĩa nằm ở phase add.
- **Không** dùng indeterminate cho phase ghi của Pivot (19-32ms, không cần).

**Wholesale** — đo trên CSV synthetic 20k dòng / 1.91MB:
`parse 36ms (4.9%) · sort 41ms (5.6%) · rules 40ms (5.4%) · json_to_sheet 131ms (17.8%) · XLSX.write 489ms (66.4%)`
⇒ Band phải theo profile này: đọc file 0–10% → parse 10–20% → sort 20–25% → rule 25–35% →
json_to_sheet 35–50% → **ghi XLSX 50–100%**.
- **KHÔNG** chunk vòng lặp rule (chỉ 5.4% cost, chunk vô nghĩa).
- `json_to_sheet` và `XLSX.write` là **hai call đồng bộ liên tiếp**, không có progress callback bên
  trong mỗi call → **yield được giữa hai call** (nên giữ band 35–50 / 50–100 riêng), nhưng không chia
  nhỏ được bên trong. Set label `Đang ghi file XLSX (có thể mất vài giây)` + bar indeterminate
  **trước khi** vào `XLSX.write`. Trung thực hơn là % giả.
  (`Papa.parse` có `step`/`chunk` callback nếu sau này cần progress theo dòng ở phase parse — chưa cần, 4.9% cost.)

Dùng `await yieldToPaint()` (`setTimeout(0)` + `requestAnimationFrame`) giữa các phase để React 18 repaint.

### D. Wholesale sang browser
`app/lib/clean.ts` port **nguyên logic** từ route (`findColumn` lowercase, `isEmpty`, `extractUtm`,
sort theo cột thời gian, 3 rule, `XLSX.utils.json_to_sheet(rows,{header:columns})`).
Chỉ khác: `XLSX.write` dùng `type:"array"` thay `"buffer"` (browser không có Buffer).
Đã verify output **byte-identical**: `Buffer.from(arrayOut).equals(Buffer.from(bufferOut)) === true` (10,124,660 bytes).
`xlsx@0.18.5` khai báo `browser: {buffer:false, crypto:false, stream:false, process:false, fs:false}`
⇒ bundle client được, không kéo node built-in.

## Phạm vi

### Trong phạm vi
- `app/lib/pivot.ts` — raw matrix + detect header + tableFromMatrix + mergeTables
- `app/lib/clean.ts` — MỚI, port từ `api/clean/route.ts`
- `app/components/ProgressBar.tsx` — MỚI
- `app/components/PivotTab.tsx` — multi-file list, sửa dòng header, progress, stats gộp
- `app/components/WholesaleTab.tsx` — xử lý browser + progress
- `app/page.tsx` — nới card `max-w-lg` → `max-w-2xl` (bảng file cần chỗ)
- `scripts/verify-pivot.ts` — MỚI, script verify chạy `node` (type-stripping, không thêm dependency)

### Ngoài phạm vi (cố ý KHÔNG đụng)
- Rule làm sạch của Wholesale (3 rule + prefix + cột thời gian) — port y nguyên, không "cải thiện"
- Format Excel output của cả 2 mode pivot (màu/bold/money/freeze) — giữ đúng như hiện tại
- Header detection cho tab Wholesale (user chỉ yêu cầu cho Pivot)
- `app/api/pivot/route.ts` — dead code từ trước commit `eb18615`, chỉ báo lại, không xóa
- `app/api/clean/route.ts` — trở thành dead code sau change này; user đã chốt giữ lại
- Dedup dòng trùng, chọn tên cột tùy ý, preview data

## Rủi ro regression

| Rủi ro | Mitigation |
|---|---|
| Detect header sai trên file đã sạch (header dòng 1) | Anchor khớp ở dòng 1 → điểm 3 → chọn dòng 1. Verify bằng file test header dòng 1 |
| Detect trượt hết (anchor score = 0) → fallback dòng 1 → lỗi rác y như bug cũ | KHÔNG fallback dòng 1; set `needsManualHeader`, pre-fill structural guess, bắt user nhập |
| Cột trùng normKey bị bỏ âm thầm → tiền = 0 mà pivot vẫn "thành công" | `tableFromMatrix` trả `droppedColumns[]`, UI hiện cảnh báo vàng ở bảng từng file |
| Override dòng header ngoài range → `matrix[i]` undefined → crash | Clamp `[1, matrix.length]`, rỗng/invalid → fallback dòng auto-detect |
| Merge làm mất tiền vì tên cột lệch | `mergeTables` remap key theo `normKey`; verify bằng test 2 file lệch cách viết cột |
| Refactor `readXlsx` làm sai số pivot | Verify script so 3 số baseline (156 / 478,735,371 / 20) trước–sau |
| Matrix build bằng `getCell` → 5.4s/file, progress đứng | Bắt buộc `eachCell` (127ms, đã verify matrix identical) + yield mỗi 500 dòng |
| Progress bar nhảy 0→100 vì band sai cost | Band theo profile đã đo; phase `XLSX.write` (66%) dùng indeterminate |
| Wholesale port sai → lệch số dòng cập nhật | Verify: chạy cùng 1 CSV qua logic route cũ và `lib/clean.ts` mới, so 3 số thống kê |
| `xlsx` không bundle được ở client | Đã verify `browser` field disable fs/buffer/stream/crypto/process. `npm run build` vẫn phải pass |
| Nhiều file lớn → hết RAM | Hiện tổng dòng cho user thấy; không xử lý ngoài phạm vi |

## Acceptance
- [ ] Upload file mẫu gốc (header dòng 7): mode 1 ra 156 bn / 478,735,371; mode 2 ra 20 nguồn / grand total 478,735,371
- [ ] File có header dòng 1 vẫn chạy đúng (không regress)
- [ ] Add cùng file 2 lần: 4414 dòng, 156 bn, **957,470,743** (không phải 742 — xem note rounding ở trên)
- [ ] **Gate phụ (chỉ 2 input đã verify)**: patient-mode TOTAL == source-mode GRAND TOTAL trên file mẫu
      đơn lẻ và file mẫu add 2 lần. KHÔNG áp cho input khác — 2 mode round khác granularity (xem note trên)
- [ ] 2 file lệch cách viết tên cột: không mất tiền
- [ ] File không detect được header: hiện yêu cầu nhập tay, KHÔNG throw lỗi "cột hiện có: BỆNH VIỆN..."
- [ ] Override dòng header 0 / rỗng / 99999: không crash
- [ ] Wholesale: 3 số thống kê khớp logic route cũ trên cùng input
- [ ] Progress bar: **không đứng yên quá 1s ở bất kỳ mốc nào**; phase blocking dùng indeterminate
- [ ] `npm run build` pass, không lỗi TypeScript
