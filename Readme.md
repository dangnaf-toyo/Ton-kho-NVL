# Quản lý Tồn kho NVL — TOYOTAKI

Dashboard quản lý tồn kho nguyên vật liệu (NVL), có in tem QR và quét QR để ghi nhận nhập/xuất kho. Chạy dưới dạng 1 trang HTML tĩnh (frontend) gọi vào 1 Google Apps Script Web App (backend), dữ liệu lưu trong Google Sheet.

## Link

- **Trang chạy thật (GitHub Pages):** https://dangnaf-toyo.github.io/Ton-kho-NVL/
- **Repo code:** https://github.com/dangnaf-toyo/Ton-kho-NVL
- **Google Sheet database:** https://script.google.com/u/0/home/projects/1hZhmLuPNHzHS-mfNYDRyxRgyN1tjFJ9A9xNqKT24N_QReNG5hmzTEiZx/edit
- **Apps Script Web App (backend):** `https://script.google.com/macros/s/AKfycbxXYRfz_wuz3RE0a8TCr4zWyzF23G5rc3SWo4ijKLcv7UkLBfdw2um1HbNwRZgAnCc/exec` (biến `GAS_URL` trong `index.html`)

## Cấu trúc code

| File | Vai trò |
|---|---|
| `index.html` | Toàn bộ frontend: UI, Chart.js cho biểu đồ kế hoạch, in tem QR (`qrcode-generator`), quét QR bằng camera (`jsQR` + `BarcodeDetector`, nhúng inline offline) |
| `code.js` | Backend Apps Script — paste vào Apps Script editor của file Google Sheet (`SHEET_ID`), deploy làm Web App. **Hiện chưa được deploy tách file `.gs` riêng trong Apps Script — toàn bộ nằm trong 1 file.** |

Deploy khi sửa `code.js`: mở Google Sheet ở link trên → Extensions → Apps Script → dán đè nội dung → Deploy → Manage deployments → chọn deployment hiện có → sửa version mới → Deploy (giữ nguyên `GAS_URL` cũ).

## Google Sheet — cấu trúc dữ liệu

File Sheet ID: `1sFRigbmMAKdKX2spRrq6IPjPeNA3hDaOECobqaXMa-8`, 6 tab:

| Tab | Vùng đọc | Nội dung |
|---|---|---|
| `Tồn Đầu Kỳ` | A4:C100 | Mã NVL, ngày đầu kỳ, tồn đầu kỳ — điểm mốc để tính tồn hiện tại |
| `Kế Hoạch` | 3 dòng/mã × 30 cột ngày, từ dòng 5 | Kế hoạch nhập/tiêu hao/tồn dự kiến 30 ngày tới, theo từng mã NVL |
| `Giao Dịch` | A4:G10000 | Lịch sử nhập/xuất (append-only): ngày, loại IN/OUT, mã, số lượng, ghi chú, tồn sau, timestamp |
| `Tem NVL` | A4:I... (append) | Tem QR đã in cho từng pallet NVL: TagNo, mã, tên, ngày nhập, số lượng, đơn vị, ghi chú, timestamp, trạng thái |
| `Cài Đặt` | A4:B100 | Tham số "bề rộng" theo mã NVL |
| `Tồn Hiện Tại` | A4:C100 | Tồn hiện tại theo mã (state, bị ghi đè mỗi lần có giao dịch) |

**⚠️ Điểm quan trọng cần biết trước khi sửa code:**

1. **4 tab `Tồn Đầu Kỳ` / `Kế Hoạch` / `Cài Đặt` / `Tồn Hiện Tại` dùng "dòng cố định theo vị trí"** — dòng thứ N trong các tab này PHẢI đúng với mã thứ N trong mảng `MATERIALS` khai báo cứng ở đầu `code.js` (hiện có 7 mã: `ADC12-DAK`, `ADC12-CHT`, `ADC12-TAQ`, `ADC6-CHT`, `HD2-HDT`, `EZDA3-HDT`, `EZDA3-MET`). **Thêm mã NVL mới bắt buộc phải:** (1) thêm mã vào cuối mảng `MATERIALS` trong code, (2) thêm 1 dòng tương ứng ở cuối cả 4 tab trên, đúng thứ tự. Sai thứ tự → dữ liệu lệch hàng loạt không có cảnh báo.
2. **`Tem NVL` và `Giao Dịch` KHÔNG dùng mảng `MATERIALS`** — chúng validate mã NVL theo file **KHSX Master** riêng (Sheet ID `1WMF1EoGsmKNVaYIQwEe9i9tw6k_dZC1gSbHJp7RNsC0`, cột V:W từ dòng 7, cache 10 phút qua `getMasterMaterials()`). File KHSX Master này còn được dùng chung bởi 3 hệ thống khác (Dashboard Đúc, In tem Đúc, Chuyển công đoạn). Vì 2 nguồn danh mục khác nhau, 1 mã mới thêm vào KHSX Master có thể tạo được tem/giao dịch nhưng **không có dòng tồn đầu kỳ/tồn hiện tại tương ứng** → tồn kho sai.
3. **Đơn vị lẫn lộn kg/tấn**: `Tem NVL` lưu số lượng theo đơn vị gốc (`kg` hoặc `pcs`), nhưng khi ghi vào `Giao Dịch`/`Tồn Hiện Tại` lại tự quy đổi sang **tấn** (chia 1000 nếu đơn vị là kg).
4. **`Tồn Hiện Tại` là dữ liệu suy ra được** (= tồn đầu kỳ + tổng nhập − tổng xuất từ `Giao Dịch`), nhưng bị ghi đè thủ công mỗi lần có giao dịch — có thể lệch nếu 1 giao dịch bị sửa/xoá tay mà quên cập nhật lại tồn.

> Cả 4 điểm trên đã được phân tích kỹ hơn (kèm đề xuất giải pháp khi chuyển sang database thật) trong file `KE_HOACH_MIGRATE_DATABASE.md` ở repo `Dashboard_SL_CL`, mục 9.5.

## Backend — các action API (`code.js`)

`doGet(e)` nhận `?action=...`, `doPost(e)` nhận JSON body `{action, ...}`. Các action chính:

| Action | Việc làm |
|---|---|
| `getAllData` | Trả toàn bộ dữ liệu (opening, planDaily, transactions, settings, currentStock) — 1 lần gọi duy nhất khi load trang |
| `getMasterMaterials` | Đọc danh mục mã/tên NVL từ KHSX Master (cache 10 phút) |
| `createTem` | Sinh Tag No (`TYK-yyyyMMdd-XXXX`), ghi tem vào `Tem NVL`, tự động ghi giao dịch nhập kho tương ứng |
| `processMultiTransaction` | Ghi nhiều giao dịch nhập/xuất 1 lần (từ giỏ hàng quét QR trên điện thoại), cập nhật trạng thái tem + tồn hiện tại |
| `checkFIFO` | Với 1 tem sắp xuất, trả danh sách tem cùng mã nhập trước đó còn tồn (kiểm tra FIFO) |
| `updateOpening` / `updateStock` / `updateSettings` / `updatePlanNhap` | Cập nhật tay các tab tương ứng |
| `recalcAllPlanStock` | Tính lại tồn dự kiến 30 ngày cho tất cả mã (dùng khi sửa tiêu hao/tồn đầu kỳ trực tiếp trong Sheet) |

## Quét QR bằng camera (frontend)

Từng bị lỗi **camera mở được nhưng không đọc được QR** — nguyên nhân là bản cũ dùng `zbar-wasm` (WASM) làm engine duy nhất, và WASM có thể không khởi tạo được trên một số điện thoại/trình duyệt (đặc biệt Safari iOS) mà không có fallback. Đã sửa (xem lịch sử commit `index.html`): ưu tiên `BarcodeDetector` (API gốc trình duyệt, nhanh, có trên Chrome/Edge Android), luôn có **jsQR** (thư viện thuần JS, không WASM) làm phương án dự phòng chạy trên mọi trình duyệt — cùng thư viện đang chạy ổn định trên hệ thống "Chuyển công đoạn" (`D:\Project\MES\Chuyển công đoạn`).

## Hệ thống liên quan

Đây là 1 trong 6 hệ thống MES nội bộ của TOYOTAKI, tất cả được ghi lại tổng thể trong `KE_HOACH_MIGRATE_DATABASE.md` (repo `Dashboard_SL_CL`):

- Dashboard Sản lượng & Giao hàng, Dashboard Chất lượng (`Dashboard_SL_CL`)
- Dashboard Đúc, In tem công đoạn Đúc, Đọc QR chuyển công đoạn (`D:\Project\MES\Dashboard Đúc`, `Intem QR`, `Chuyển công đoạn`)
- **Quản lý tồn kho NVL (dự án này)**

4/6 hệ thống trên (Đúc, In tem, Chuyển công đoạn, NVL) đọc chung 1 file **KHSX Master** — sửa cấu trúc file đó cần kiểm tra ảnh hưởng tới cả 4.
