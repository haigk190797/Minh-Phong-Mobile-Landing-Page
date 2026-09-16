---
name: landing-page-dien-thoai
description: Tạo hoặc chỉnh sửa landing page bán hàng (điện thoại, phụ kiện, sản phẩm bán lẻ nói chung) dựng từ dữ liệu có sẵn trong các thư mục cục bộ — ảnh sản phẩm + bảng giá Excel — thành 1 trang HTML tĩnh responsive, có thể copy phong cách từ 1 website mẫu do người dùng đưa link, và có thể kèm form thu số điện thoại khách ghi thẳng vào Google Sheet qua Google Apps Script. Dùng skill này ngay khi người dùng nhờ tạo landing page / trang bán hàng / trang giới thiệu sản phẩm từ ảnh và giá có sẵn trong thư mục, kể cả khi họ không dùng đúng từ "landing page" (ví dụ: "làm trang web bán máy giúp tui", "dựng trang giới thiệu giá cho shop"). Cũng dùng khi người dùng muốn thêm/sửa phần thu khách hàng (lead capture) vào Google Sheet trên 1 landing page đã có.
---

# Landing page bán hàng từ dữ liệu thư mục

Quy trình đầy đủ để dựng 1 landing page bán hàng tĩnh (HTML/CSS/JS, không framework) từ dữ liệu sản phẩm người dùng đã có sẵn trên máy, đúc kết từ lần đầu làm cho một shop điện thoại (Minh Phong Mobile). Áp dụng được cho bất kỳ shop bán lẻ nào có cấu trúc tương tự: vài thư mục sản phẩm, mỗi thư mục có ảnh + 1 file Excel giá.

## Khi nào dùng

- Người dùng đưa 1 hoặc nhiều thư mục chứa ảnh sản phẩm + bảng giá, nhờ dựng landing page
- Người dùng đưa link 1 website có sẵn làm mẫu giao diện ("làm giống trang này", "giao diện tui muốn là...")
- Người dùng muốn thêm form thu SĐT khách hàng, lưu vào Google Sheet, trên landing page

## Quy trình

### 1. Đọc dữ liệu sản phẩm từ thư mục

Dùng `Glob` quét toàn bộ thư mục gốc để thấy cấu trúc (mỗi thư mục con thường là 1 nhóm sản phẩm, vd "MÁY MỚI", "MÁY 99%"). Với mỗi thư mục: ảnh sản phẩm dùng `Read` trực tiếp được; file Excel giá thường **không** đọc trực tiếp được vì:

- Không có Python torng môi trường Windows mặc định (`python`/`python3` chỉ là App Execution Alias trỏ tới Microsoft Store, không chạy được) → đừng cố `pip install`/`markitdown` trước khi kiểm tra `Get-Command python` có thật không.
- File Excel có thể đang mở trong Excel trên máy người dùng → `System.IO.Compression.ZipFile]::OpenRead` sẽ báo "process cannot access the file".

**Cách chắc ăn nhất** (không phụ thuộc Python):
1. Copy file `.xlsx` sang scratchpad bằng PowerShell (`Copy-Item ... -Force`) để né file lock.
2. `.xlsx` thực chất là file zip — giải nén bằng `[System.IO.Compression.ZipFile]::ExtractToDirectory` (không truyền `entryNameEncoding` — tham số đó nhận `System.Text.Encoding`, không phải bool, truyền `$true` sẽ lỗi).
3. Đọc `xl/sharedStrings.xml` bằng `Read` — với bảng giá đơn giản (2 cột: Tên máy / Giá), toàn bộ dữ liệu thường nằm gọn trong danh sách `<si><t>...</t></si>` theo đúng thứ tự, đọc trực tiếp là đủ, không cần parse `sheet1.xml` để ghép tọa độ ô.

Nếu Python thật sự có sẵn (`Get-Command python` trả về kết quả hợp lệ, không phải store alias), dùng skill `xlsx` bình thường (markitdown/openpyxl) sẽ nhanh hơn.

### 2. Tham khảo giao diện mẫu (nếu người dùng đưa link)

Dùng Browser tool (`preview_start` với `url`, rồi `computer` screenshot + `get_page_text`) ghé thăm site mẫu. Ghi nhận cụ thể chứ đừng chỉ nhìn qua:
- Bảng màu chính (thường 2-3 màu: nền, chữ, accent)
- Bố cục hero/banner đầu trang
- Cách trình bày card sản phẩm (ảnh, tên, giá — có giá gạch ngang không, badge giảm giá không, nút CTA chữ gì)
- Có thanh điều hướng dính đáy kiểu app di động không (phổ biến ở site bán điện thoại Việt Nam)

Không cần/không nên copy y nguyên toàn bộ site (site mẫu thường có hàng trăm sản phẩm, nhiều trang) — chỉ lấy "DNA thiết kế" để áp cho đúng số sản phẩm thật của người dùng.

### 3. Chuẩn hoá yêu cầu

Nếu có skill `prompt-8-buoc` trong môi trường, dùng nó để tóm tắt yêu cầu thành 1 brief rõ ràng gửi người dùng duyệt trước khi code — đặc biệt hữu ích vì yêu cầu ban đầu kiểu "tạo landing page giúp tui" luôn thiếu ngữ cảnh.

Dùng `AskUserQuestion` hỏi đúng những thông tin **không thể tự suy luận được** từ dữ liệu có sẵn — thường chỉ 2 thứ:
- Tên shop/thương hiệu hiển thị
- Số điện thoại/Zalo để gắn vào nút CTA

Đừng hỏi những gì đã có trong dữ liệu (giá, tên sản phẩm, ảnh) hay có thể suy luận hợp lý.

### 4. Build trang HTML

- 1 file `index.html` duy nhất, CSS inline trong `<style>`, không dùng framework/CDN nặng (trang chạy được cả khi mở trực tiếp bằng file://, không cần build step).
- Copy toàn bộ ảnh sản phẩm sang 1 thư mục `assets/` với tên file thuần ASCII, không dấu, không khoảng trắng, không ký tự đặc biệt (đặc biệt tránh dấu `%` trong tên thư mục/file — dễ vỡ khi host lên web thật vì `%` là ký tự escape trong URL). Đường dẫn ảnh gốc trong thư mục dữ liệu của người dùng có thể có tiếng Việt có dấu/khoảng trắng — không sửa/xoá thư mục gốc, chỉ copy sang `assets/`.
- Cấu trúc: header sticky + hero + (tuỳ chọn) trust bar + 1 section riêng cho mỗi nhóm sản phẩm (dùng đúng tag/badge phân biệt, vd "Máy Mới 100%" vs "Máy 99% - Like New") + section lý do chọn shop + contact/footer + thanh CTA dính đáy cho mobile (gọi điện + Zalo).
- Responsive bắt buộc — test cả desktop lẫn mobile viewport trước khi báo xong (xem bước 5).

### 5. Test bằng preview server cục bộ

Môi trường Windows thường **không có Node/Python** để chạy `npx serve`/`python -m http.server`. Đồng thời `navigate` file:// trực tiếp trong Browser pane chỉ render "static snapshot" — ảnh và request thật **không load được**, dễ nhầm tưởng ảnh bị lỗi trong khi thực ra chỉ là do preview mode.

Cách chắc ăn: dùng script PowerShell HTTP server có sẵn trong `scripts/serve.ps1` của skill này:
1. Copy `scripts/serve.ps1` vào scratchpad của session (hoặc dùng thẳng nếu path này đọc được).
2. Tạo `.claude/launch.json` ở project root, trỏ `runtimeExecutable: powershell.exe`, `runtimeArgs` gọi script kèm `-Root "<đường dẫn tuyệt đối tới thư mục landing page>"` — **truyền đường dẫn có dấu tiếng Việt qua `runtimeArgs`** (argv), **không** hard-code trong nội dung file `.ps1` (file `.ps1` do `Write` tạo ra có thể bị `powershell.exe` đọc sai encoding, biến chuỗi tiếng Việt thành ký tự rác — đây là lỗi thực tế đã gặp).
3. `preview_start` với `name` trong launch.json → mở tab preview trỏ `http://localhost:<port>` → giờ `computer` screenshot, `resize_window` (preset mobile/desktop), `read_console_messages`, `read_network_requests` đều hoạt động thật.
4. `preview_stop` khi xong.

### 6. (Tuỳ chọn) Thu khách hàng vào Google Sheet

Nếu người dùng muốn form thu SĐT khách:

1. Thêm 1 mini-form dưới mỗi card sản phẩm (input `tel` + nút Gửi), gắn `data-product="<tên sản phẩm + phân loại>"` trên thẻ `<form>` để biết khách quan tâm sản phẩm nào.
2. JS gửi bằng `fetch(url, { method: 'POST', body: formData })` — **KHÔNG dùng `mode: 'no-cors'`.** Đây là lỗi tốn nhiều thời gian nhất khi làm phần này lần đầu: `no-cors` làm response luôn "opaque" (không đọc được), nên code chỉ có thể đoán mò và luôn hiện "Gửi thành công" dù request thật sự thất bại (từng gặp lỗi 403 bị che giấu hoàn toàn, tưởng cấu hình Google sai trong khi vấn đề chỉ nằm ở `no-cors`). Google Apps Script Web App (được deploy đúng cách) trả CORS header hợp lệ, gọi bằng `fetch` mặc định (không set `mode`) + đọc `res.json()` là đủ, tin cậy hơn nhiều và cho phép hiện đúng lỗi thật khi có sự cố.
3. Dùng file mẫu `references/code-gs-template.gs` của skill này làm Google Apps Script `doPost` — tự tạo sheet "Leads" với cột Thời gian / Số điện thoại / Sản phẩm nếu chưa có, ghi 1 dòng mỗi lần submit.
4. Hướng dẫn người dùng tự deploy (không tự làm thay được vì cần đăng nhập Google của họ — không có sẵn trong Browser tool, và không được tự ý xử lý OAuth thay người dùng):
   - Tạo Google Sheet mới (`sheets.new`)
   - Tiện ích mở rộng → Apps Script → dán code → Lưu
   - Triển khai → Triển khai mới → loại "Ứng dụng web" → **Thực thi với quyền của: Tôi**, **Ai có quyền truy cập: Bất kỳ ai** (không phải "Bất kỳ ai có Tài khoản Google" — 2 lựa chọn rất dễ nhầm)
   - Lần đầu deploy sẽ hiện màn hình xin quyền "(Unverified)" — bắt buộc bấm **Review Permissions → chọn tài khoản → Advanced/Nâng cao → Go to [dự án] (unsafe) → Allow**. Bỏ qua bước này là nguyên nhân phổ biến nhất gây lỗi 403 dù mọi cấu hình khác đều đúng.
   - Copy URL `.../exec`, dán vào hằng số endpoint trong `index.html`.
5. Nếu người dùng không rành kỹ thuật, cân nhắc dựng 1 trang hướng dẫn trực quan (Artifact) có mockup từng bước thay vì chỉ giải thích bằng chữ — hiệu quả hơn hẳn khi người dùng cần làm theo từng cú click. Tận dụng widget `copy code` để họ dán trực tiếp không phải gõ tay.
6. **Luôn test thử 1 lần submit thật** trước khi báo hoàn thành — không chỉ dựa vào việc code chạy không lỗi console. Nếu vẫn lỗi, debug bằng cách gọi `fetch` KHÔNG có `no-cors` qua `javascript_tool`/console để đọc được status code và body thật, thay vì đoán từ thông báo mù mờ trên UI.

## Việc không nên làm

- Đừng sửa/xoá dữ liệu gốc của người dùng (thư mục ảnh, file Excel) — chỉ đọc và copy sang nơi khác.
- Đừng tự ý đăng nhập/thao tác tài khoản Google của người dùng để deploy Apps Script — luôn hướng dẫn họ tự làm, chỉ hỗ trợ debug qua thông tin họ cung cấp lại (log lỗi, ảnh chụp màn hình).
- Đừng dùng `mode: 'no-cors'` cho fetch tới Google Apps Script (xem mục 6.2).
- Đừng copy nguyên xi toàn bộ nội dung site mẫu — chỉ lấy phong cách, số lượng sản phẩm phải khớp đúng dữ liệu thật của người dùng.
