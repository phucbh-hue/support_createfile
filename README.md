# support_createfile

Công cụ nội bộ: điền một mẫu Word (.docx) hàng loạt từ file dữ liệu (Excel/CSV) bằng Claude.

Người dùng tải mẫu và file dữ liệu lên, mô tả việc cần làm (ví dụ "mỗi người nhận 2 biên bản giao hàng: Lucky draw và Tour"), tùy chọn nhập các trường cần lấy, số khách hàng và số file dự kiến. AI đọc cả hai file, lập kế hoạch điền mẫu, trang web sinh từng file .docx và cho tải về (từng file hoặc một file .zip).

## Cách hoạt động

1. Trình duyệt đọc mẫu .docx và file dữ liệu, chuyển thành văn bản: mỗi đoạn văn, dòng bảng, ô của mẫu có một ID.
2. `api/generate.js` gửi phần văn bản đó cùng mô tả tới Claude API (API key nằm ở máy chủ) và nhận về một kế hoạch JSON: điền gì vào đâu cho từng file, dòng nào cần xóa, có cần thêm dòng nào không.
3. Trình duyệt áp kế hoạch lên mẫu để tạo file Word, giữ nguyên định dạng của mẫu. Phép tính tiền và đọc số bằng chữ do chương trình thực hiện, không để AI tính nhẩm.

File Word được tạo ngay trên máy người dùng, máy chủ không lưu file. Phần chữ của mẫu và dữ liệu có đi qua Claude API, nên chỉ dùng với dữ liệu được phép xử lý theo quy định của công ty.

## Cấu trúc

- `public/` - giao diện (`index.html`, `app.js`, `style.css`) và lõi xử lý `core.js`
- `api/generate.js` - gọi Claude, trả kết quả dạng stream
- `api/config.js` - cho giao diện biết cách đăng nhập và ai đang đăng nhập
- `api/auth/` và `lib/auth.js` - đăng nhập Google, phiên làm việc, mã truy cập
- `test/` - test phần lõi (`npm test`)
- `dev-server.mjs` - chạy thử ở máy cá nhân

## Biến môi trường

| Biến | Bắt buộc | Ý nghĩa |
|---|---|---|
| `ANTHROPIC_API_KEY` | Có | API key Claude. Không commit vào repo. |
| `APP_PASSWORD` | Nên có | Mã truy cập trang, tránh người ngoài dùng API key của công ty. |
| `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `SESSION_SECRET` | Không | Đặt đủ cả ba thì trang dùng đăng nhập Google thay cho mã truy cập. |
| `ALLOWED_EMAIL_DOMAINS`, `ALLOWED_EMAILS` | Không | Giới hạn tài khoản Google được vào (cách nhau bằng dấu phẩy). |
| `CLAUDE_MODEL` | Không | Mặc định `claude-opus-5-5`. |
| `CLAUDE_EFFORT` | Không | Mặc định `medium`. |

## Đăng nhập Google

1. Google Cloud Console > APIs & Services > Credentials > Create credentials > OAuth client ID, loại "Web application".
2. Authorized redirect URIs: `https://<tên miền của trang>/api/auth/callback` (chạy ở máy cá nhân: `http://localhost:3000/api/auth/callback`).
3. Đặt `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `SESSION_SECRET` (chuỗi ngẫu nhiên dài) trên Vercel rồi deploy lại.

## Chạy ở máy cá nhân

```bash
npm install
ANTHROPIC_API_KEY=... node dev-server.mjs   # mở http://localhost:3000
npm test
```

## Triển khai trên Vercel

1. Import repo này vào Vercel (Framework Preset: Other, không cần build command).
2. Thêm `ANTHROPIC_API_KEY` và `APP_PASSWORD` trong Settings > Environment Variables.
3. Deploy. Hàm `api/generate.js` được đặt thời gian chạy tối đa 300 giây trong `vercel.json`.

## Giới hạn hiện tại

- Mẫu phải là .docx; chỉ điền phần thân tài liệu (chưa điền header, footer, text box).
- Mỗi lần xử lý một mẫu và một file dữ liệu. File dữ liệu rất lớn (hàng trăm bản ghi) nên tách nhỏ để không vượt thời gian xử lý.
- Kết quả do AI lập nên cần người kiểm tra lại trước khi gửi cho khách hàng; phần "Ghi chú của AI" nêu các giả định và chỗ cần xem lại.
