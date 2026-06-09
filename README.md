# QuynhPrivate Portal - Trình Đọc Tin Nhắn Microsoft Teams (TTS)

QuynhPrivate Portal là ứng dụng Web App hiện đại (giao diện Dark Mode cao cấp) được phát triển bằng React, TypeScript và Vite. Tính năng cốt lõi của ứng dụng là tích hợp với Microsoft Teams để tự động lắng nghe và đọc to tin nhắn cá nhân (1-on-1 chats) bằng công nghệ chuyển đổi văn bản thành giọng nói (Text-to-Speech).

---

## 1. Công nghệ Sử dụng
*   **Frontend**: React (Functional Components + Hooks), TypeScript
*   **Build Tool**: Vite
*   **Xác thực**: `@azure/msal-browser` (Microsoft Authentication Library) để đăng nhập bảo mật OAuth.
*   **API tích hợp**: Microsoft Graph API (`/me/chats`, `/chats/{id}/messages`).
*   **Chuyển đổi Giọng nói**: Web Speech API (`window.speechSynthesis`) giọng đọc tiếng Việt tự nhiên của trình duyệt.
*   **Định kiểu giao diện**: CSS Vanilla cao cấp (Dark Mode, responsive).

---

## 2. Cấu hình Môi trường Server (Port, Host, Allowed Hosts)
Bạn có thể cấu hình các thông số chạy của Development Server thông qua tệp cấu hình [.env.development](file:///.env.development) ở thư mục gốc:
```env
# Cổng chạy của máy chủ phát triển
PORT=5173

# Địa chỉ IP của host để chạy (Ví dụ: 127.0.0.1, localhost, hoặc 0.0.0.0 để mở rộng cho mạng LAN)
HOST=127.0.0.1

# Cho phép truy cập từ tên miền công khai (Đặt 'true' để mở cho mọi Host, hoặc nhập tên miền cụ thể)
ALLOWED_HOSTS=true
```

---

## 3. Cách Vận hành Cục bộ (Localhost)

Trước khi khởi chạy, cài đặt đầy đủ các gói thư viện:
```bash
npm install
```

### A. Khởi chạy Chế độ Phát triển (Dev Mode)
```bash
npm run dev
```
Truy cập ứng dụng tại địa chỉ: `http://localhost:<PORT>` (mặc định: `http://localhost:5173`).

### B. Đóng gói sản phẩm tối ưu (Build Production)
```bash
npm run build
```
Các tệp tĩnh tối ưu sẽ được biên dịch và lưu vào thư mục `dist/`.

---

## 4. Hướng dẫn Triển khai (Deploy) dự án lên VPS DirectAdmin
Dự án được viết hoàn toàn bằng Client-side React, do đó bạn có hai phương án triển khai lên VPS DirectAdmin chạy CentOS 9:

### PHƯƠNG ÁN 1: Triển khai Tĩnh (Static Deployment - Khuyên dùng)
Vì dự án chạy trực tiếp trên trình duyệt (không cần server-side Node.js chạy ngầm), đây là cách tối ưu và nhẹ nhất.

1.  **Biên dịch cục bộ**: Chạy lệnh `npm run build` trên máy cá nhân để tạo thư mục `dist/`.
2.  **Upload lên VPS**: Upload toàn bộ các file bên trong thư mục `dist/` vào thư mục gốc tên miền của DirectAdmin:
    `/home/[user]/domains/[domain_name]/public_html/`
3.  **Cấu hình chuyển tiếp đường dẫn (SPA Routing)**:
    Để tránh lỗi `404 Not Found` khi F5 lại trang, hãy tạo file `.htaccess` trong thư mục `public_html` với nội dung sau:
    ```apache
    <IfModule mod_rewrite.c>
      RewriteEngine On
      RewriteBase /
      RewriteRule ^index\.html$ - [L]
      RewriteCond %{REQUEST_FILENAME} !-f
      RewriteCond %{REQUEST_FILENAME} !-d
      RewriteRule . /index.html [L]
    </IfModule>
    ```

---

### PHƯƠNG ÁN 2: Chạy Server Phát triển (Dev Server) trên VPS & cấu hình Reverse Proxy
Nếu bạn muốn chạy trực tiếp mã nguồn trên VPS và tự cập nhật mã nguồn (Dev Server):

#### Bước 1: Đồng bộ và Phân quyền Thư mục
Đăng nhập SSH Terminal với quyền root và chạy chuỗi lệnh sau để dọn dẹp phân quyền sở hữu về tài khoản user DirectAdmin của bạn:
```bash
# Đổi chủ sở hữu thư mục về user quản lý tên miền
chown -R quynhprivate:quynhprivate /home/quynhprivate/domains/quynhprivate.sagoker.vn/public_html

# Cấp quyền đọc/ghi chuẩn cho thư mục (755) và file tĩnh (644)
find /home/quynhprivate/domains/quynhprivate.sagoker.vn/public_html -type d -exec chmod 755 {} +
find /home/quynhprivate/domains/quynhprivate.sagoker.vn/public_html -type f -exec chmod 644 {} +

# Khôi phục riêng quyền thực thi cho các script chạy ngầm của Node.js
chmod -R +x /home/quynhprivate/domains/quynhprivate.sagoker.vn/public_html/node_modules/.bin
```

#### Bước 2: Cấu hình Môi trường chạy trên VPS
Không cần thay đổi mã nguồn [vite.config.ts](file:///vite.config.ts). Bạn chỉ cần cấu hình trực tiếp các biến trong file [.env.development](file:///.env.development) trên VPS để cho phép truy cập qua tên miền:
```env
PORT=5173
HOST=127.0.0.1
ALLOWED_HOSTS=true
```

#### Bước 3: Cấu hình Reverse Proxy trên DirectAdmin
Do không thể cho Node.js chạy trực tiếp cổng `80` (xung đột với Apache/Nginx của DirectAdmin), ta tạo một cổng chuyển tiếp:
1.  Đăng nhập DirectAdmin quyền **Admin** -> Vào **Server Manager** -> Chọn **Custom HTTPD Configurations**.
2.  Chọn tên miền `quynhprivate.sagoker.vn`.
3.  Tại ô tùy chỉnh văn bản trống lớn phía trên cùng (dành cho file `httpd.conf`), dán chính xác đoạn mã cấu hình Apache:
    ```apache
    ProxyRequests Off
    ProxyPreserveHost On

    <Location />
        ProxyPass http://127.0.0.1:5173/
        ProxyPassReverse http://127.0.0.1:5173/
    </Location>
    ```
4.  Bấm **SAVE**. Web Server sẽ tự động hòa trộn cấu hình và nạp lại hệ thống.

#### Bước 4: Khởi chạy dự án ngầm
Do phân vùng `/home` của DirectAdmin dính cờ bảo mật `noexec` ngăn gọi trực tiếp file nhị phân trong `bin`, ta khởi chạy gián tiếp thông qua trình biên dịch Node.js:
```bash
node ./node_modules/vite/bin/vite.js
```
*Để tiến trình chạy ngầm liên tục khi tắt cửa sổ SSH, bạn nên dùng gói quản lý tiến trình `pm2` để khởi chạy:*
```bash
pm2 start ./node_modules/vite/bin/vite.js --name "quynhprivate-dev"
```

---

## 5. Tài liệu đính kèm
*   [walkthrough.md](file:///walkthrough.md): Hướng dẫn vận hành hệ thống và thiết lập Client ID để liên kết Microsoft Teams.
