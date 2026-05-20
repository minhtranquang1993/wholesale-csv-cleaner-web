# Wholesale CSV Cleaner - Web

Tool làm sạch dữ liệu Wholesale CSV online, deploy trên Vercel.

## Deploy lên Vercel

1. Push repo lên GitHub
2. Vào [vercel.com](https://vercel.com), import repo
3. Vercel tự detect Next.js và deploy

## Chạy local

```bash
npm install
npm run dev
```

Mở http://localhost:3000

## Cách sử dụng

1. Upload file CSV
2. Nhấn "Làm sạch và tải XLSX"
3. Download file kết quả

## Rules xử lý

- **Rule 1:** Nếu `first_source_url` rỗng và `url` không rỗng → gán `first_source_url = prefix + url`
- **Rule 2:** Lấy `utm_source` từ URL gán vào `first_user_source`
- **Rule 3:** Lấy `utm_medium` từ URL gán vào `first_user_medium`
