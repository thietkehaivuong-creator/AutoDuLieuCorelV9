# 📖 AutoDuLieuCorelV9 - PFI Search Tool

## ✨ Tính năng chính

✅ **Tìm kiếm PFI** - Nhập đơn hàng, tìm ngay dữ liệu sản phẩm  
✅ **Hiển thị đầy đủ** - 20 cột dữ liệu (Customer, PFI, Description, Sizes, FAO, v.v.)  
✅ **Định dạng ngày linh hoạt** - dd/mm/yyyy, mm/dd/yyyy, mm/yyyy, custom formats  
✅ **Copy dữ liệu** - Copy từng ô hoặc tất cả  
✅ **Insert vào CorelDRAW** - Tự động thêm text box với thông tin  
✅ **Lưu credentials an toàn** - Không lộ username/password  
✅ **API REST** - Swagger UI để test & quản lý  

---

## 📋 Yêu cầu hệ thống

- **Windows 10/11** (64-bit)
- **CorelDRAW X7** trở lên (2024, 2023, 2022, v.v.)
- **.NET 8.0 Runtime**
- **Docker Desktop** (tùy chọn)

---

## 🚀 Cài đặt nhanh (5 phút)

### **Bước 1: Cài .NET 8.0 Runtime**

```powershell
# Mở PowerShell như Administrator
winget install Microsoft.DotNet.Runtime.8

# Kiểm tra
dotnet --version
```

### **Bước 2: Clone Repository**

```powershell
mkdir C:\AutoDuLieuCorelV9
cd C:\AutoDuLieuCorelV9

# Clone từ GitHub
git clone https://github.com/thietkehaivuong-creator/AutoDuLieuCorelV9.git .

# HOẶC tải ZIP và giải nén vào C:\AutoDuLieuCorelV9
```

### **Bước 3: Cài Docker**

```powershell
winget install Docker.DockerDesktop

# Khởi động lại PowerShell sau khi cài
docker --version
```

### **Bước 4: Chạy API**

```powershell
cd C:\AutoDuLieuCorelV9

# Khởi chạy
docker-compose up -d

# Chờ 30 giây, kiểm tra
curl http://localhost:5000/swagger
```

### **Bước 5: Cài Macro vào CorelDRAW**

```
1. Mở CorelDRAW
2. Tools → Macros → Edit Macros
3. File → Open → Chọn: Macros/frmPFISearch.bas
4. Ctrl + S
5. Đóng VBA Editor
```

### **Bước 6: Chạy Macro**

```
1. Tools → Macros → Run Macro
2. Chọn: frmPFISearch
3. Click Run
```

---

## 💻 Giao diện sử dụng

```
┌─────────────────────────────────────────────┐
│  AUTO DU LIEU - TIM KIEM PFI                │
├─────────────────────────────────────────────┤
│  Nhap PFI: [24001              ] [TIM]     │
│  Trang thai: San sang                       │
│                                             │
│  ┌─ Du lieu san pham ───────────────────┐  │
│  │ Customer   │ ABC SEAFOOD            │  │
│  │ PFI        │ 24001                  │  │
│  │ Description│ TUNA LOIN              │  │
│  │ Sizes      │ 2KG                    │  │
│  │ Qty        │ 100 CTN                │  │
│  │ FAO        │ 71;77                  │  │
│  │ ...        │ ...                    │  │
│  └────────────────────────────────────┘  │
│                                             │
│  Dinh dang ngay: [dd/mm/yyyy        ]     │
│                                             │
│  [COPY TAT CA] [COPY DONG] [INSERT]       │
│  [CAI DAT] [DONG]                         │
│                                             │
└─────────────────────────────────────────────┘
```

---

## 📝 Hướng dẫn sử dụng

### **1. Tìm kiếm PFI**

```
1. Nhập PFI (Đơn hàng): ví dụ 24001
2. Click "TIM KIEM"
3. Kết quả sẽ hiển thị 20 cột dữ liệu
```

### **2. Chọn định dạng ngày**

```
Các format có sẵn:
✓ dd/mm/yyyy   (02/06/2026)
✓ mm/dd/yyyy   (06/02/2026)
✓ yyyy/mm/dd   (2026/06/02)
✓ mm/yyyy      (06/2026)
✓ yyyy-mm-dd   (2026-06-02)
✓ dd mmm yyyy  (02 Jun 2026)
✓ mmm yyyy     (Jun 2026)
... và nhiều format khác

Hoặc nhập custom format: mm/yyyy, yyyy.mm, ddMMMyyyy
```

### **3. Copy dữ liệu**

```
1️⃣  COPY TAT CA
    → Copy tất cả 20 cột vào Clipboard
    → Dán vào Excel, Word, v.v.

2️⃣  COPY DONG
    → Chọn 1 dòng trong grid
    → Click "COPY DONG"
    → Chỉ copy giá trị của dòng đó

3️⃣  INSERT COREL
    → Tự động thêm Text Box vào CorelDRAW
    → Chứa: PFI, Customer, Product name
```

### **4. Cấu hình Credentials**

```
1. Click "CAI DAT"
2. Mở trình duyệt: http://localhost:5000/swagger
3. Tìm: POST /api/settings/credentials
4. Nhập username & password
5. Click "Execute"

Credentials sẽ được lưu vào:
C:\AutoDuLieuCorelV9\data\config\credentials.json
(File này KHÔNG được commit lên GitHub - bảo mật)
```

---

## 🔗 API Endpoints

Tất cả endpoint có sẵn tại: `http://localhost:5000/swagger`

### **Tìm kiếm**

```
GET /api/data/search/{pfi}
→ Tìm sản phẩm theo PFI

Ví dụ:
GET http://localhost:5000/api/data/search/24001
```

### **Dữ liệu**

```
GET /api/data/all
→ Lấy tất cả dữ liệu sản phẩm
```

### **Định dạng ngày**

```
POST /api/data/format-date
Body: {"inputDate": "02/06/2026", "targetFormat": "dd/mm/yyyy"}
→ Format lại ngày tháng
```

### **FAO**

```
GET /api/data/fao/{code}
→ Tên FAO từ code

Ví dụ:
GET http://localhost:5000/api/data/fao/71
→ Kết quả: "Western Central Pacific Ocean"
```

### **PPDB (Phương pháp đánh bắt)**

```
GET /api/data/ppdb/{code}
→ Tên phương pháp từ code

Ví dụ:
GET http://localhost:5000/api/data/ppdb/PS
→ Kết quả: "Purse seine"
```

### **Cài đặt**

```
GET /api/settings/status
→ Kiểm tra đã cấu hình credentials chưa

POST /api/settings/credentials
Body: {"username": "ma.thach", "password": "Anhthu123!!!"}
→ Lưu thông tin đăng nhập an toàn
```

---

## 🧪 Kiểm tra cài đặt

### **Test 1: API hoạt động**

```powershell
# Mở PowerShell hoặc trình duyệt
curl http://localhost:5000/swagger

# Hoặc mở: http://localhost:5000/swagger
```

### **Test 2: Tìm PFI**

```
Trên Swagger UI:
1. Mở: GET /api/data/search/{pfi}
2. Nhập: 24001
3. Click "Try it out"
4. Kết quả: JSON object với dữ liệu sản phẩm
```

### **Test 3: Macro trong CorelDRAW**

```
1. Mở CorelDRAW
2. Tools → Macros → Run Macro
3. Chọn frmPFISearch
4. Click "Run"
5. Nên hiển thị cửa sổ tìm kiếm
6. Nhập: 24001
7. Click "TIM KIEM"
```

---

## 🐛 Khắc phục sự cố

### **❌ "API không kết nối được"**

```powershell
# 1. Kiểm tra Docker chạy
docker ps

# 2. Kiểm tra logs
docker-compose logs -f api

# 3. Nếu lỗi, reset Docker
docker-compose down -v
docker-compose up -d
```

### **❌ "Macro không hiển thị"**

```
1. Mở CorelDRAW
2. Tools → Macros → Organize Macros
3. Visual Basic → Chọn "All Databases"
4. Tìm frmPFISearch

Nếu không thấy:
  - File → Open → Chọn frmPFISearch.bas
  - Ctrl + S
  - Đóng VBA Editor
  - Mở lại CorelDRAW
```

### **❌ ".NET không cài"**

```powershell
# Kiểm tra
dotnet --version

# Nếu lỗi, cài lại
winget install Microsoft.DotNet.Runtime.8

# Hoặc tải từ: https://dotnet.microsoft.com/download/dotnet/8.0
```

---

## 📁 Cấu trúc thư mục

```
C:\AutoDuLieuCorelV9\
├── .env                              ← Cấu hình môi trường
├── .gitignore                        ← Ignore files
├── docker-compose.yml                ← Docker config
│
├── AutoDuLieuApi\                    ← Backend .NET
│   ├── AutoDuLieuApi.csproj
│   ├── Program.cs
│   ├── appsettings.json
│   ├── Dockerfile
│   ├── Models/
│   ├── Services/
│   └── Controllers/
│
├── Macros\                           ← VBA Macro
│   ├── frmPFISearch.bas
│   └── README.md
│
├── data/
│   └── config/
│       └── credentials.json          ← Lưu thông tin (KHÔNG commit)
│
└── README.md                         ← File này
```

---

## ✅ Checklist hoàn tất

- [ ] Cài .NET 8.0
- [ ] Clone code từ GitHub
- [ ] Cài Docker Desktop
- [ ] `docker-compose up -d`
- [ ] Kiểm tra: http://localhost:5000/swagger
- [ ] Cài Macro vào CorelDRAW
- [ ] Test tìm PFI: 24001
- [ ] Copy dữ liệu thành công
- [ ] Insert vào CorelDRAW thành công

---

## 🎉 Chúc mừng!

Ứng dụng đã sẵn sàng sử dụng. Bất kỳ câu hỏi gì, hãy liên hệ!
