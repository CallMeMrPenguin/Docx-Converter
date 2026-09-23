# Future Plan: ULN to DOCX Converter — Subscription & Commercial Release

> Tài liệu tổng hợp toàn bộ thảo luận kiến trúc, bảo mật, và lộ trình thương mại hóa.
> Cập nhật: 2026-09-23

---

## 1. Quyết Định Kiến Trúc: EXE (Không phải Web App)

### Lý do chọn EXE thay vì Web App

| Tiêu chí | Web App | EXE App (Chọn) |
|----------|---------|-----------------|
| Bảo vệ logic | Gần như không thể — F12 DevTools thấy hết | Tốt đến rất tốt với Nuitka/Cython |
| Subscription enforcement | Rất yếu — `window.isLicensed = true` bypass | Mạnh — hardware-bound license |
| Chất lượng DOCX | Giảm (phải dùng docx.js) | Giữ nguyên (pywin32 COM → MS Word) |
| Effort chuyển đổi | Viết lại ~70% renderer | Thêm ~15-20% (login + license) |
| Cross-platform | Mọi OS | Chỉ Windows (đúng target audience) |
| VPS cost | ~$5/tháng | ~$5/tháng |
| Offline support | Không | Có (grace period) |

### Tại sao Web App không khả thi

1. **Core renderer dùng pywin32 COM** — gọi trực tiếp MS Word. Không chạy được trên browser.
2. **Chuyển sang web = viết lại toàn bộ renderer** bằng `docx.js` → chất lượng output giảm đáng kể.
3. **Mọi JavaScript/WASM trên browser đều bị reverse-engineer**:
   - DevTools → Sources tab: xem toàn bộ source code
   - `de4js`, `js-beautifier` deobfuscate trong vài giây
   - WASM cũng bị `wasm-decompile`, `wasm2wat` reverse dễ dàng
   - Network tab: thấy mọi API call, bypass subscription check
4. **Subscription bypass quá dễ** trên browser — copy source tự host, hoặc override biến.

---

## 2. Kiến Trúc Tổng Quan

```
┌─────────────────────────────────────────┐
│              VPS (Lightweight)           │
│  ┌─────────────────────────────────────┐│
│  │ Auth Server (FastAPI/Flask)         ││
│  │ License Validation API              ││
│  │ Subscription Management (Stripe)    ││
│  │ Update Distribution (GitHub Tags)   ││
│  │ PostgreSQL (Users + Licenses)       ││
│  └─────────────────────────────────────┘│
│  CPU: 1 vCPU | RAM: 1GB | Disk: 10GB   │
│  Cost: ~$5-10/month                     │
└─────────────────────────────────────────┘
                    ▲
                    │ HTTPS (License check + Update check)
                    ▼
┌─────────────────────────────────────────┐
│         Desktop EXE (Client)             │
│  ┌─────────────────────────────────────┐│
│  │ Login / Activation UI               ││
│  │ License Token (encrypted, cached)   ││
│  │ ULN Parser (.pyd — Cython native)   ││
│  │ ULN Compiler (.pyd — Cython native) ││
│  │ ULN Renderer (.pyd — Cython native) ││
│  │ Auto-updater (existing updater.py)  ││
│  └─────────────────────────────────────┘│
│  Đóng gói: Nuitka → single .exe binary  │
│  Requires: MS Word installed             │
└─────────────────────────────────────────┘
```

---

## 3. Chiến Lược Bảo Mật (6 Layers)

### Layer 1: Cython Compile Core Modules → .pyd (BẮT BUỘC)

Compile 3 file core thành native binary (`.pyd` = native DLL). Reverse engineering cần chuyên gia và tốn hàng tuần.

```python
# setup_cython.py
from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize([
        "uln_parser.py",      # → uln_parser.pyd
        "uln_renderer.py",    # → uln_renderer.pyd
        "uln_compiler.py",    # → uln_compiler.pyd
    ], compiler_directives={'language_level': "3"})
)

# Build command:
# python setup_cython.py build_ext --inplace
```

**So sánh các cách compile:**

| Cách tạo | Output | Bên trong chứa | Reverse được? |
|-----------|--------|-----------------|---------------|
| PyInstaller `.exe` | `.exe` + `.dll` | Python bytecode (.pyc) | **DỄ** — extract → decompile → source gốc (30 phút) |
| **Cython `.pyd`** | `.pyd` (= `.dll`) | **Native C compiled code** | **Rất khó** — chỉ thấy assembly |
| **Nuitka `.exe`** | `.exe` | **Native C compiled code** | **Rất khó** — tương đương reverse C program |

> **QUAN TRỌNG**: `.dll` từ PyInstaller và `.pyd` từ Cython nhìn ngoài giống nhau nhưng bên trong KHÁC HOÀN TOÀN.
> - PyInstaller DLL = Python bytecode đội lốt DLL → KHÔNG an toàn
> - Cython .pyd = Native C binary đội lốt DLL → AN TOÀN

### Layer 2: Nuitka Đóng Gói Toàn Bộ App (BẮT BUỘC)

```bash
# Nuitka compile toàn bộ app thành 1 native binary
python -m nuitka --standalone --onefile \
  --include-module=uln_parser \
  --include-module=uln_renderer \
  --include-module=uln_compiler \
  --windows-icon=icon.ico \
  gui_app.py
```

### Layer 3: RSA Signed License Token (BẮT BUỘC)

```
❌ SAI: Server trả về { valid: true/false } → client quyết định
   → User patch binary để luôn đọc true

✅ ĐÚNG: Server trả về signed payload + essential data
   → Client KHÔNG THỂ hoạt động nếu không có payload hợp lệ
```

**Thiết kế đúng:**

```
Client → POST /api/validate { machine_id, license_key }
Server → Response:
  {
    "token": "eyJhbGciOi...",           // JWT signed bởi server PRIVATE KEY
    "expires_at": "2026-10-23T...",      // Hết hạn sau 72h
    "feature_flags": { ... },            // Tính năng được bật
    "signature": "RSA_SIGN(payload)"     // Chữ ký RSA-2048
  }

Client verify signature bằng PUBLIC KEY embedded trong binary.
Nếu token hết hạn → PHẢI gọi server lại.
Nếu không có token hợp lệ → app KHÔNG compile được.
```

### Layer 4: Gắn Logic Quan Trọng Vào License Token (NÊN LÀM)

```python
# ❌ SAI — License check tách biệt khỏi core logic
if not is_licensed():
    show_error("Vui lòng mua license")
    return
# ... compiler chạy bình thường ...

# ✅ ĐÚNG — Core logic PHỤ THUỘC vào license token
def compile(self, uln_text, license_token):
    config = decrypt_compile_config(license_token)  # Lấy config từ token
    if not config:
        raise RuntimeError("Invalid license")
    # Config chứa rendering parameters THIẾT YẾU
    # Không có token → không có config → không compile được đúng
```

### Layer 5: Hardware Binding (BẮT BUỘC)

```python
import hashlib, uuid, platform, subprocess

def get_machine_id():
    factors = []
    factors.append(str(uuid.getnode()))                    # MAC address
    factors.append(platform.processor())                   # CPU info
    
    # Windows disk serial
    result = subprocess.run(
        ['wmic', 'diskdrive', 'get', 'serialnumber'],
        capture_output=True, text=True
    )
    factors.append(result.stdout.strip())
    
    # Kết hợp và hash
    raw = "|".join(factors)
    return hashlib.sha256(raw.encode()).hexdigest()[:32]
```

**Hiệu quả**: 1 license key chỉ chạy trên 1-2 máy cụ thể. Share key → không dùng được.

### Layer 6: Anti-Tampering & Periodic Validation (NÊN LÀM)

```python
# Kiểm tra integrity của chính binary
import hashlib, sys

def self_integrity_check():
    exe_path = sys.executable
    with open(exe_path, 'rb') as f:
        current_hash = hashlib.sha256(f.read()).hexdigest()
    
    # So sánh với hash gốc (lấy từ server khi validate license)
    if current_hash != expected_hash_from_server:
        sys.exit(1)  # Binary đã bị patch/modify
```

**Periodic validation**: Heartbeat check mỗi 24-72h. Offline grace period 7 ngày.

---

## 4. Đánh Giá Mức Bảo Mật Theo Cấu Hình

| Cấu hình | Effort | Bảo mật | Ai crack được? |
|-----------|--------|---------|----------------|
| PyInstaller + if/else check | Thấp | **Yếu** | Bất kỳ dev nào biết Google |
| PyInstaller + PyArmor | Trung bình | **Trung bình** | RE có kinh nghiệm |
| **Nuitka + RSA license + hardware binding** | **Cao** | **Mạnh** | **Chỉ chuyên gia RE** |
| Nuitka + Cython .pyd + server-dependent logic + anti-tamper | Rất cao | **Rất mạnh** | Gần như không đáng crack |

---

## 5. Lộ Trình Triển Khai

### Phase 1: VPS Backend (1-2 tuần)

```
├── FastAPI auth server
│   ├── POST /api/register (đăng ký tài khoản)
│   ├── POST /api/login (đăng nhập, trả JWT)
│   ├── POST /api/validate-license (validate + machine_id)
│   ├── POST /api/activate (kích hoạt license key)
│   └── GET  /api/check-update (kiểm tra phiên bản mới)
│
├── PostgreSQL Schema
│   ├── users (id, email, password_hash, created_at)
│   ├── subscriptions (id, user_id, plan, status, expires_at)
│   ├── licenses (id, user_id, license_key, machine_id, activated_at)
│   └── login_sessions (id, user_id, token, ip, created_at)
│
├── Stripe Integration
│   ├── Checkout Session (tạo payment link)
│   ├── Webhook handler (payment success → activate subscription)
│   └── Subscription lifecycle (create, renew, cancel, expire)
│
└── Admin Dashboard (đơn giản)
    ├── Xem danh sách users
    ├── Xem subscriptions
    ├── Revoke/ban license
    └── Usage analytics
```

### Phase 2: Desktop Client Integration (1-2 tuần)

```
├── Login/Register Dialog (tkinter)
│   ├── Email + Password login form
│   ├── Register new account
│   ├── "Remember me" checkbox
│   └── License activation input
│
├── License Token Manager
│   ├── Encrypted local cache (AES-256)
│   ├── Auto-refresh khi token sắp hết hạn
│   ├── Offline grace period (7 ngày)
│   └── Machine fingerprint generation
│
├── App Integration
│   ├── License check trước khi compile
│   ├── Subscription status badge trên header
│   ├── Feature gating (free vs premium)
│   └── Graceful degradation khi offline
│
└── Update existing updater.py
    ├── Tích hợp license validation vào update flow
    └── Signed update packages
```

### Phase 3: Build Pipeline & Hardening (1 tuần)

```
├── Cython Compilation
│   ├── setup_cython.py cho 3 core modules
│   ├── Build script tự động (.pyd output)
│   └── Test .pyd modules hoạt động đúng
│
├── Nuitka Packaging
│   ├── Nuitka build script (gui_app.py → .exe)
│   ├── Include .pyd modules
│   ├── Windows icon + metadata
│   └── Code signing certificate (optional)
│
├── Anti-Tampering
│   ├── Binary integrity check
│   ├── Certificate pinning cho HTTPS
│   └── Debug detection (anti-debugger)
│
└── Installer
    ├── Rebuild installer với license flow
    ├── First-run activation wizard
    └── Uninstaller (clean license cache)
```

### Phase 4: Testing & Launch (1 tuần)

```
├── Security Testing
│   ├── Tự thử reverse Nuitka binary
│   ├── Tự thử bypass license
│   ├── Test offline grace period
│   └── Test hardware binding (đổi máy)
│
├── User Flow Testing
│   ├── Register → Pay → Activate → Use
│   ├── License expire → Renew
│   ├── Transfer license (đổi máy)
│   └── Update app (giữ license)
│
└── Soft Launch
    ├── Beta test với 5-10 users
    ├── Monitor server logs
    ├── Fix edge cases
    └── Public release
```

---

## 6. Chi Phí Ước Tính

| Hạng mục | Chi phí/tháng |
|----------|---------------|
| VPS (1 vCPU, 1GB RAM) | $5-10 |
| Domain + SSL | $1-2 |
| Stripe fee | 2.9% + $0.30/transaction |
| Code signing cert (optional) | ~$70-200/năm |
| **Tổng cố định** | **~$7-12/tháng** |

---

## 7. Ghi Chú Kỹ Thuật Quan Trọng

1. **KHÔNG BAO GIỜ** dùng PyInstaller đơn thuần cho bản thương mại — extract source trong 30 phút.
2. **KHÔNG BAO GIỜ** để server chỉ trả `{ valid: true/false }` — phải dùng RSA signed token.
3. **LUÔN** gắn license vào machine ID — ngăn share key.
4. **LUÔN** có offline grace period — user mất mạng vẫn dùng được vài ngày.
5. **LUÔN** test tự crack app của mình trước khi release — nếu mình crack được thì người khác cũng crack được.
6. Cython `.pyd` và Nuitka binary là **2 layer bảo vệ độc lập** — nên dùng cả 2.
7. Certificate pinning ngăn MITM attack — user không thể dùng proxy để fake server response.
