# 🎓 Ứng dụng Xử lý Bảng Điểm Tổng Hợp

## Cấu trúc thư mục

```
bang_diem_app/
├── app.py                         ← code chính
├── requirements.txt
├── README.md
└── data/                          ← đặt các file KetQua vào đây
    ├── KetQua_BSNT_Ngoai.xlsx
    ├── KetQua_BSNT_San.xlsx
    ├── KetQua_BSNT_Nhi.xlsx
    ├── KetQua_BSNT_TMH.xlsx       ← đã có
    ├── KetQua_CKI_Ngoai.xlsx
    ├── KetQua_CKI_San.xlsx
    ├── KetQua_CKI_Nhi.xlsx
    ├── KetQua_CKI_TMH.xlsx
    ├── KetQua_CKI_Noi.xlsx
    └── KetQua_CKI_NgoaiTimMach.xlsx
```

### File KetQua_<chuyên ngành>.xlsx là gì?
File do bạn tự tạo bằng notebook `merge_chon_sheet_colab.ipynb` (đã có từ trước).
Cột cần có: `Tên môn học`, `MSMH`, `Trọng số điểm Thường xuyên`, `Trọng số điểm GK`, `Trọng số điểm CK`, `Điểm đạt`

---

## Chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
# Mở: http://localhost:8501
```

---

## Deploy Streamlit Cloud (miễn phí, 5 phút)

1. **Tạo repo GitHub** → upload toàn bộ thư mục `bang_diem_app/`
   > ⚠ Bao gồm thư mục `data/` với đủ các file KetQua

2. Vào **[share.streamlit.io](https://share.streamlit.io)** → đăng nhập GitHub
3. **New app** → chọn repo → Main file: `app.py` → **Deploy**
4. URL dạng: `https://<tên-repo>.streamlit.app`

---

## Thêm chuyên ngành mới

1. Tạo file `KetQua_<tên>.xlsx` bằng notebook `merge_chon_sheet_colab.ipynb`
2. Đặt vào thư mục `data/`
3. Thêm vào dict `CHUYEN_NGANH` trong `app.py`:
```python
"Tên hiển thị": "KetQua_<tên>.xlsx",
```

---

## Cách dùng app

| Bước | Thao tác |
|------|---------|
| 1 | Upload 1 hoặc nhiều file Excel bảng điểm |
| 2 | Chọn chuyên ngành từ dropdown |
| 3 | Nhấn **Xử lý ngay** |
| 4 | Xem preview → **Tải file Excel** |

**File kết quả có 3 sheet:**
| Sheet | Nội dung |
|-------|---------|
| `Daydu` | Toàn bộ thông tin từng học viên + trọng số điểm |
| `LopHocPhan` | Danh sách lớp học phần (unique) + cột Rớt môn |
| `SV_LopHocPhan` | MSSV, Họ tên, Mã lớp học phần |
