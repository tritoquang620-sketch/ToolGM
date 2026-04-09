# ASN Tool GM

Web tool chạy trên GitHub + Render, tối ưu iPhone/PWA.

## Chức năng
- PDF to IMG: upload tối đa 50 PDF, nhóm CPT / OP / GP, tải ZIP ảnh.
- PDF to EXCEL: xuất file Excel nhiều sheet, tô màu xen kẽ, dòng tổng ASN chuẩn.
- Packing:
  - Mã Đơn: thêm, sửa, xoá, import Excel.
  - Mã Đôi: thêm, sửa, xoá, import Excel.

## Quy tắc tính đang dùng
- CPT: Line No bắt đầu từ C2.
- OP: Line No bắt đầu từ C1.
- GP: Line No là D2 JOB hoặc GP JOB.
- TỔNG ASN / PCS lẻ = số thùng lẻ thực tế.
- Mã đơn có dư: tính 1 thùng lẻ.
- Mã đôi có dư: cả cặp chỉ tính 1 thùng lẻ chung.

## Chạy local
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```

## Import Excel
### Mã Đơn
Cột theo thứ tự:
- Item
- Rev
- Packing

### Mã Đôi
Cột theo thứ tự:
- Item 1
- Rev 1
- Item 2
- Rev 2
- Packing
