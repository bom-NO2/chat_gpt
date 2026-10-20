# S3-L1 – Thiết kế CSDL Sprint 3

**Thiết kế đề xuất để review**, không phải migration hoặc chức năng đã chạy.

- Tài liệu chính: [4 bước](Thiet-ke-CSDL-Sprint3-4-buoc.md).
- [ERD nối nhóm sách (SVG)](erd-noi-voi-quan-ly-sach.svg), [PNG](erd-noi-voi-quan-ly-sach.png), [mã DOT](erd-noi-voi-quan-ly-sach.dot).
- [ERD chi tiết Sprint 3 (SVG)](erd-chi-tiet-sprint3.svg), [PNG](erd-chi-tiet-sprint3.png), [mã DOT](erd-chi-tiet-sprint3.dot).
- [Từ điển JSON](schema-de-xuat.json).
- Word ở `Sản phẩm yêu cầu/Lê Hoàng Lâm/S3-L1/` trong repo.
- Nhật ký AI: `docs/prompt/lam-s3-l1.md`.

Đủ bảy bảng theo thẻ; giữ giá/phí và địa chỉ lịch sử; mật khẩu chỉ có hash. Nguồn hiện có là SQL master S4-L4, chưa có đúng bản tham khảo WebsiteBanSach_Sprint3 và đặc tả phụ thuộc cuối. Mọi phần mới và khác tên được nêu rõ trong tài liệu.

**ChiTietDonHang (*)** trên sơ đồ là tên logic theo thẻ; nguồn vật lý là ChiTietDon. Nét liền: PK–FK; nét đứt: đọc/sao chép lúc chốt đơn, không là FK. DonHang–chi tiết tối thiểu một dòng là ràng buộc nghiệp vụ cần giao dịch, không chỉ FK. NguoiCapNhat có thể NULL.

## Tái tạo

Cài Python với `python-docx` và Graphviz `dot`, rồi chạy từ gốc repo:

```sh
python docs/thiet-ke-sprint3/generate_design.py
```

Lệnh chỉ ghi tệp tài liệu/sơ đồ, không kết nối hoặc thay CSDL. Các ví dụ SQL trong tài liệu chưa được thực thi; không chạy trên CSDL nhóm khi chưa rà lại tên vật lý/migration.
