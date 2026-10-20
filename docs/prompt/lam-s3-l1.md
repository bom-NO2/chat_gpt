# Nhật ký AI – Lâm – S3-L1

## 1. Công cụ AI

OpenAI Codex trong môi trường cloud; Python/python-docx để soạn Word; Graphviz để sinh ERD SVG/PNG từ mã DOT. Không tự ghi tên/phiên bản mô hình khi không có thông tin xác nhận. Ngày lập: 21/10/2026, theo ngày/múi giờ người dùng (Asia/Saigon).

## 2. Prompt nguyên văn

```text
Mô tả: Thiết kế theo 4 bước như S4-L1 cho các bảng: TaiKhoan, GioHang, DonHang, ChiTietDonHang, DiaChiKhachHang, BangPhiGiao, CauHinhCuaHang.
Đầu vào (cần có trước khi làm): Bản tham khảo: WebsiteBanSach_Sprint3/database/01_tao_bang.sql (phần tài khoản, giỏ, đơn); WebsiteBanSach_Sprint3/docs/dac_ta_sprint3_gio_hang_dat_hang.md mục 4; tài liệu 4 bước của S4-L1.
Đầu ra (phải nộp): Tài liệu 4 bước và ERD phần Sprint 3, nối với ERD quản lý sách.
Điều kiện đạt: Giải thích được vì sao mỗi bảng tồn tại; đơn hàng lưu giá lúc đặt, phí giao, Tỉnh/Thành, Phường/Xã; mật khẩu chỉ lưu bản băm.
Phụ thuộc: S3-H1 (đặc tả UC15, UC26 – Hùng), S3-Y1 (đặc tả UC03, UC04, UC05 – Huy)
Nhật ký AI: file docs/prompt/lam-s3-l1.md trong repo (công cụ AI, prompt nguyên văn, phần giữ lại, phần tự sửa và lý do). Thiếu file này thì thẻ không được Pass.
```

## 3. Nguồn đã có và nguồn chưa có

- Đã đọc SQL master trong WebsiteBanSach-S4-L4.zip được cung cấp ở các lượt trước, cùng model, báo cáo chung và mẫu 06. SQL nền có SHA-256 `5cb74aa1e7441b4653279ab3d2c07dea78bcb1a77fa1ff2c4491ec08d226f601`.
- Không có đúng đường dẫn/bản WebsiteBanSach_Sprint3 hoặc file dac_ta_sprint3_gio_hang_dat_hang.md mục 4 trong hồ sơ hiện tại. Không coi tên CSDL NhaSachOnline_S3 trong SQL master là bằng chứng đã có đúng bản tham khảo.
- Chưa có tài liệu S4-L1 riêng và đặc tả S3-H1/S3-Y1 cuối để xác nhận mọi quyết định. Các khác biệt và giả định cần chốt được nêu ngay đầu/cuối tài liệu.
- Không truy cập hoặc sửa Trello, không có bình luận review/Pass của người kiểm tra.

## 4. Phần giữ lại

- Bốn bước: thực thể/khóa/thuộc tính; quan hệ; chuẩn hóa; từ điển dữ liệu và ERD.
- 7 bảng theo thẻ; bổ sung trên ERD những bảng có sẵn cần nối: KhachHang, Sach, NhanVien và nhóm danh mục/tác giả.
- Khóa giỏ (MaKhachHang,MaSach), khóa chi tiết (MaDon,MaSach), khóa chính và ràng buộc từ SQL nguồn.
- MatKhauHash là chuỗi hash, không cột mật khẩu rõ.
- DonGia lịch sử ở chi tiết, PhiGiao và địa chỉ nhận độc lập ở đơn. Tổng tiền dùng giá/phí chốt, không JOIN giá hiện hành để tính lại.
- Các tuyến/bậc phí theo cấu trúc SQL nguồn; không tạo mức tiền thật hoặc mã hành chính giả.

## 5. Phần AI điều chỉnh và lý do

- Phân biệt tên logic ChiTietDonHang theo thẻ và tên vật lý ChiTietDon trong nguồn; ghi rõ cần chọn ánh xạ hoặc migration, không tạo hai bảng song song.
- Đề xuất PhuongXa ở sổ địa chỉ và bản sao trên đơn; tăng độ dài TinhThanh; đề xuất TuyenGiao/KhoiLuongTinhPhiGram trên đơn để giải thích phí đã chốt.
- Đề xuất CauHinhCuaHang dạng một cửa hàng, điểm gửi và khối lượng đóng gói; CHECK MaCauHinh=1 chỉ giới hạn tối đa một dòng, phải có bước khởi tạo.
- Tách luồng sao chép địa chỉ/đọc cấu hình bằng nét đứt khỏi FK. Đơn không phụ thuộc MaDiaChi để đọc địa chỉ lịch sử.
- Ghi chính xác quan hệ: tài khoản có thể không có hồ sơ khách; mỗi đơn phải có ít nhất một dòng theo nghiệp vụ nhưng FK không tự ngăn đơn rỗng; người nội bộ cập nhật có thể NULL.
- Không khẳng định snapshot giá/địa chỉ vi phạm 3NF: phụ thuộc của chúng là khóa giao dịch tại thời điểm phát sinh.
- Đề xuất unique index lọc cho địa chỉ mặc định, FK NO ACTION cho lịch sử đơn/sách và giao dịch nguyên tử khi đặt hàng. Không khẳng định giao dịch một mình chống hai POST lặp; cần quyết định idempotency riêng nếu đặc tả yêu cầu.
- Không bịa lịch sử Phường/Xã của đơn cũ; có phương án chuyển tiếp NULL khi nâng cấp trước khi đủ dữ liệu để ép NOT NULL.
- Sinh Markdown, Word và hai ERD từ một nguồn từ điển; kèm JSON và DOT để tái tạo. Word dùng trang ngang cho sơ đồ, không cắt mất trường/quan hệ.

## 6. Phần Lâm tự sửa và lý do

Chưa có thay đổi thủ công của Lâm được cung cấp. Không ghi nhận các điều chỉnh của AI là phần Lâm tự sửa. Sau khi Lâm đối chiếu S3-H1/S3-Y1 và nguồn Sprint 3 thật, bổ sung mục này với thay đổi và lý do thực tế.

## 7. Kiểm tra và trạng thái

- Đã kiểm tra sự tồn tại/khớp kiểu của 8 FK trong bảy bảng và liên kết KhachHang.MaTaiKhoan tới TaiKhoan; kiểm tra tài liệu có đủ 7 bảng, 4 bước, 2 hình và 5 ý/hình.
- Đã sinh thành công hai SVG/PNG bằng Graphviz và kiểm tra cấu trúc DOCX. Các kiểm tra này là kiểm tra tài liệu/mô hình, không phải chạy schema hoặc ứng dụng trên SQL Server.
- Không sửa, chạy hoặc thêm DROP/ALTER/CREATE vào database của nhóm; không thực thi SQL minh họa trong tài liệu.
- Các ca TK/GH/DH/DC/PG/CH là tiêu chí nghiệm thu thiết kế cần chạy khi triển khai, không ghi Pass.
- Trạng thái: hoàn thành bộ tài liệu thiết kế đề xuất dựa trên nguồn hiện có; còn cần đối chiếu bản tham khảo Sprint 3, đặc tả phụ thuộc và review của nhóm. Lưu ở nhánh bàn giao được cho phép trên bom-NO2/chat_gpt, không coi là PR đã merge vào repository WebsiteBanSach.
