# Nhật ký AI – Lâm – S4-03 – Báo cáo sản phẩm cá nhân Sprint 4

## 1. Công cụ AI

OpenAI Codex trong phiên cloud, hỗ trợ đọc hồ sơ, đối chiếu mã/bằng chứng và soạn Word bằng Python python-docx. Không xác nhận phiên bản mô hình cụ thể vì phiên không cung cấp tên model chính xác. Thời điểm lập bản nháp: 19/10/2026, theo ngày của người dùng (Asia/Saigon).

## 2. Prompt nguyên văn

```text
 viết giúp tôi
Theo buổi 5.1, báo cáo gồm sản phẩm của từng thành viên và bản chung do nhóm trưởng gộp. Viết báo cáo sản phẩm các thẻ mình phụ trách theo mẫu 06 (8 mục trùng báo cáo chung); mục không phụ trách ghi “Không phụ trách”. Mục cần viết: 4 (S4-L1), 5 (S4-L3), 6 (S4-L2, S4-L4), 8 (quá trình).
Đầu vào (cần có trước khi làm): Sản phẩm các thẻ của mình (sơ đồ, màn hình, code, phiếu kiểm thử), thẻ Trello, nhật ký AI, commit trên GitHub; mẫu “06. Mẫu báo cáo sản phẩm cá nhân” trong Sản phẩm yêu cầu/Mẫu biểu.
Đầu ra (phải nộp): File “Báo cáo sản phẩm cá nhân Sprint 4.docx” trong Sản phẩm yêu cầu/Lê Hoàng Lâm.
Điều kiện đạt: Mỗi thẻ có sản phẩm trình bày ở đúng mục; mỗi hình ghi 5 ý; code có trích đoạn chính, giải thích và link commit; mục 8 có giờ khớp bình luận trên thẻ.
Phụ thuộc: Các thẻ của mình
Nhật ký AI: file docs/prompt/lam-s4-03.md trong repo (công cụ AI, prompt nguyên văn, phần giữ lại, phần tự sửa và lý do). Thiếu file này thì thẻ không được Pass.
```

## 3. Đầu vào đã kiểm tra

- tonkho-lam.zip: bản tham khảo và S4-L2 v3, nhật ký lam-s4-l2.md, SQL/Models/Program.
- WebsiteBanSach-master.zip và ZIP Huy/hung/huyen/kiet: mã nhóm; không thấy mẫu 06, sản phẩm L1/L3 hoặc lịch sử Trello của Lâm.
- Sản phẩm S4-L4 đã tích hợp trong phiên: code, log build Release, 38 kiểm tra HTTP và 10 Chromium với EF InMemory, ảnh bằng chứng.
- Commit bàn giao tệp thực tế trên bom-NO2/chat_gpt: e9fef93. Không coi đây là commit cá nhân của Lâm trên kho WebsiteBanSach.
- Đã hỏi thông tin còn thiếu để hoàn thiện đúng mẫu và hồ sơ.

## 4. Phần giữ lại trong bản nháp

- Cấu trúc tám vị trí, các vị trí 1/2/3/7 ghi Không phụ trách theo yêu cầu người dùng. Chưa có tên mục chính thức nên không tự khẳng định đúng mẫu 06.
- Mục 6 có mô tả sản phẩm L2/L4, trích đoạn chính, giải thích, bốn hình có năm ý mô tả tạm, kết quả test và giới hạn.
- Mục 4/5 có chỗ để nhận sản phẩm thật, không tạo sơ đồ hoặc đặc tả rồi nhận là sản phẩm Lâm đã nộp.
- Mục 8 có bảng giờ chờ đối chiếu; không lấy timestamp của AI hoặc commit để thay giờ Trello.
- Cảnh báo bản nháp, thiếu mẫu/phiếu/commit/Trello, khác biệt tên CSDL S3/S4 và việc chưa xác minh SQL Server thật.
- File Word nằm đúng đường dẫn yêu cầu trong bản sao bàn giao: Sản phẩm yêu cầu/Lê Hoàng Lâm/Báo cáo sản phẩm cá nhân Sprint 4.docx.

## 5. Phần tự sửa và lý do

### Điều chỉnh do AI sau đối chiếu

- Tách bằng chứng thực tế của phiên khỏi các kết quả test cũ và ảnh mô phỏng trong tài liệu đầu vào.
- Dẫn link commit bàn giao ZIP có thật, ghi rõ chưa có commit/PR code trên repository nhóm; không tạo link giả từ SHA cục bộ.
- Giới hạn vùng hiển thị hai ảnh danh sách rất dài trong Word để tránh ảnh vượt trang; ảnh gốc vẫn được nhúng nguyên vẹn và bản đầy đủ có trong gói mã.
- Dùng 5 ý tạm cho mỗi hình, ghi rõ cần đối chiếu mẫu chính thức để không trình bày đó là yêu cầu đã xác nhận.

### Phần Lâm tự sửa

Chưa có nội dung chỉnh sửa thủ công của Lâm được cung cấp. Lâm bổ sung phần này sau khi đọc, sửa và đối chiếu thực tế, nêu rõ nội dung và lý do; không nhận sửa của AI thành sửa cá nhân.

## 6. Trạng thái sau lượt đầu

Bản nháp, chưa đủ điều kiện Pass hoặc nộp cuối. Cần mẫu 06, sản phẩm S4-L1/L3, mô tả/bình luận Trello kèm giờ, phiếu kiểm thử và commit/PR cá nhân thực tế. Hồ sơ được lưu trên nhánh bàn giao đã được người dùng cho phép sử dụng, không thay main hoặc đẩy vào repository nhóm khi chưa có quyền.

## 7. Lượt cập nhật theo phạm vi hai ZIP người dùng chỉ định

### Prompt nguyên văn

```text
viết báo cáo dựa trên mấy cái này là đc
```

Tệp người dùng gửi kèm: tonkho-lam.zip và WebsiteBanSach-S4-L4.zip.

### Kiểm tra nguồn

- tonkho-lam.zip: SHA-256 `9095ba362753bd5e5e0231df0cb5df98f272be72fd7d5de2c24029fdaa6ff544`.
- WebsiteBanSach-S4-L4.zip: SHA-256 `dd7abab39ad748f37d9c40449064509a9f3030bac292ce2a365a64a8405fc56b`, khớp gói tích hợp master đã bàn giao.
- Hai ZIP không chứa mẫu 06, sản phẩm mang mã S4-L1/S4-L3, lịch sử giờ Trello hoặc phiếu kiểm thử có người xác nhận.

### Phần giữ lại

Mục 6 được viết từ SQL/Model/kết nối L2 và code/bằng chứng L4 có thật. Giữ bốn hình có năm ý, trích đoạn code và giải thích, kết quả 38 HTTP + 10 Chromium, giới hạn EF InMemory và link commit bàn giao ZIP có thật. Không yêu cầu thêm tài liệu để tiếp tục phạm vi người dùng đã chọn.

### Điều chỉnh do AI và lý do

- Viết lại báo cáo thành bản tổng hợp dựa trên hai ZIP, bỏ nhãn bản nháp và các ô chờ điền trên bìa.
- Giữ đủ tám mục; mục 1/2/3/7 ghi Không phụ trách. Tên mục đặt theo nội dung vì không có mẫu gốc, không khẳng định trùng tên mẫu 06.
- Mục 4/5 giải thích nguồn không có sản phẩm L1/L3, không coi đó là mục không phụ trách hoặc tự vẽ sản phẩm để lấp chỗ thiếu.
- Viết lại mục 8 thành các bước thực hiện, phối hợp, khó khăn, cách xử lý và sử dụng AI. Cột giờ ghi Không có dữ liệu, không lấy timestamp công cụ làm giờ Trello.
- Phân biệt CSDL S4 trong gói L2 v3 với S3 ở master, ảnh test host với ảnh mô phỏng, commit bàn giao với commit cá nhân trên repository nhóm.

### Phần Lâm tự sửa

Chưa có thay đổi thủ công do Lâm cung cấp; không ghi nhận sửa của AI thành phần Lâm tự sửa.

### Trạng thái sau cập nhật

Đã soạn và lưu báo cáo theo phạm vi hai ZIP được người dùng chấp nhận, đúng đường dẫn tên file yêu cầu. Không kết luận đã đạt mọi điều kiện thẻ: mẫu chính thức, sản phẩm L1/L3, giờ Trello, kiểm thử SQL Server và commit/PR cá nhân vẫn không có bằng chứng trong hồ sơ. Các giới hạn đã được ghi trong báo cáo.
