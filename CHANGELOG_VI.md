Nhật kí thay đổi



Phiên bản 0.6.0 – 2025-01-20

Tính năng mới

• Đã thêm chức năng kiểm tra chính tả. Từ menu ngữ cảnh, người dùng có thể kiểm tra xem từ hiện tại có đúng hay không và nhận gợi ý nếu sai.

• Đã thêm chức năng nhập và xuất podcast thông qua tệp OPML.

• Đã thêm hỗ trợ tìm kiếm Podcast Index bên cạnh iTunes. Người dùng có thể nhập API key và API secret miễn phí (được tạo chỉ bằng địa chỉ email).

• Đã thêm hỗ trợ giọng nói SAPI4, cho cả việc đọc thời gian thực và tạo sách nói.

• Đã thêm cơ chế OCR dự phòng tự động cho các PDF không truy cập được: khi không tìm thấy văn bản có thể trích xuất, tài liệu sẽ được nhận dạng bằng OCR.

• Đã thêm hỗ trợ từ điển thông qua Wiktionary. Khi nhấn phím Applications, các định nghĩa sẽ được hiển thị và, khi có sẵn, cả từ đồng nghĩa và bản dịch sang ngôn ngữ khác.

• Đã thêm chức năng nhập bài viết từ Wikipedia với tìm kiếm, chọn kết quả và nhập trực tiếp vào trình soạn thảo.

• Đã thêm phím tắt Shift+Enter trong mô-đun RSS để mở bài viết trực tiếp trên trang web gốc.

Cải tiến

• Việc chọn micro hiện luôn được ứng dụng tôn trọng.

• Trong cửa sổ podcast, khi nhấn Enter trên một tập, NVDA sẽ thông báo ngay “đang tải”, giúp người dùng biết rõ thao tác đã được thực hiện.

• Trong kết quả tìm kiếm podcast, nhấn Enter sẽ đăng ký podcast đã chọn.

• Đã sửa và cải thiện nhãn cho các phím tắt Ctrl+Shift+O và Podcast Ctrl+Shift+P.

• Tốc độ phát và âm lượng hiện được lưu trong cài đặt và áp dụng cho tất cả các tệp âm thanh.

• Đã thêm thư mục bộ nhớ đệm riêng cho các tập podcast. Người dùng có thể giữ tập podcast thông qua mục “Giữ podcast” trong menu Phát. Bộ nhớ đệm sẽ tự động được dọn dẹp khi vượt quá dung lượng do người dùng thiết lập (Tùy chọn → Âm thanh).

• Đã cải thiện đáng kể việc tải bài viết RSS bằng cách sử dụng libcurl với giả lập Chrome và iPhone, đảm bảo khả năng tương thích với khoảng 99% trang web.

• Đã thêm trạng thái đã đọc / chưa đọc cho các bài viết RSS, với hiển thị rõ ràng trong danh sách RSS.

• Chức năng Thay thế tất cả hiện hiển thị số lượng thay thế đã thực hiện.

• Đã thêm nút Xóa podcast khi duyệt thư viện podcast bằng phím Tab.

Sửa lỗi

• Đã loại bỏ mục “pending update” dư thừa trong menu Trợ giúp (việc cập nhật đã được xử lý tự động).

• Đã sửa lỗi khi mở tệp MP3 và nhấn Ctrl+S khiến tệp bị lưu và dẫn đến hỏng tệp.

• Đã sửa lỗi giao diện khi “Batch Audiobooks” hiển thị là “(B)… Ctrl+Shift+B” (đã loại bỏ nhãn dư thừa).

• Đã sửa lỗi dấu ngoặc kép thông minh: khi được bật, dấu ngoặc kép thông thường giờ đây được thay thế chính xác bằng dấu ngoặc kép kiểu chữ.

• Đã sửa lỗi khi sử dụng “Đi tới dấu trang” làm tốc độ phát bị đặt lại về 1.0.

• Đã sửa lỗi khi phát podcast đã tải xuống nhưng hệ thống vẫn tải lại thay vì sử dụng phiên bản trong bộ nhớ đệm.

Phím tắt

• F1 hiện mở hướng dẫn.

• F2 hiện kiểm tra cập nhật.

• F7 / F8 hiện cho phép chuyển đến lỗi chính tả trước đó hoặc tiếp theo.

• F9 / F10 hiện cho phép chuyển nhanh giữa các giọng nói đã lưu trong mục yêu thích.

Cải tiến dành cho nhà phát triển

• Lỗi không còn bị bỏ qua một cách im lặng: tất cả các mẫu let \_ = đã được loại bỏ và lỗi hiện được xử lý rõ ràng (truyền tiếp, ghi log hoặc xử lý bằng cơ chế dự phòng phù hợp).

• Dự án hiện sẽ không biên dịch nếu còn cảnh báo: cả cargo check và cargo clippy đều phải chạy sạch, với các lint chặt chẽ hơn và loại bỏ allow khi có thể.

• Đã loại bỏ các triển khai tùy chỉnh kiểu strlen / wcslen. Độ dài chuỗi và bộ đệm UTF-16 hiện được lấy trực tiếp từ dữ liệu do Rust quản lý, thay vì quét bộ nhớ thủ công.

• Việc xử lý DLL đã được làm gọn và thống nhất xoay quanh libloading, tránh logic tải tùy chỉnh và phân tích PE.

• Đã loại bỏ các helper tự viết cho việc phân tích byte; toàn bộ việc phân tích byte hiện sử dụng from\_le\_bytes / from\_be\_bytes trên các slice đã được kiểm tra.

Những thay đổi này giúp giảm việc sử dụng unsafe không cần thiết, loại bỏ các hành vi không xác định tiềm ẩn và làm cho mã nguồn trở nên chuẩn mực hơn, ổn định hơn và dễ bảo trì hơn.

 



Phien ban 0.5.9 - 2025-01-13
Tinh nang moi
• Them kha nang sap xep RSS tu menu ngu canh (len/xuong/den vi tri), co kiem tra vi tri khong hop le.
• Them menu ngu canh cho bai viet: mo trang goc va chia se qua WhatsApp, Facebook va X.
• Them phim tat Esc de quay lai danh sach RSS tu bai viet da nhap.
• Them che do podcast: tim kiem, dang ky va nghe; sap xep dang ky; Esc dung phat va quay lai danh sach; Enter tren tap phat bat dau phat.
• Them dieu chinh toc do phat cho podcast va file MP3.
• Them Ctrl+T de di den thoi gian cu the.
• Them nut nghe thu giong doc sau hop chon am luong.
• Them chuc nang regex cho Tim va Thay the, phong cach Notepad++.
• Them nhap RSS tu file OPML va TXT.
• Them tuy chon trong Cai dat de bat "Mo voi Novapad" trong File Explorer, ke ca ban portable.
Cai tien
• Cai thien lua chon toc do, cao do va am luong giong doc, ton trong gioi han toi da cua TTS.
• Nhieu cai tien cho RSS de tai du cac bai viet ma khong di chuyen focus NVDA khi cap nhat.
• Cai thien phat am thanh voi menu rieng, thong bao thoi gian bang Ctrl+I va am luong toi 300%.
• Them cac phim tat con thieu cho mot so chuc nang.
• Sap xep lai menu Chinh sua voi submenu cho cac chuc nang lam sach van ban.
• Sap xep lai Tuy chon thanh cac the, co Ctrl+Tab va Ctrl+Shift+Tab de chuyen the.
• Khac phuc loi doc bai viet: trinh doc RSS gio doc du bai viet nhu tren trinh duyet.
Sua loi
• Sua loi lam sach Markdown loai bo so o dau dong.
• Sua AltGr+Z kich hoat Undo.
• Sua loi khi ghi sach noi khong the dung nhanh.
Ban dia hoa
• Them dich tieng Viet (cam on Anh Duc Nguyen).

Phiên bản 0.5.7 - 2026-01-05
Tính năng mới
• Thêm tính năng Sách nói hàng loạt để chuyển đổi nhiều tệp/thư mục cùng lúc.
• Thêm hỗ trợ cho các tệp Markdown (.md).
• Thêm lựa chọn bảng mã khi mở các tệp văn bản.
• Thêm tùy chọn trong terminal hỗ trợ tiếp cận để thông báo khi có dòng mới bằng NVDA.
Cải tiến
• Quá trình ghi âm sách nói giờ đây sẽ lưu trực tiếp sang định dạng MP3 khi được chọn.
• Người dùng có thể chọn vị trí của dấu sao (\*) báo hiệu "chưa lưu" trên tiêu đề cửa sổ.
• Cải thiện độ ổn định của hệ thống cập nhật trong nhiều tình huống khác nhau.
• Thêm mục "Xóa dấu gạch nối" trong menu Chỉnh sửa để sửa lỗi ngắt dòng sau khi quét OCR.

Phiên bản 0.5.6 - 2026-01-04
Sửa lỗi
Cải thiện Tìm trong các tệp để khi nhấn Enter sẽ mở tệp chính xác tại vị trí đoạn văn bản được chọn.
Cải tiến
Thêm hỗ trợ định dạng PPT/PPTX (mở dưới dạng văn bản).
Khi mở các định dạng không phải văn bản thuần túy, phần mềm sẽ lưu thành tệp .txt để tránh lỗi định dạng (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
Thêm ghi âm podcast từ micro và âm thanh hệ thống (Menu Tệp, Ctrl+Shift+R).

Phiên bản 0.5.5 – 2026-01-03
Tính năng mới
• Thêm Terminal hỗ trợ tiếp cận được tối ưu hóa cho nội dung đầu ra lớn và trình đọc màn hình (Ctrl+Shift+P).
• Thêm cài đặt để lưu tùy chỉnh người dùng vào thư mục hiện tại (chế độ portable).
Sửa lỗi
• Cải thiện đoạn trích dẫn trong Tìm trong các tệp để phần xem trước luôn khớp với kết quả tìm thấy.

Phiên bản 0.5.4 – 2026-01-03
Cải tiến
• Sửa lỗi Chuẩn hóa khoảng trắng (Ctrl+Shift+Enter).
• Thêm hỗ trợ HTML/HTM (mở dưới dạng văn bản).

Phiên bản 0.5.3 – 2026-01-02
Tính năng mới
• Thêm tính năng Tìm trong các tệp.
• Thêm các công cụ văn bản mới: Chuẩn hóa khoảng trắng, Ngắt dòng cứng và Loại bỏ thẻ Markdown.
• Thêm Thống kê văn bản (Alt+Y).
• Thêm các lệnh danh sách mới trong menu Chỉnh sửa:
• Sắp xếp các mục (Alt+Shift+O)
• Giữ lại các mục duy nhất (Alt+Shift+K)
• Đảo ngược các mục (Alt+Shift+Z)
• Thêm Trích dẫn / Bỏ trích dẫn các dòng (Ctrl+Q / Ctrl+Shift+Q).
Bản địa hóa
• Thêm ngôn ngữ tiếng Tây Ban Nha.
• Thêm ngôn ngữ tiếng Bồ Đào Nha.
Cải tiến
• Khi đang mở tệp EPUB, lệnh Lưu sẽ tự động chuyển thành Lưu mới thành và xuất nội dung dưới dạng tệp .txt để tránh làm hỏng tệp EPUB.

## 0.5.2 - 2026-01-01

* Thêm Nhật ký thay đổi.
* Thêm tùy chọn "Mở với Novapad" và liên kết tệp trong quá trình cài đặt.
* Cải thiện ngôn ngữ cho các thông báo (lỗi, hộp thoại, xuất sách nói).
* Thêm lựa chọn phần khi dùng "Chia nhỏ sách nói dựa trên văn bản", với tùy chọn "Bắt buộc dấu đánh dấu ở đầu dòng".
* Thêm tính năng nhập bản phụ đề YouTube với lựa chọn ngôn ngữ, mốc thời gian và cải thiện xử lý tiêu điểm.

## 0.5.1 - 2025-12-31

* Cập nhật tự động có xác nhận, cải thiện thông báo và xử lý lỗi.
* Cải tiến xuất sách nói (chia nhỏ theo văn bản, SAPI5/Media Foundation, điều khiển nâng cao).
* Cải tiến TTS (tạm dừng/tiếp tục, từ điển thay thế, danh sách yêu thích).
* Thêm menu Hiển thị và các bảng giọng đọc/yêu thích, chỉnh màu và cỡ chữ.
* Tự động chọn ngôn ngữ theo hệ thống và cải thiện bản địa hóa.
* Thiết lập quy trình đóng gói cho Windows (MSI/NSIS).

## 0.5.0 - 2025-12-27

* Tái cấu trúc theo mô-đun (trình soạn thảo, xử lý tệp, menu, tìm kiếm).
* Cập nhật quy trình đóng gói trên Windows và tệp README/giấy phép.
* Sửa lỗi điều hướng phím TAB trong cửa sổ Trợ giúp.

## 0.5 - 2025-12-27

* Nâng cấp phiên bản sơ bộ.

## 0.1.0 - 2025-12-25

* Phiên bản phát hành đầu tiên: Cấu trúc dự án và tệp README.
