# Nhật ký thay đổi

Phiên bản 0.6.0 – 2025-01-20
Tính năng mới
• Thêm trình kiểm tra chính tả. Từ menu chuột phải, người dùng có thể kiểm tra xem từ hiện tại có đúng không và nếu không, sẽ nhận được các gợi ý sửa lỗi.
• Thêm tính năng nhập và xuất podcast thông qua tệp OPML.
• Thêm hỗ trợ tìm kiếm trên Podcast Index cùng với iTunes. Người dùng có thể nhập API key và secret miễn phí (tạo chỉ bằng địa chỉ email).
• Thêm hỗ trợ cho các giọng đọc SAPI4, áp dụng cho cả việc đọc thời gian thực và tạo sách nói.
• Thêm tính năng tự động chuyển sang OCR cho các tệp PDF không hỗ trợ tiếp cận: khi không tìm thấy văn bản có thể trích xuất, tài liệu sẽ được nhận dạng qua OCR.
• Thêm hỗ trợ từ điển bằng Wiktionary. Nhấn phím Applications sẽ hiển thị các định nghĩa, và khi có sẵn, sẽ hiện cả từ đồng nghĩa cùng bản dịch sang các ngôn ngữ khác.
• Thêm tính năng nhập bài viết Wikipedia với khả năng tìm kiếm, chọn kết quả và nhập trực tiếp vào trình soạn thảo.
• Thêm phím tắt Shift+Enter trong mô-đun RSS để mở trực tiếp bài viết trên trang web gốc.
Cải tiến
• Việc lựa chọn Micro giờ đây luôn được ứng dụng tuân thủ chính xác.
• Trong cửa sổ podcast, nhấn Enter vào một tập tin giờ đây sẽ thông báo ngay lập tức "đang tải" qua NVDA để xác nhận thao tác.
• Trong kết quả tìm kiếm podcast, nhấn Enter giờ đây sẽ đăng ký theo dõi podcast đã chọn.
• Sửa và cải thiện các nhãn cho phím tắt Ctrl+Shift+O và Podcast Ctrl+Shift+P.
• Tốc độ phát và âm lượng giờ đây được lưu trong cài đặt và duy trì cho tất cả các tệp âm thanh.
• Thêm một thư mục bộ nhớ đệm (cache) riêng cho các tập podcast. Người dùng có thể giữ lại các tập phim qua mục "Giữ podcast" trong menu Phát lại. Bộ nhớ đệm sẽ tự động được dọn dẹp khi vượt quá kích thước do người dùng thiết lập (Tùy chọn → Âm thanh).
• Cải thiện đáng kể việc tải bài viết RSS bằng cách sử dụng giả lập libcurl với cấu hình Chrome và iPhone, đảm bảo tương thích với khoảng 99% các trang web.
• Thêm trạng thái đã đọc / chưa đọc cho các bài viết RSS, với chỉ báo rõ ràng trong danh sách RSS.
• Tính năng Thay thế tất cả giờ đây sẽ báo cáo số lượng thay thế đã thực hiện.
• Thêm nút Xóa Podcast khi điều hướng thư viện podcast bằng phím Tab.
Sửa lỗi
• Loại bỏ mục "bản cập nhật đang chờ" thừa trong menu Trợ giúp (việc cập nhật đã được xử lý tự động).
• Sửa lỗi nhấn Ctrl+S trên tệp MP đang mở gây lưu đè và làm hỏng tệp.
• Sửa lỗi giao diện khiến "Sách nói hàng loạt" hiển thị thành "(B)… Ctrl+Shift+B" (loại bỏ nhãn thừa).
• Sửa lỗi ngoặc kép thông minh: khi được bật, các dấu ngoặc kép thông thường giờ đây sẽ được thay thế chính xác bằng ngoặc kép thông minh.
• Sửa lỗi sử dụng "Đi tới dấu trang" làm đặt lại tốc độ phát về 1.0.
• Sửa lỗi các tập podcast đã tải về lại bị tải lại thay vì sử dụng bản lưu trong bộ nhớ đệm.
Phím tắt
• F1 giờ đây mở hướng dẫn Trợ giúp.
• F2 giờ đây kiểm tra các bản cập nhật.
• F7 / F8 giờ đây nhảy đến lỗi chính tả trước đó hoặc tiếp theo.
• F9 / F10 giờ đây chuyển đổi nhanh giữa các giọng đọc yêu thích.
Cải tiến cho nhà phát triển
• Các lỗi không còn bị bỏ qua một cách im lặng: tất cả các mẫu `let _ =` đã bị loại bỏ, và các lỗi hiện được xử lý rõ ràng (truyền đi, ghi nhật ký hoặc xử lý bằng các phương án dự phòng phù hợp).
• Dự án giờ đây sẽ không thể biên dịch nếu có cảnh báo (warnings): cả `cargo check` và `cargo clippy` phải vượt qua một cách sạch sẽ.
• Các triển khai tùy chỉnh như strlen / wcslen đã bị loại bỏ. Độ dài chuỗi và bộ đệm UTF-16 giờ đây được lấy trực tiếp từ dữ liệu do Rust quản lý thay vì quét bộ nhớ.
• Việc xử lý DLL đã được làm sạch và hợp nhất xung quanh `libloading`, tránh các logic trình nạp tùy chỉnh và phân tích cú pháp PE.
• Các trình hỗ trợ phân tích cú pháp byte tự viết đã bị loại bỏ; tất cả việc phân tích byte hiện sử dụng `from_le_bytes` / `from_be_bytes` tiêu chuẩn trên các lát cắt (slices) đã được kiểm tra.
Những thay đổi này giúp giảm việc sử dụng mã không an toàn (unsafe) không cần thiết, loại bỏ các hành vi không xác định tiềm ẩn và làm cho mã nguồn trở nên chuẩn mực, mạnh mẽ và dễ bảo trì hơn.

Phiên bản 0.5.9 - 2025-01-13
Tính năng mới
• Thêm tính năng sắp xếp lại RSS từ menu chuột phải (lên/xuống/đến vị trí) với kiểm tra vị trí hợp lệ.
• Thêm menu ngữ cảnh cho bài viết với các tùy chọn mở trang web gốc và chia sẻ qua WhatsApp, Facebook và X.
• Thêm phím tắt Esc để quay lại danh sách RSS từ các bài viết đã nhập.
• Thêm chế độ podcast: tìm kiếm, đăng ký, lắng nghe; sắp xếp lại các đăng ký; phím Esc dừng phát và quay lại danh sách; phím Enter trên một tập phim để bắt đầu phát.
• Thêm điều khiển tốc độ phát cho podcast và tệp MP3.
• Thêm Ctrl+T để đi tới một mốc thời gian cụ thể.
• Thêm nút nghe thử giọng đọc sau hộp chọn âm lượng.
• Thêm tính năng tìm kiếm và thay thế bằng biểu thức chính quy (Regex) theo phong cách Notepad++.
• Thêm tính năng nhập RSS từ tệp OPML và TXT.
• Thêm tùy chọn để bật "Mở bằng Novapad" trong File Explorer, bao gồm cả các bản portable.
Cải tiến
• Cải thiện việc chọn tốc độ/cao độ/âm lượng giọng đọc, tuân thủ các giới hạn tối đa của TTS.
• Nhiều cải tiến RSS để tải xuống tất cả bài viết mà không làm di chuyển tiêu điểm NVDA trong quá trình cập nhật.
• Cải thiện việc phát âm thanh với menu chuyên dụng, thông báo thời gian bằng Ctrl+I và âm lượng lên tới 300%.
• Thêm các phím tắt còn thiếu cho một số chức năng.
• Tổ chức lại menu Chỉnh sửa với menu con dọn dẹp văn bản.
• Tổ chức lại Tùy chọn thành các tab, với điều hướng bằng Ctrl+Tab và Ctrl+Shift+Tab.
• Trình đọc RSS hiện tải toàn bộ nội dung bài viết, khớp với chế độ xem trên trình duyệt.
Sửa lỗi
• Sửa lỗi dọn dẹp Markdown làm xóa mất các số ở đầu dòng.
• Sửa lỗi AltGr+Z kích hoạt lệnh hoàn tác.
• Sửa lỗi hủy ghi âm sách nói để quá trình dừng lại nhanh chóng.
Bản địa hóa
• Thêm bản dịch tiếng Việt (cảm ơn Anh Đức Nguyễn).

Phiên bản 0.5.8 - 2026-01-10
Tính năng mới
• Thêm điều khiển âm lượng cho micro và âm thanh hệ thống khi ghi âm podcast.
• Thêm tính năng mới để nhập bài viết từ các trang web hoặc nguồn cấp dữ liệu RSS, bao gồm các nguồn quan trọng nhất cho mỗi ngôn ngữ.
• Thêm chức năng xóa tất cả dấu trang cho tệp hiện tại.
• Thêm chức năng xóa các dòng trùng lặp và các dòng trùng lặp liên tiếp.
• Thêm chức năng đóng tất cả các tab hoặc cửa sổ ngoại trừ cái hiện tại.
• Thêm mục Quyên góp trong menu Trợ giúp cho tất cả các ngôn ngữ.
Cải tiến
• Cải thiện terminal hỗ trợ tiếp cận để ngăn chặn một số lỗi treo máy.
• Cải thiện và sửa lỗi các phím truy cập và phím tắt trong toàn bộ ứng dụng.
• Sửa lỗi đóng cửa sổ phát âm thanh nhưng âm thanh không dừng.
• Thêm hộp thoại xác nhận cho các hành động quan trọng (ví dụ: xóa dòng trùng lặp, xóa dấu gạch nối cuối dòng, xóa tất cả dấu trang). Không có hộp thoại nào hiển thị khi hành động đó không thể thực hiện.
• Thêm khả năng xóa các nguồn RSS/trang web khỏi thư viện bằng cách chọn chúng và nhấn phím Delete.
• Thêm menu chuột phải trong cửa sổ RSS để chỉnh sửa hoặc xóa các nguồn RSS/trang web.
• Loại bỏ cài đặt di chuyển cài đặt sang thư mục hiện tại; ứng dụng hiện tự động xử lý việc này dựa trên vị trí (nếu thư mục chứa file exe tên là "novapad portable" hoặc nằm trên ổ đĩa di động, cài đặt sẽ vào thư mục `config` của exe, nếu không sẽ vào `%APPDATA%\Novapad`).

Phiên bản 0.5.7 - 2026-01-05
Tính năng mới
• Thêm tính năng Sách nói hàng loạt để chuyển đổi nhiều tệp/thư mục cùng lúc.
• Thêm hỗ trợ cho các tệp Markdown (.md).
• Thêm lựa chọn bảng mã khi mở các tệp văn bản.
• Thêm tùy chọn trong terminal hỗ trợ tiếp cận để thông báo khi có dòng mới bằng NVDA.
Cải tiến
• Ghi âm sách nói giờ đây lưu trực tiếp sang MP3 khi được chọn.
• Người dùng giờ đây có thể chọn vị trí dấu sao (*) báo hiệu chưa lưu trên tiêu đề cửa sổ.
• Cải thiện độ ổn định của hệ thống cập nhật.
• Thêm mục "Xóa dấu gạch nối" trong menu Chỉnh sửa để sửa lỗi ngắt dòng OCR.

Phiên bản 0.5.6 - 2026-01-04
Sửa lỗi
  Cải thiện Tìm trong các tệp để nhấn Enter sẽ mở tệp chính xác tại đoạn văn bản đã chọn.
Cải tiến
  Thêm hỗ trợ PPT/PPTX (mở dưới dạng văn bản).
  Mở các định dạng không phải văn bản giờ đây sẽ lưu thành .txt để tránh lỗi định dạng (PDF/DOC/DOCX/EPUB/HTML/PPT/PPTX).
  Thêm ghi âm podcast từ micro và âm thanh hệ thống (Menu Tệp, Ctrl+Shift+R).

Phiên bản 0.5.5 – 2026-01-03
Tính năng mới
• Thêm terminal hỗ trợ tiếp cận được tối ưu hóa cho đầu ra lớn và trình đọc màn hình (Ctrl+Shift+P).
• Thêm cài đặt để lưu cài đặt người dùng trong thư mục hiện tại (chế độ portable).
Sửa lỗi
• Cải thiện đoạn trích dẫn Tìm trong các tệp để phần xem trước luôn khớp với kết quả tìm thấy.

Phiên bản 0.5.4 – 2026-01-03
Cải tiến
• Sửa lỗi Chuẩn hóa khoảng trắng (Ctrl+Shift+Enter).
• Thêm hỗ trợ HTML/HTM (mở dưới dạng văn bản).

Phiên bản 0.5.3 – 2026-01-02
Tính năng mới
• Thêm tính năng Tìm trong các tệp.
• Thêm các công cụ văn bản mới: Chuẩn hóa khoảng trắng, Ngắt dòng cứng và Loại bỏ Markdown.
• Thêm Thống kê văn bản (Alt+Y).
• Thêm các lệnh danh sách mới trong menu Chỉnh sửa:
• Sắp xếp các mục (Alt+Shift+O)
• Giữ lại các mục duy nhất (Alt+Shift+K)
• Đảo ngược các mục (Alt+Shift+Z)
• Thêm Trích dẫn / Bỏ trích dẫn các dòng (Ctrl+Q / Ctrl+Shift+Q).
Bản địa hóa
• Thêm bản dịch tiếng Tây Ban Nha.
• Thêm bản dịch tiếng Bồ Đào Nha.
Cải tiến
• Khi mở tệp EPUB, lệnh Lưu giờ đây tự động chuyển thành Lưu mới thành và xuất nội dung dưới dạng tệp .txt để tránh làm hỏng EPUB.

## 0.5.2 - 2026-01-01
- Thêm nhật ký thay đổi.
- Thêm các tùy chọn mở bằng Novapad và liên kết tệp trong khi cài đặt.
- Cải thiện bản địa hóa thông báo (lỗi, hộp thoại, xuất sách nói).
- Thêm lựa chọn phần khi dùng "Chia nhỏ sách nói dựa trên văn bản", với tùy chọn "Bắt buộc dấu đánh dấu ở đầu dòng".
- Thêm tính năng nhập phụ đề YouTube với lựa chọn ngôn ngữ, mốc thời gian và cải thiện xử lý tiêu điểm.

## 0.5.1 - 2025-12-31
- Cập nhật tự động có xác nhận, cải thiện thông báo và xử lý lỗi.
- Cải tiến xuất sách nói (chia nhỏ theo văn bản, SAPI5/Media Foundation, điều khiển nâng cao).
- Cải tiến TTS (tạm dừng/tiếp tục, từ điển thay thế, danh sách yêu thích).
- Menu Hiển thị và các bảng giọng đọc/yêu thích, màu chữ và cỡ chữ.
- Ngôn ngữ mặc định theo hệ thống và cải thiện bản địa hóa.
- Đóng gói cho Windows (MSI/NSIS).

## 0.5.0 - 2025-12-27
- Tái cấu trúc theo mô-đun (trình soạn thảo, xử lý tệp, menu, tìm kiếm).
- Quy trình đóng gói trên Windows và cập nhật README/giấy phép.
- Sửa lỗi điều hướng phím TAB trong cửa sổ Trợ giúp.

## 0.5 - 2025-12-27
- Nâng cấp phiên bản sơ bộ.

## 0.1.0 - 2025-12-25
- Phiên bản phát hành đầu tiên: Cấu trúc dự án và tệp README.