# Nhật ký thay đổi

Phiên bản 0.6.8 – 2026-03-24

Có gì mới
• Đã thêm một mục mới trong menu Phát để chép lời bất kỳ tệp âm thanh hoặc video nào bằng Whisper. Trong Tùy chọn có một phần mới tên là “AI và Chuyển lời”, nơi bạn có thể chọn mô hình, bật hỗ trợ CUDA tùy chọn cho card đồ họa NVIDIA, giữ nguyên ngôn ngữ gốc và bật hoặc tắt dấu thời gian.
• Đã thêm khả năng dùng đọc chính tả bằng giọng nói ngoại tuyến, với cách hoạt động giống như chép lời âm thanh. Mặc định, nhấn `Ctrl+Shift+Space` để bắt đầu đọc chính tả và nhấn lại đúng phím tắt đó để kết thúc; có thể tùy chỉnh phím tắt trong phần Tùy chọn. Từ lần kích hoạt thứ hai trở đi, việc đọc chính tả sẽ nhanh hơn vì bộ máy vẫn sẵn sàng trong bộ nhớ; trên các PC có dưới 4 GB RAM, việc nạp sẵn và tái sử dụng này sẽ tự động bị tắt.
• Tìm kiếm podcast giờ mặc định dùng `iTunes + Spreaker`, với bộ lọc loại bỏ kết quả trùng lặp khi cùng một podcast xuất hiện trên cả hai nền tảng.
• Đã cải thiện tìm kiếm và duyệt podcast Apple: tìm kiếm podcast, duyệt theo danh mục và top podcast theo danh mục giờ dùng quốc gia thư mục podcast đã chọn. Trong Tùy chọn > RSS / Podcast, có thể để `Tự động` để dùng quốc gia hệ thống hoặc tự chọn một quốc gia khác.
• Sonarpad hiện cũng đã có bản cho Mac, dù hiện chỉ hỗ trợ một phần chức năng. Liên kết dự án: https://github.com/Ambro86/Sonarpad-Mac

Cải tiến
• Đã thêm hơn 50 quốc gia có thể chọn cho thư mục podcast, giúp người dùng chọn được nhiều danh mục quốc gia hơn.
• "Phat am thanh tu streaming..." gio cung cho phep tim kiem tren YouTube bang bat ky chuoi van ban nao, hoac dan lien ket cua mot kenh hoac playlist YouTube de hien thi cac ket qua cua no.
• Đã cải thiện cách hiển thị kết quả trong "Phat am thanh tu streaming...": các mục YouTube giờ bao gồm tiêu đề, thời lượng, kênh và lượt xem theo định dạng rõ ràng hơn.
• Đã thêm mục yêu thích YouTube cho kênh và danh sách phát trong "Phat am thanh tu streaming...": có thể thêm từ kết quả bằng menu ngữ cảnh, mở trực tiếp từ danh sách Yêu thích truy cập bằng Tab ngay sau trường URL/truy vấn YouTube và xóa sau đó cũng từ chính danh sách đó bằng menu ngữ cảnh. Trong kết quả tìm kiếm YouTube, menu ngữ cảnh chỉ khả dụng cho kênh và danh sách phát.
• Đã cải thiện focus khi dùng "Phat am thanh tu streaming...", để cửa sổ tiến trình ổn định hơn trong lúc tải xuống và chuyển đổi.
• Đã thêm hai thao tác đọc mới trong menu Giọng nói: `Câu trước` và `Câu tiếp theo`, với phím tắt có thể cấu hình để nhảy trong khi đọc văn bản.
• Phím tắt mặc định của `Chạy tệp bằng trình thông dịch` giờ là `Ctrl+Shift+F5`, để `Shift+F5` có thể được dùng mặc định cho `Câu trước`.
• Đã thêm quản lý hồ sơ giọng nói trong Tùy chọn > Giọng nói: có thể thêm, đổi tên và xóa hồ sơ.
• Đã mở rộng trong Tùy chọn > Âm thanh các lựa chọn cho khoảng tua lùi khi phát, với các giá trị mới từ 1 giây đến 2 giờ.
• Đã thêm bản dịch tiếng Nga nhờ Dmitriy.
• Đã thêm trong Tùy chọn > Âm thanh một lựa chọn mới cho định dạng tên phần của sách nói: `Tiêu đề + số`, `Chỉ số` hoặc `Số + tiêu đề`.
• Đã thêm hành động trong menu ngữ cảnh bài RSS để thêm bài viết vào mục yêu thích.
• Nguồn RSS "Yêu thích" có thể bị xóa và sẽ được tạo lại tự động khi thêm bài viết mới vào mục yêu thích.
• Đã thêm phím tắt RSS để di chuyển nguồn lên/xuống: `Ctrl+Shift+Mũi tên lên` và `Ctrl+Shift+Mũi tên xuống`.
• Đã cải thiện cửa sổ RSS với phần xem trước bài viết tích hợp, giúp có thể xem ngay nội dung tại đó và chuyển nhanh tới bằng phím Tab trước khi mở toàn bộ bài viết trong trình soạn thảo.
• Đã thêm trong RSS một mục rõ ràng “Tải thêm tin tức” ở cuối nguồn khi còn bài khác; nhấn Enter sẽ tải khối tiếp theo và đưa focus tới bài viết mới đầu tiên.
• Trong từ điển giọng nói, khi thêm hoặc sửa một mục thay thế, giờ có thêm ô «Phân biệt chữ hoa/thường» để quyết định mỗi phép thay thế có phân biệt hay bỏ qua chữ hoa chữ thường.
Sửa lỗi
• Đã sửa chức năng nhập từ Wikipedia: trên một số trang, các đoạn trích dẫn trong bài không được nhập đúng.
• Đã cải thiện bộ phân tích trang web: trên một số trang WordPress, các mục danh sách và một số tiêu đề phần không được đưa vào.
• Khi dùng “Đi đến dòng”, ô nhập giờ sẽ được điền sẵn bằng dòng hiện tại.
• Đã sửa xuất OPML cho podcast và RSS, vì vậy các tệp xuất ra giờ đã được iTunes chấp nhận.
• Đã thêm thông báo xác nhận được bản địa hóa cho việc nhập và xuất OPML podcast và nguồn RSS đúng cách.
• Đã sửa lỗi trong "Phat am thanh tu streaming..." khiến chương trình có thể trông như bị treo khi nhập chuỗi tìm kiếm và chọn một kênh YouTube từ kết quả, thay vì mở danh sách video của kênh đó.
• Đã sửa lỗi khiến danh sách tệp đang mở hiển thị trong menu Trợ giúp thay vì menu Cửa sổ.
• Đã sửa một trường hợp biên khi phát trực tuyến: phát âm thanh có thể bắt đầu nhưng cửa sổ “Đang tải streaming” vẫn mở khi tệp đã tải xuống đã đúng định dạng đích.
• Đã sửa hành vi chuyển đổi với streaming MP3: khi luồng đã là MP3 và người dùng chọn bitrate MP3 cụ thể (ví dụ 128 kbps), Sonarpad giờ sẽ mã hóa lại theo bitrate đã chọn thay vì bỏ qua bước chuyển đổi.
• Đã sửa phím tắt `Alt+Shift+L`: giờ đây phím này mở đúng danh sách chương trong khi phát.
• Đã sửa phím tắt `Alt+Shift+T`: giờ đây phím này khởi động đúng chức năng “Chép lời âm thanh hiện tại” thay vì mở menu Công cụ.
• Nếu đang có âm thanh được phát, khi bắt đầu chép lời Sonarpad giờ sẽ tự động tạm dừng âm thanh đó trước khi bắt đầu.
• Đã sửa lỗi khiến khi nhập một bài viết từ Wikipedia, quá trình nhập có thể thành công nhưng phần văn bản bài viết lại không hiển thị trên màn hình.
• Đã bổ sung hỗ trợ chương podcast nhúng trong tệp media cục bộ (ví dụ metadata chương của MP3): khi feed/URL không cung cấp chương, Sonarpad sẽ nạp chương từ tệp đã tải ở chế độ nền, giúp phát bắt đầu ngay và áp dụng chương ngay khi sẵn sàng.
• Đã sửa việc nạp chương cho các tập podcast đã tải xuống rồi mở như tệp media cục bộ thông thường: các chương nhúng giờ cũng khả dụng trong trường hợp này, không chỉ khi bắt đầu phát từ cửa sổ Podcast.
• Đã sửa bước hoàn tất cuối cùng của sách nói MP3 với SAPI4 và SAPI5: tệp đầu ra cuối cùng giờ được hoàn tất đúng cách, tránh các tệp không đầy đủ hoặc kém ổn định sau các lần xuất dài.
• Đã thêm thanh tiến trình rõ ràng cho giai đoạn hoàn tất trong mọi chế độ tạo sách nói: sau giai đoạn tạo, Sonarpad giờ thông báo và hiển thị riêng giai đoạn hoàn tất với tiến trình nhìn thấy được.
• Đã sửa lỗi giọng hội thoại: các tham số tốc độ/cao độ/âm lượng của giọng hội thoại thứ nhất và thứ hai giờ được áp dụng đúng trong quá trình tổng hợp giọng nói.
• Đã cải thiện nhận diện mã hóa cho tệp `.txt` tiếng Nhật: thêm fallback Shift_JIS/CP932 an toàn cho các trường hợp mojibake, đồng thời giữ nguyên hành vi hiện có với UTF/diacritics/tiếng Trung.
• Tái cấu trúc an toàn nội bộ: chuyển sang triển khai safe ở những nơi có thể và giảm mạnh số dòng mã unsafe.
Phiên bản 0.6.7 – 2026-03-02
Cải tiến
• Ora il programma riesce a gestire Sostituisci tutto in modo massivo su file grandi con un gran numero di sostituzioni.
• Đã cập nhật bản dịch tiếng Ba Lan nhờ DJ Graco.
• Đã thêm bản dịch tiếng Litva.
• Đã thêm bản dịch tiếng Trung.
• Từ bây giờ, các bản dựng beta thường xuyên sẽ được phát hành trong mục Releases của dự án, để người dùng có thể thử nghiệm các thay đổi mới trước bản phát hành ổn định tiếp theo.
• Đã thêm phím tắt `Ctrl+.` để chèn ký tự dấu ba chấm (…).
• Đã cải thiện hỗ trợ chương podcast: điều hướng chương hiện ổn định hơn, kể cả với các tập phát trực tiếp/streaming không có chương nhúng trong tệp MP3, bằng cách dùng metadata chương từ feed/URL làm fallback khi có sẵn. Đã thêm phím tắt `Ctrl+Alt+PageUp` (chương trước) và `Ctrl+Alt+PageDown` (chương tiếp theo).
• Đã tổ chức lại thư mục đầu ra về `Documents\\Sonarpad`: tệp giờ được lưu vào các thư mục con riêng `audiobooks`, `documents`, `recordings` và `media`, kèm tự động chuyển dữ liệu từ đường dẫn cũ.
• Đã cải thiện hỗ trợ cho tệp văn bản rất lớn (kể cả 60 MB): mở tệp và điều hướng theo từng dòng mượt hơn, đặc biệt khi dùng trình đọc màn hình.
• Đã cập nhật hướng dẫn cho tất cả ngôn ngữ và làm mới tài nguyên bản địa hóa trên toàn bộ ứng dụng, bao gồm nội dung quyên góp và bản dịch trình cài đặt NSIS (thêm chuỗi cài đặt mới cho tiếng Trung giản thể và tiếng Litva, đồng thời hoàn thiện bản dịch tiếng Ukraina của setup).
• Da bo sung ho tro proxy mang toan cuc (HTTP/HTTPS va SOCKS5/SOCKS5H) cho cac tinh nang truc tuyen, kem kiem tra khi luu Tuy chon: proxy khong hop le se duoc canh bao va xoa tu dong.
• Da them tinh nang moi trong Cong cu: "Phat am thanh tu streaming...", cho phep dan URL (YouTube hoac lien ket media truc tiep), chon dinh dang dau ra va ho so chat luong/bitrate (bao gom chat luong/bitrate goc cho MP3 va MP4) va phat ngay trong trinh phat cua Sonarpad.
• Da them ho tro phim da phuong tien Play/Pause cua he thong (tai nghe/ban phim): nay co the dieu khien ca phat media va tam dung/tiep tuc doc van ban (uu tien trinh phat media khi ca hai dang hoat dong).
• Da them muc moi trong Tep > Tep gan day: "Xoa tep gan day" de xoa nhanh danh sach tai lieu vua mo.
• Đã mở rộng tùy chọn bitrate trong Chuyển đổi âm thanh và cài đặt ghi podcast: thêm mức thấp hơn (64/96 kbps) và nâng MP3 lên tối đa 320 kbps, đồng bộ cả kiểm tra hợp lệ và xử lý bộ mã hóa.
• Đã mở rộng tùy chọn chia sách nói theo thời lượng lên đến 60 phút.
• Đã cải thiện chia sách nói theo số phần: người dùng có thể nhập thủ công số phần, với kiểm tra hợp lệ từ 1 đến 100.
• Đã thêm chế độ mới Xem > Chỉ đọc để ngăn chỉnh sửa nhầm văn bản, đồng thời vẫn giữ khả năng đọc và điều hướng đầy đủ tài liệu.
• Đã thêm thanh tiến trình có thể truy cập trong quá trình cập nhật chương trình, giúp trình đọc màn hình theo dõi tiến độ tải xuống theo thời gian thực.
• Đã thêm thanh trạng thái mới dạng yên lặng ở cửa sổ chính với số ký tự, số từ và vị trí dòng/cột (ví dụ: "Ký tự (kể cả khoảng trắng): 11. | Từ: 2. | Ln 1, Col 12"), không làm ảnh hưởng tiêu điểm của NVDA.
• Đã thêm tùy chọn mới trong menu Xem cho tự động xuống dòng, giúp bật/tắt nhanh chế độ xuống dòng mà không cần mở Tùy chọn.
• Đã thêm trong Chỉnh sửa > Văn bản các thao tác tăng/giảm thụt lề, với phím tắt Ctrl+Shift+. (thụt vào) và Ctrl+Shift+, (thụt ra), vì khi bật “Hiển thị giọng nói trong trình soạn thảo” thì phím Tab được dành cho việc điều hướng bảng giọng nói.
• Đã thêm hiển thị ngày/giờ theo ngôn ngữ cho bài RSS và tập podcast, với định dạng tự điều chỉnh theo ngôn ngữ giao diện.
• Đã thêm thao tác mới trong menu ngữ cảnh RSS để chia sẻ bài đã chọn qua email.
• Đã thêm tùy chọn xác nhận xóa chi tiết trong Tùy chọn > RSS và podcast: với RSS (nguồn/bài/cả hai/không) và với podcast (podcast/tập/cả hai/không).
• Đã thêm sao chép nhanh RSS có thể cấu hình bằng Ctrl+C (Tùy chọn > RSS và podcast): sao chép tiêu đề, URL, nội dung bài viết hoặc tất cả.
• Đã hợp nhất luồng RSS: “Thêm nguồn” giờ chấp nhận cả URL feed và từ khóa (tự động tạo feed Google News), không cần chức năng tìm kiếm riêng.
• Khi nhấn Ctrl+A, chương trình giờ sẽ thông báo hoàn tất thao tác để phản hồi rõ ràng hơn cho trình đọc màn hình.
• Đã thêm Shift+F3 cho "Tìm trước đó" trong menu Chỉnh sửa, bổ sung cho F3 "Tìm tiếp theo".
• Đã cải thiện thông báo thay thế với dạng số ít/số nhiều chính xác (ví dụ: “1 mục đã thay thế” và “2 mục đã thay thế”).
• Đã thêm trong cửa sổ từ điển tùy chọn chọn ngôn ngữ tra cứu, mặc định là Auto (ngôn ngữ giao diện) và có thể chọn thủ công.
• Đã thêm tab Phím tắt mới trong Tùy chọn để tùy chỉnh tổ hợp phím, kèm phát hiện xung đột và cảnh báo khi một phím tắt đã được gán cho hành động khác.
• Đã thêm hỗ trợ ban đầu cho tham số dòng lệnh: `-h`/`--help` hiển thị trợ giúp nhanh và `--version` hiển thị phiên bản chương trình.
• Đã làm rõ hơn cách chỉnh tay tốc độ và cao độ: các ô chỉnh tay giờ dùng thang lấy 100 làm mốc, trong đó 100 tương ứng giá trị bình thường.
• Đã cải thiện chọn giọng Microsoft trong cả Tùy chọn > Giọng nói và bảng giọng nói của trình soạn thảo: thêm combobox ngôn ngữ đã bản địa hóa để lọc giọng theo ngôn ngữ, đồng thời vẫn giữ chế độ “chỉ giọng đa ngôn ngữ” là một danh sách duy nhất không chia theo ngôn ngữ (ẩn combobox ngôn ngữ khi chế độ này bật).
• Đã thêm cấu hình giọng cho hội thoại trong Tùy chọn > Giọng nói với điều hướng đầy đủ bằng Tab, dùng cùng mô hình giọng của giao diện chính (engine, bộ lọc ngôn ngữ Edge, giọng và tốc độ/cao độ/âm lượng có nhãn); thêm giọng hội thoại thứ hai tùy chọn với cùng nhóm điều khiển (engine, bộ lọc ngôn ngữ Edge, giọng, tốc độ/cao độ/âm lượng) để luân phiên hội thoại; quy tắc hội thoại được lưu trong cấu hình `.ini`, không sửa đổi văn bản tài liệu.
• Đã cải thiện nhãn Hoàn tác: mục Chỉnh sửa > Hoàn tác giờ hiển thị hành động sẽ được hoàn tác (ví dụ chỉnh sửa văn bản, thêm/bỏ comment dòng hoặc chèn thẻ giọng nói), đồng thời vẫn bị vô hiệu khi không có gì để hoàn tác.
Sửa lỗi
• Đã sửa hỗ trợ mở tệp RTF: tệp `.rtf` giờ được trích xuất và hiển thị thành văn bản dễ đọc, không còn hiển thị mã RTF thô (ví dụ `{\\rtf1...}`).
• Đã sửa mở tệp văn bản tiếng Trung mã hóa GB18030/GBK: Sonarpad giờ phát hiện và giải mã đúng, tránh hiện tượng chữ lỗi (mojibake).
• Đã cải thiện tạo sách nói M4B với metadata và marker chương; đã sửa lỗi "chipmunk" (giọng quá cao/quá nhanh) trong tệp M4B tạo ra.
• Đã sửa giao diện bitrate trong cửa sổ lưu sách nói: đã bỏ các nhãn hardcoded bằng tiếng Ý và thêm tùy chọn 64 kbps vào danh sách bitrate có thể chọn.
• Đã sửa "Lưu tất cả" (Ctrl+Shift+S): giờ đây tất cả tài liệu đang mở đã chỉnh sửa đều được phát hiện ổn định (kể cả tab mới/chưa lưu), và Lưu tất cả sẽ lưu đúng từng tài liệu, mở "Lưu thành" khi cần.
• Đã sửa thứ tự bài RSS từ Google News: khi có ngày xuất bản, bài viết giờ được hiển thị từ mới nhất đến cũ nhất.
• Đã sửa gán nhãn cho NVDA trong cửa sổ từ điển: ô tìm kiếm và combobox ngôn ngữ giờ đọc đúng nhãn.
• Đã sửa điều hướng bàn phím trong cửa sổ Thuộc tính RSS/Podcast: Tab/Shift+Tab giờ đi tới nút OK, Enter kích hoạt OK, Esc đóng an toàn và tiêu điểm quay lại đúng danh sách RSS/Podcast.
• Đã sửa lịch sử hoàn tác trong RSS/Podcast: Ctrl+Z giờ hỗ trợ hoàn tác nhiều cấp cho thao tác xóa (bài/tập và nguồn), không chỉ thao tác gần nhất.
• Đã cải thiện thông báo khi xóa trong RSS/Podcast với thông điệp rõ ràng (đã xóa RSS, đã xóa bài RSS, đã xóa tập podcast).
• Đã cải thiện xử lý tiêu điểm sau khi xóa/hoàn tác trong RSS/Podcast: với RSS, mục nguồn đầu tiên được chọn lại ổn định khi cần, đồng thời giảm lặp thông báo của trình đọc màn hình trong quá trình chọn lại có độ trễ.

Phiên bản 0.6.6 – 2026-02-13
Cải tiến
• Đã thêm "Định dạng tự động cho TTS" trong menu Chỉnh sửa để chuẩn bị nhanh văn bản cho đọc giọng nói (xóa markdown/dấu ngoặc kép và ghép lại các dòng bị ngắt).
• Đã cải thiện chèn thẻ giọng nói: khi có văn bản được chọn, thẻ giờ được áp dụng đúng cho cả một dòng đơn và vùng chọn nhiều dòng.
• Đã thêm tùy chọn trong cài đặt Âm thanh để chọn thư mục mặc định lưu sách nói (mặc định: Documents\\Sonarpad Audiobooks).
• Trong hộp thoại lưu sách nói, khi bật chế độ chia phần, đã thêm tùy chọn mới (bật mặc định) để tạo thư mục con riêng cho các phần được tạo ra.
• Xuất sách nói giờ lưu MP3 stereo với bitrate do người dùng chọn cho giọng Edge, SAPI5 và SAPI4.
• Đã thêm hỗ trợ giọng SAPI5 32-bit qua bridge, để dùng cả các giọng chỉ có trong engine 32-bit.
• Đã tổ chức lại các tính năng giọng nói vào menu riêng "Giọng nói và âm thanh" và thêm/làm rõ mục "Chuyển đổi âm thanh", dùng để chuyển đổi mọi tệp đa phương tiện được hỗ trợ sang MP3, AAC, OGG, Opus, FLAC, WAV và AIFF.
• Đã thêm khả năng xóa từng bài RSS và từng tập podcast riêng lẻ (phím Delete + menu ngữ cảnh có xác nhận), không cần xóa toàn bộ nguồn RSS/podcast, kèm hoàn tác lần xóa gần nhất (bài/tập đơn lẻ hoặc toàn bộ nguồn RSS/podcast).
• Đã thêm chức năng xuất nguồn RSS sang OPML trong cửa sổ RSS, giúp lưu và nhập lại các nguồn hiện tại một cách dễ dàng.
• Đã thêm tính năng "Tìm RSS theo từ khóa" trong cửa sổ RSS: khi nhập từ khóa, Sonarpad sẽ tự động tạo URL RSS Google News và mở hộp thoại thêm nguồn với thông tin đã điền sẵn, giúp tạo feed theo chủ đề chỉ trong một bước.
• Đã thêm bản dịch tiếng Serbia nhờ Mila Kuran.
• Đã thêm bản dịch tiếng Ukraina nhờ Ivan Shtefuriak.
• Đã thêm mở nhiều tệp media cùng lúc: khi mở nhiều tệp sẽ tạo hàng đợi phát thay vì thay thế tệp hiện tại.
• Đã thêm phím tắt tua biến thiên khi phát: với mức cơ bản 1 phút, Trái/Phải tua 60 giây, Shift+Trái/Phải tua 20 giây, và Ctrl+Trái/Phải tua 3 phút.
• Đã thêm phím tắt chuyển bài trước/sau trong trình phát: Ctrl+PageUp và Ctrl+PageDown.
• Đã thêm mục "Đặt lại âm lượng" và gom các thao tác đặt lại vào submenu riêng "Đặt lại" trong menu Phát lại, cùng với "Đặt lại tốc độ" và "Đặt lại cao độ".
• Cải tiến trình cài đặt: setup.exe giờ cho phép chọn giữa liên kết tất cả kiểu tệp được hỗ trợ hoặc tự chọn từng phần mở rộng; MSI cũng hỗ trợ chọn theo từng phần mở rộng trong cây tính năng (mặc định giữ nguyên: bật tất cả).
• Đã thêm menu mới "Cửa sổ" với mục "Tài liệu đang mở..." để chuyển nhanh đến bất kỳ tệp nào đang mở.
• Đã cập nhật mục Xem > Phông chữ: thay bộ chọn đầy đủ bằng submenu nhanh với các phông phổ biến (Arial, Calibri, Consolas, Segoe UI, Tahoma, Verdana, Times New Roman, Georgia), đồng thời giữ nguyên cỡ chữ hiện tại.
• Đã cải thiện cách đọc RSS và podcast với hai kiểu thông báo tách biệt: nút nguồn sẽ báo "mục mới" khi feed/podcast có nội dung mới, còn từng bài RSS và từng tập podcast sẽ báo "chưa đọc"/"chưa phát"; có thể tắt hành vi này trong Tùy chọn.
Sửa lỗi
• Đã sửa trích xuất văn bản EPUB cho sách có chứa chú thích HTML inline (<!-- ... -->): văn bản chương giờ được phân tích đúng thay vì bị bỏ qua một phần hoặc toàn bộ.
• Đã sửa từ điển Wiktionary tiếng Tây Ban Nha và cơ chế cache từ điển: các từ như "agua" giờ được tìm đúng, và các mục cache cũ kiểu "Không tìm thấy từ" sẽ không còn được tái sử dụng.
• Đã sửa mã hóa khi nhập bài RSS cho một số nguồn tiếng Tây Ban Nha (ví dụ El Mundo): dấu tiếng Tây Ban Nha và ký tự "ñ" giờ được giữ đúng trong trình soạn thảo tạm.
• Đã sửa giải mã ANSI cho các tệp Trung Âu (ví dụ tiếng Séc/tiếng Ba Lan): Sonarpad giờ phân biệt UTF-8 và ANSI tốt hơn và chọn đúng bảng mã (bao gồm Windows-1250), tránh lỗi vỡ dấu.
• Đã sửa lỗi lưu nguồn RSS có tham số trong URL (ví dụ `rss.aspx?c=...`): các feed này giờ được lưu và khôi phục đúng sau khi khởi động lại Sonarpad.
• Đã sửa lỗi mở các tệp con trỏ Google Drive (`.gdoc`, `.gsheet`, `.gslides`) từ menu ngữ cảnh Explorer: khi đọc trực tiếp bị lỗi “Incorrect function (os error 1)”, Sonarpad giờ dùng fallback shell-open để tài liệu vẫn mở đúng.
• Đã sửa đọc tệp Excel legacy `.xls` (Excel 2010): các tệp nhị phân cũ giờ được nhận diện/giải mã đúng thay vì hiển thị văn bản lỗi (ví dụ `ÐÏ_à¡±...`).
• Đã sửa luồng thông báo lỗi chính tả: lỗi sẽ được đọc lại khi bạn rà soát văn bản sau đó, và cùng một lỗi sẽ được báo lại nếu bị xóa rồi gõ lại.
• Đã sửa các thao tác văn bản theo dòng (ví dụ Ctrl+Q / Ctrl+Shift+Q, sắp xếp/đảo/duy nhất/gộp dòng): khi chọn một dòng bằng Shift+Mũi tên xuống, các dòng liền kề không còn bị dính hoặc bị cắt.
• Đã sửa xử lý chọn nhiều dòng cho các thao tác theo dòng (Ctrl+Q / Ctrl+Shift+Q và các công cụ liên quan): khi RichEdit trả về dấu xuống dòng dạng CR-only, Sonarpad giờ chuẩn hóa đúng để xử lý đủ tất cả dòng đã chọn mà không cắt mất ký tự đầu dòng.
• Đã mở rộng chuẩn hóa đầu vào TTS cho các ký hiệu hiển thị của khoảng trắng/tab/xuống dòng (␠/U+2420, ␣/U+2423, ␉/U+2409, ␊/U+240A, ␍/U+240D, ␤/U+2424), vốn có thể gây lặp đoạn với giọng đa ngôn ngữ.
• Đã tinh chỉnh bước làm sạch văn bản cho Edge TTS bằng một pipeline kiểm tra duy nhất: chuẩn hóa khoảng trắng lạ/vô hình, rút gọn các chuỗi dấu câu dài (như "...", "!!!", "???"), và bỏ qua các đoạn chỉ gồm dấu câu để tránh lặp vòng khi phát.
• Đã sửa thông báo thời gian phát (Ctrl+I) cho luồng MP3/podcast: thời gian hiện tại giờ được giới hạn theo tổng thời lượng của track, và phát sẽ tự dừng nếu vị trí vượt quá điểm kết thúc.
• Đã cải thiện phạm vi bản địa hóa của trình cài đặt: setup.exe giờ bao gồm thêm tiếng Séc, Ba Lan, Pháp và Serbia, trong khi MSI được giữ thành một gói en-US duy nhất để tránh gây rối trong bản phát hành.
• Đã sửa dọn dẹp khi gỡ cài đặt các mục menu ngữ cảnh: mục "Mở bằng Sonarpad" giờ được xóa ổn định, kể cả trong các kịch bản registry cũ.
• Đã sửa độ ổn định của tạm dừng/tiếp tục với SAPI5: tạm dừng bằng F4 giờ hoạt động đúng và khi tiếp tục sẽ quay lại đúng vị trí mong đợi thay vì phát lại từ đầu.
• Đã sửa luồng tạm dừng + tua + tiếp tục khi phát media: sau khi tạm dừng và tua bằng Trái/Phải, nhấn Space giờ sẽ tiếp tục ổn định tại vị trí hiện tại thay vì dừng hẳn hoặc phát lại từ đầu.

Phiên bản 0.6.5 – 2026-02-07
Cải tiến
• Bản dịch tiếng Tây Ban Nha được cải thiện nhờ Arturo Fernandez Rivas.
• Đã thêm tùy chọn tách sách nói EPUB theo chương.
• Nhập RSS giờ dùng tab tạm riêng (tiêu đề đã bản địa hóa); Lưu dưới dạng sẽ chuyển thành tài liệu bình thường.
• Thông báo từ trình đọc màn hình giờ cũng được gửi tới JAWS khi có sẵn.
Sửa lỗi
• Đọc từ vị trí con trỏ (F5) giờ bắt đầu chính xác tại con trỏ. Trước đây có thể bắt đầu vài dòng phía trên vì offset con trỏ không khớp với vị trí CRLF/UTF-16.
• Đã sửa lỗi vẽ lại: khi gõ đè lên vùng chọn, phần văn bản phía trước có thể biến mất cho tới khi di chuyển vùng chọn.
• Sửa parser chương EPUB: trang bìa hoặc chỉ có hình ảnh không còn bị đọc CSS (ví dụ "padding") hay tiêu đề "Sconosciuto".
• Đã sửa lỗi chia audiobook từ EPUB theo thời gian: Edge TTS có thể lỗi khi gặp đoạn rỗng hoặc quá dài ("Edge audio not sent").
• Bài viết RSS giờ giải mã các thực thể HTML (ví dụ &quot;, &amp;, &lt;, &gt;).
• Lưu/Lưu dưới tên giờ đây đề xuất tên tệp hiện có khi lưu các định dạng không nên ghi đè (ví dụ: EPUB), thay vì dòng đầu tiên.
• Đã sửa lỗi khiến podcast có tập mới không được thông báo là chưa phát, đồng thời đổi "Chưa nghe" thành "Chưa phát" cho chuyên nghiệp hơn.

Phiên bản 0.6.4 – 2026-02-05
Cải tiến
• Chương trình đã được đổi tên thành Sonarpad để nhấn mạnh âm thanh và audio, là điểm then chốt của chương trình.
• Thêm lựa chọn track âm thanh trong menu Phát lại cho các tệp đa phương tiện có nhiều track âm thanh (ví dụ: MKV với nhiều ngôn ngữ).
• Podcast giờ đây hiển thị rõ ràng những tập chưa nghe với tiền tố "Chưa nghe" trước tên.
• Hệ thống thẻ mới để đổi giọng trong văn bản. Ví dụ:
  - Giọng Microsoft (Edge): <voice edge it-IT-IsabellaNeural>Xin chào</voice>
  - Giọng SAPI5: <voice sapi5 Microsoft Helena Desktop>Xin chào</voice>
  - Giọng SAPI4: <voice sapi4 #1>Xin chào</voice>
• Bổ sung danh mục podcast.
• Thêm tùy chọn trong menu ngữ cảnh để tạo sách nói từ phần chọn.
Sửa lỗi
• Khắc phục lỗi khiến sách nói SAPI4 có thể được tạo ra khác với mong đợi.
• Cửa sổ Tìm trong tệp: nhấn Enter trên một kết quả giờ mở đúng vị trí đoạn trích và Esc quay lại kết quả.
• Cửa sổ Tùy chọn: chỉnh lại bố cục hiển thị ở các tab Chung, Giọng nói, Trình soạn thảo và Âm thanh để tránh thiếu hoặc bị cắt điều khiển.
• Đã sửa lỗi dấu trang khi thay đổi tốc độ phát.
• Đã sửa lỗi Podcast Index và danh mục không hiển thị đúng.
• Đã sửa lỗi dấu nháy đơn làm ngắt đọc: không còn chế độ đọc riêng cho hội thoại, dùng thẻ giọng nói.

Phiên bản 0.6.3 – 2026-01-30
Cải tiến
• Cải thiện phát hiện micrô.
• Thêm phát lại tức thì cho tất cả các định dạng.
Sửa lỗi
• Sửa lỗi sập trong cửa sổ danh mục podcast.

Phiên bản 0.6.2 – 2026-01-30
Tính năng mới
• Thêm hỗ trợ chạy tệp (Shift+F5). Người dùng có thể chọn trình thông dịch (ví dụ: python) trong Tùy chọn, tìm kiếm trên máy tính, và nhấn Shift+F5 để chạy tập lệnh hiện tại. Các tệp HTML mở trong trình duyệt.
• Thêm hỗ trợ cho các tệp con trỏ Google Docs (.gdoc, .gsheet, .gslides), tự động mở trong trình duyệt mặc định.
• Thêm hỗ trợ định dạng sách nói M4B (Apple/AAC).
• Thêm tùy chọn "Hiển thị tập" trong menu ngữ cảnh kết quả tìm kiếm podcast để duyệt và phát các tập mà không cần đăng ký.
• Thêm tính năng "Đi đến dòng" (menu Chỉnh sửa hoặc Ctrl+J) để nhảy nhanh đến số dòng cụ thể.
• Thêm tùy chọn menu ngữ cảnh để sắp xếp nguồn cấp RSS và podcast (theo bảng chữ cái hoặc theo ngày).
• Thêm nguồn cấp RSS mặc định bằng tiếng Việt.
• Thêm hộp kiểm tra micrô trong hộp thoại ghi âm để kiểm tra mức độ trước khi bắt đầu.
• Thêm "Hiển thị mô tả" cho các tập podcast trong menu ngữ cảnh.
• Thêm hỗ trợ cho các định dạng âm thanh/video mở rộng qua FFmpeg: mkv, avi, mov, m4v, webm, mpg, ts, wmv, flv, vob, 3gp, flac, ogg, wma, aiff.
• Thêm đọc phụ đề đồng bộ (srt, vtt, ass, sub, sbv, lrc, smi) với NVDA hoặc giọng đã chọn. Chương trình tìm kiếm tệp phụ đề có cùng tên với tệp phương tiện. Thêm tùy chọn "Nhập phụ đề" và "Xóa phụ đề" trong menu Phát lại cho các tệp có tên khác nhau.
• Thêm liên kết tệp cho tất cả các định dạng âm thanh/video mới được hỗ trợ trong menu ngữ cảnh "Mở bằng Sonarpad".
• Thêm cài đặt để điều chỉnh cao độ của bất kỳ tệp nào.
• Thêm tùy chọn trong Cài đặt Chung để bật hoặc tắt báo cáo lỗi ẩn danh. Thêm mục trong menu Trợ giúp để tạo tệp ZIP chẩn đoán.
• Thêm tùy chọn sử dụng giọng khác cho đối thoại, cả cho đọc trực tiếp và tạo sách nói.
• Thêm trình duyệt danh mục podcast để khám phá podcast theo danh mục (kinh doanh, nghệ thuật, thể thao, v.v.).
Cải tiến
• Mở tệp âm thanh/video từ Explorer giờ mở trực tiếp giao diện trình phát thay vì trình soạn thảo văn bản.
• Loại bỏ yêu cầu OCR cho PDF không thể truy cập; OCR giờ được thực hiện tự động để cải thiện tốc độ và trải nghiệm người dùng.
• Cải thiện Terminal Trợ năng: đọc NVDA giờ nhớ dòng cuối cùng đã đọc để liên tục tốt hơn.
• SAPI 4: Tạo sách nói giờ hoàn toàn song song và gần như tức thì. Thêm yêu cầu để chọn số quy trình đồng thời.
• SAPI 4: Loại bỏ nút cổ chai WAV-MP3 bằng cách chuyển đổi các đoạn song song trong quá trình tổng hợp.
• SAPI 4: Cải thiện xử lý lỗi và dọn dẹp tự động các tệp tạm thời.
• Hộp thoại Tìm: Đổi tên "Regex" thành "Biểu thức chính quy" để rõ ràng hơn và thêm các bản dịch còn thiếu cho các tùy chọn tìm kiếm.
• Sách nói M4B: Xử lý đầu ra tốt hơn; chia theo phần/đánh dấu giờ tạo ra một tệp M4B duy nhất với siêu dữ liệu chương bao gồm tiêu đề và tác giả.
• Trình phát: Sửa độ chính xác của dấu trang và thông báo thời gian khi tốc độ phát không phải 1.0x.
• Khôi phục điều hướng Ctrl+Tab và Ctrl+Shift+Tab trong Tùy chọn.
• Thêm tùy chọn trong menu Phát lại để đặt lại ngay tốc độ về Bình thường (1.0x).
• Cập nhật tất cả các phụ thuộc lên phiên bản mới nhất để có hiệu suất và độ ổn định tốt hơn.
• Tích hợp FFmpeg với tải DLL động để đảm bảo khả năng tương thích mà không chặn khởi động.
• Cập nhật bộ lọc tải xuống podcast để bao gồm các định dạng âm thanh/video mới.
• Ngăn Ctrl+S lưu các tệp âm thanh/video để tránh hỏng.
• Cải thiện nhập bản ghi YouTube làm cho nó mạnh mẽ và linh hoạt hơn.
• Cải thiện độ bền của việc chia sách nói thành các phần, đảm bảo không mất văn bản nào.
• Trình cài đặt giờ hoàn toàn đa ngôn ngữ, hỗ trợ tiếng Ý, Anh, Tây Ban Nha, Bồ Đào Nha, Thụy Điển và Việt Nam dựa trên ngôn ngữ hệ thống của người dùng. Tiếng Anh là mặc định cho các hệ thống không được hỗ trợ.
• Danh mục podcast: nhấn Enter trên một danh mục giờ xác nhận lựa chọn (tương đương nút OK).
• Cải thiện hệ thống phát hiện treo để tránh dương tính giả khi có hộp thoại modal mở (thông báo lỗi, "không tìm thấy văn bản").
Sửa lỗi
• Sửa lỗi changelog không mở khi khởi động.
• Sửa lỗi yêu cầu OCR không xuất hiện cho PDF không thể truy cập được mở từ Explorer.
• Sửa lỗi khởi động có thể gây mất tiêu điểm hoặc đóng cửa sổ ngay sau khi mở.
• Sửa lỗi nghiêm trọng trong tìm kiếm regex ngăn tìm văn bản, bao gồm các vấn đề với "Tìm kiếm vòng" và tùy chọn "Dấu chấm tương đương dòng mới" với kết thúc dòng Windows.
Bản địa hóa
• Thêm bản dịch tiếng Ba Lan.
• Thêm bản dịch tiếng Pháp.
• Thêm bản dịch tiếng Séc (cảm ơn Radek Žalud và Jiri Holzinger).

Phiên bản 0.6.1 – 2026-01-20

Sửa lỗi

• Đã sửa lỗi khi bật “Hiển thị giọng nói trong trình soạn thảo” và phát podcast thì quá trình phát bị dừng.

• Đã sửa lỗi khiến một số podcast không thể thêm bằng URL do địa chỉ bị cắt ngắn.

• Đã sửa lỗi không thể thêm các URL thông thường trong chức năng RSS feed.

• Đã sửa lỗi khiến ngôn ngữ Wikipedia hiển thị ở nhiều tab cài đặt khác nhau.

• Đã loại bỏ việc tạo một số tệp debug vốn vẫn được tạo ngay cả ở chế độ release.

Cải tiến

• Cải thiện hỗ trợ cho giọng nói Microsoft, hiện được phát bằng phương thức chuyên biệt với user agent khác.

• Đã thêm hỗ trợ cho tệp MP4.



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
• Thêm tùy chọn để bật "Mở bằng Sonarpad" trong File Explorer, bao gồm cả các bản portable.
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
• Loại bỏ cài đặt di chuyển cài đặt sang thư mục hiện tại; ứng dụng hiện tự động xử lý việc này dựa trên vị trí (nếu thư mục chứa file exe tên là "sonarpad portable" hoặc nằm trên ổ đĩa di động, cài đặt sẽ vào thư mục `config` của exe, nếu không sẽ vào `%APPDATA%\Sonarpad`).

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
- Thêm các tùy chọn mở bằng Sonarpad và liên kết tệp trong khi cài đặt.
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











