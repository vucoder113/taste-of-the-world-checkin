# data-kh-checkin

Trang báo cáo khách tham dự Taste of the World. Chọn bốn file Excel riêng để xem số đăng ký, check-in buổi chiều, khách mời/check-in tiệc tối và kết quả theo nhân viên.

## Chạy local

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Mở `http://127.0.0.1:5000`.

## Triển khai Render

Repository có `render.yaml` để tạo Web Service chạy bằng Gunicorn. Từ Render, chọn **New > Blueprint** và kết nối repository này. Cấu hình bật `REQUIRE_AUTH=true`, Render tạo `SECRET_KEY`, và cần cung cấp hai biến bí mật `APP_PASSWORD` và `DATABASE_URL`.

Để lấy `DATABASE_URL`, trong Supabase mở **Connect > Connection string**, chọn dạng **URI** và **Transaction pooler**, rồi sao chép URI vào biến `DATABASE_URL` trong Render. Thay placeholder mật khẩu bằng mật khẩu database của dự án Supabase; ứng dụng kết nối qua TLS. Không gửi URI hoặc mật khẩu cho người khác và không lưu chúng vào GitHub. Không cần tạo bảng thủ công: ứng dụng tự tạo schema riêng `app_private` và bảng lưu báo cáo khi kết nối lần đầu; schema và bảng bị thu hồi quyền truy cập công khai.

Sau khi Render triển khai thành công, đăng nhập và tải lên bốn file Excel một lần. Từ đó báo cáo được nạp lại từ Supabase khi website mở hoặc Render khởi động lại. Có thể tải bộ Excel mới lên để thay thế báo cáo đã lưu. File khách mời được lưu trong database Supabase; chỉ dùng dự án và thông tin kết nối được bảo vệ, và chỉ chia sẻ URL/mật khẩu website với người được phép.

## Các file Excel

Trang nhận bốn file `.xlsx` trong cùng một lượt:

- Danh sách khách hàng đăng ký tham dự.
- Danh sách khách check-in buổi chiều (mọi dòng trong file đều là khách đã check-in).
- Danh sách khách mời tiệc tối.
- Danh sách chỉ tiêu nhân viên.

Ứng dụng nhận dạng từng file theo tên/cột, kể cả khi tiêu đề bảng nằm dưới dòng tiêu đề lớn. Báo cáo giữ riêng các danh sách, không gộp/deduplicate khách giữa các file. Mọi dòng trong file check-in buổi chiều được tính là khách đã check-in. Số khách tiệc tối được tính theo cột `SL Khách`; check-in tối chỉ tính số lượng ở các dòng có dấu trong cột `Check in`. Dòng thiếu số lượng vẫn hiện trong dữ liệu chi tiết nhưng không bị tự tính thành một khách; trang sẽ cảnh báo số dòng cần rà soát. Nếu file không có thông tin VIP, biểu đồ đánh dấu là “Chưa phân loại”, không tự gán thành không VIP. Có thể mở chi tiết nhân viên hoặc nhóm “Chưa xác định” để xem các trường dữ liệu gốc.

Các ô trạng thái trong chi tiết khách được đối chiếu giữa danh sách đăng ký, buổi chiều và tiệc tối theo số điện thoại/email duy nhất; chỉ dùng tên khi tên duy nhất ở cả hai danh sách và thông tin liên hệ không mâu thuẫn. Ô check-in tối dựa trên dấu `Check in` của danh sách tiệc tối. Số liệu tổng hợp theo nhân viên gán tên đầy đủ khớp chính xác hoặc bí danh viết tắt khớp duy nhất với tên/chức vụ trong file Target; “TGĐ”, “TDG” và “Sếp tùng” được quy về ANH TÙNG theo cách ghi trong dữ liệu. Tên không khớp hoặc mơ hồ được cộng vào dòng “Chưa đối chiếu Target”, không tự suy đoán nhân viên. Bảng nhân viên được xếp theo phòng ban rồi theo tên; tỷ lệ tham dự là `(check-in chiều + check-in tối) / max(đăng ký, target)` của từng nhân viên: nếu đăng ký thấp hơn Target thì dùng Target làm mẫu số, nếu không thì dùng số đăng ký. Hiển thị `—` nếu cả đăng ký và Target đều bằng 0.

Khi chạy trực tiếp bằng `python app.py`, ứng dụng tự nạp các file `.xlsx` trong `~/Downloads/15.10` nếu tìm thấy đủ bốn file nhận diện được. Có thể đổi thư mục bằng biến môi trường `REPORT_DATA_DIR`. Workbook tổng hợp bốn sheet của phiên bản cũ vẫn được nhận qua API.

## Riêng tư dữ liệu

Workbook tải lên được đọc trong bộ nhớ; ứng dụng chỉ lưu báo cáo đã phân tích vào schema riêng trong Supabase, không lưu workbook gốc. Chạy local không cấu hình `DATABASE_URL` thì báo cáo hiện tại chỉ tồn tại trong bộ nhớ đến khi server khởi động lại. Bản cấu hình Render yêu cầu đăng nhập bằng `APP_PASSWORD`; thay mật khẩu nếu bị lộ và không chia sẻ thông tin đăng nhập database qua GitHub hoặc mã nguồn.
