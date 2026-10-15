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

Repository có `render.yaml` để tạo Web Service chạy bằng Gunicorn. Từ Render, chọn **New > Blueprint** và kết nối repository này. Trước khi mở trang, đặt giá trị bí mật cho biến môi trường `APP_PASSWORD`; `SECRET_KEY` được Render tạo tự động và `REQUIRE_AUTH=true` bật trang đăng nhập. Không lưu mật khẩu hoặc dữ liệu Excel trong GitHub.

Mỗi lần triển khai, khởi động lại hoặc scale service có thể làm mất báo cáo đang giữ trong bộ nhớ. Sau khi mở website, đăng nhập rồi tải lên lại bốn file Excel. Bản miễn phí của Render có thể ngủ khi không hoạt động. Dùng dữ liệu khách thật chỉ khi tài khoản/mật khẩu được giữ riêng và chỉ chia sẻ URL với người được phép.

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

Workbook tải lên được đọc trong bộ nhớ và không được ghi vào đĩa. Báo cáo nạp tự động được đọc từ thư mục đã cấu hình; báo cáo hiện tại chỉ tồn tại trong bộ nhớ đến khi server khởi động lại. Bản cấu hình Render bật đăng nhập bằng `APP_PASSWORD`; thay mật khẩu nếu bị lộ và không chia sẻ mật khẩu qua GitHub hoặc mã nguồn.
