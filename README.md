# dss_py_control
Tập hợp các mã lệnh phục vụ thực hiện điều khiển lưới trong OpenDSS bằng Python

Quy trình thao tác:
Bước 1: Thay đổi cài đặt thông số mô phỏng trong file settings với các thông số như sau:

OPENDSS_LINK= r"đường_dẫn_tới_folder_lưới_mô_phỏng" # chú ý có "" và r
ALGO_LINK = r"đường_dẫn_tới_thuật_toán" # chú ý có "" và r
STEPSIZE = "thời_gian" #Thời gian ở đây gồm giá trị + đơn vị (h/m): Ví dụ: 1h, 15m
NUMBERS = số bước mô phỏng # Giá trị là số thực
VIEW_TIMEPERSTEP = giá trị thời gian mỗi lần mô phỏng "đơn vị là giây


Về yêu cầu file thuật toán; Phải đảm bảo thuật toán trả về kết quả dữ liệu có các định dạng như sau

Với nguồn: “Tên P Q” (Ví dụ: Gen1 500 600) mặc định là kW
Với biến áp: “Tên Phân_cấp_chỉnh” (VD: Reg1 0.97)
Với các phần tử đóng cắt: “Tên Trạng_thái” (VD: Switch1 Active; Switch2 Noactive)

Khi thực hiện chạy file main.py, kết quả trả về gồm có 2 file:
1. Mã OpenDSS phục vụ nhập vào OpenDSS-G #Điều này nhằm mục đích để kết quả hiển thị chi tiết hơn; hỗ trợ được thêm bởi OpenDSS-Viewer
2. Folder csv chứa toàn bộ monitor tất cả các phần tử
