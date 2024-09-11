import requests
import json
import time

error_count = 0
max_error_count = 2

with open("diemthi2024.txt", "w", encoding='utf-8') as f:
    for x in range(48000001, 49000000):
        try:
            data_url = "https://sgddt.dongnai.gov.vn/tracuu/export/" + str(x) + ".json"
            response = requests.get(data_url, allow_redirects=True)

            if response.status_code == 200:
                info = response.json()

                if info and all(k in info for k in ('HO_TEN', 'SOBAODANH', 'DIEM_THI', 'NGAY_SINH', 'GIOI_TINH', 'CMND')):
                    thong_tin_hs = "Họ tên: {}  SBD: {}     Điểm: {}    Ngày sinh: {}   Giới tính: {}   Số cmnd: {}".format(
                        info['HO_TEN'], info['SOBAODANH'], info['DIEM_THI'], info['NGAY_SINH'], info['GIOI_TINH'], info['CMND'])
                    f.write(thong_tin_hs + "\n")
                    print("Đang dò điểm thi của số báo danh: " + str(x))

                error_count = 0
                time.sleep(0.1)
            else:
                print(f"Số báo danh {x} không có dữ liệu.")
                error_count += 1

        except requests.exceptions.RequestException as e:
            print(f"Yêu cầu lỗi với SBD {x}: {e}")
            error_count += 1

        if error_count > max_error_count:
            print("Đã dò hết số báo danh có trong tỉnh Đồng Nai!")
            break
