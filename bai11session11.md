#  BỘ HỒ SƠ KỸ THUẬT: HỘI NGHỊ KHOA HỌC SINH VIÊN PTIT 2027


## PHẦN 1: THIẾT KẾ KIẾN TRÚC HỆ THỐNG

### 1. Sơ đồ Luồng Dữ liệu (Data Flow) & Phân định Nguồn - Hiển thị

Hệ thống quản lý và vận hành Hội nghị Khoa học Sinh viên PTIT 2027 được thiết kế theo mô hình tích hợp khép kín giữa các ứng dụng văn phòng, trong đó phân định rõ ràng giữa **Nguồn dữ liệu gốc (Source)** và **Các kênh hiển thị/báo cáo (Display)**:

```text
[ CSDL NGUỒN: Google Sheets / Excel ]
 ├── Sheet 1: DanhMucPhong (Đơn giá, Phòng, Chủ trì)  ──┐
 └── Sheet 2: DanhSachBaoCao (Mã TG, Tên, Khoa,...)      │  (Hàm VLOOKUP tự động tra cứu)
                                                        ▼
                                       [ BẢNG THEO DÕI & LỌC DỮ LIỆU (EXCEL) ]
                                                        │
                  ┌─────────────────────────────────────┴─────────────────────────────────────┐
                  ▼                                                                           ▼
       [ TÀI LIỆU KẾ HOẠCH (WORD) ]                                              [ SLIDE KHAI MẠC (POWERPOINT) ]
       - Nguồn: Soạn thảo thủ công / Markdown                                    - Nguồn: Trích xuất số liệu tổng hợp từ Excel
       - Hiển thị: Văn bản in phát Ban Giám khảo                                - Hiển thị: Trình chiếu trực quan cho 500 đại biểu
       - Cấu trúc: Auto TOC, Heading Styles                                      - Cấu trúc: Rule 10/20/30, Assertion Titles

```

* **Thành phần Nguồn (Source):** Bảng tính Excel (`DanhSachBaoCao` và `DanhMucPhong`) đóng vai trò là cơ sở dữ liệu gốc (Database). Mọi thông tin về phân bổ phòng, thù lao, danh sách báo cáo viên đều được lưu trữ tập trung tại đây.
* **Thành phần Hiển thị (Display):**
* *Tài liệu Word (A):* Kênh hiển thị văn bản mang tính pháp lý, thể hiện kế hoạch chi tiết cho Ban Giám khảo nghiên cứu trước ngày hội nghị.
* *Slide PowerPoint (C):* Kênh hiển thị trực quan, chắt lọc các chỉ số cốt lõi từ Excel để trình chiếu trước toàn thể đại biểu trong phiên khai mạc.



---

### 2. Cấu trúc Phân cấp Heading cho Tài liệu Word (Kế hoạch Hội nghị 2027)

Tài liệu Kế hoạch được thiết kế tuân thủ cấu trúc phân cấp từ Cấp 1 (H1) đến Cấp 3 (H3) nhằm phục vụ tính năng tạo Mục lục tự động (Auto TOC):

* `# 1. Giới thiệu Chung về Hội nghị Khoa học Sinh viên PTIT 2027` (H1)
* `## 1.1. Mục đích và Ý nghĩa Khoa học` (H2)
* `### 1.1.1. Định hướng đổi mới sáng tạo trong sinh viên` (H3)
* `### 1.1.2. Thúc đẩy hợp tác nghiên cứu giữa các viện/khoa` (H3)


* `## 1.2. Chủ đề trọng tâm năm 2027` (H2)


* `# 2. Cơ cấu Tổ chức và Ban Chỉ đạo` (H1)
* `## 2.1. Hội đồng Khoa học và Ban Giám khảo` (H2)
* `## 2.2. Ban Thư ký và Ban Kỹ thuật vận hành` (H2)


* `# 3. Lịch trình và Thời gian Triển khai` (H1)
* `## 3.1. Giai đoạn tiếp nhận và phản biện hồ sơ` (H2)
* `## 3.2. Chương trình chi tiết ngày hội nghị chính thức` (H2)


* `# 4. Cơ chế Khen thưởng và Giải thưởng` (H1)
* `## 4.1. Cơ cấu giải thưởng chính thức` (H2)
* `## 4.2. Quyền lợi và kinh phí hỗ trợ công bố khoa học` (H2)


* `# 5. Quy chế Làm việc và Đảm bảo Kỹ thuật` (H1)
* `## 5.1. Quy định đối với báo cáo viên và tác giả` (H2)
* `## 5.2. Phương án dự phòng sự cố kỹ thuật phòng họp` (H2)



---

## PHẦN 2: SẢN PHẨM HOÀN CHỈNH

### 1. Tài liệu Kế hoạch (Word - Thành phần A)

*Tài liệu được định dạng hoàn toàn bằng Heading Styles (Không dùng định dạng thủ công).*

```markdown
# 1. Giới thiệu Chung về Hội nghị Khoa học Sinh viên PTIT 2027

Hội nghị Khoa học Sinh viên Học viện Công nghệ Bưu chính Viễn thông (PTIT) năm 2027 là sự kiện thường niên nhằm thúc đẩy phong trào nghiên cứu khoa học, ứng dụng công nghệ số và trí tuệ nhân tạo trong khối sinh viên toàn học viện.

## 1.1. Mục đích và Ý nghĩa Khoa học
### 1.1.1. Định hướng đổi mới sáng tạo trong sinh viên
Tạo môi trường học thuật để sinh viên tiếp cận các bài toán thực tế từ doanh nghiệp, phát triển tư duy giải quyết vấn đề thông qua các đề tài nghiên cứu chuyên sâu.

### 1.1.2. Thúc đẩy hợp tác nghiên cứu giữa các viện/khoa
Xây dựng mạng lưới kết nối giữa sinh viên các ngành Công nghệ thông tin, Viễn thông, Điện tử và Quản trị kinh tế.

## 1.2. Chủ đề trọng tâm năm 2027
Chủ đề chính: *"Chuyển đổi số và Trí tuệ nhân tạo kiến tạo tương lai số bền vững"*.

---

# 2. Cơ cấu Tổ chức và Ban Chỉ đạo

## 2.1. Hội đồng Khoa học và Ban Giám khảo
Hội đồng bao gồm các Giáo sư, Phó Giáo sư, Tiến sĩ đầu ngành từ các Khoa chuyên môn và chuyên gia cố vấn công nghệ từ các tập đoàn đối tác chiến lược.

## 2.2. Ban Thư ký và Ban Kỹ thuật vận hành
* Trị sự chung: Phòng Khoa học Công nghệ & Hợp tác Quốc tế.
* Ban Kỹ thuật: Phụ trách hạ tầng mạng, hệ thống trực tuyến, trực phòng họp và điều phối thiết bị trình chiếu.

---

# 3. Lịch trình và Thời gian Triển khai

## 3.1. Giai đoạn tiếp nhận và phản biện hồ sơ
1. Tiếp nhận tóm tắt đề tài (Abstract): Từ ngày 01/10/2026 đến hết ngày 30/11/2026.
2. Vòng phản biện kín (Blind Review): Từ ngày 01/12/2026 đến 31/12/2026.
3. Hoàn thiện bản thảo toàn văn (Full-paper): Trước ngày 15/02/2027.

## 3.2. Chương trình chi tiết ngày hội nghị chính thức
Thời gian diễn ra: Ngày 20/03/2027 tại Hội trường lớn cơ sở Hà Nội.

---

# 4. Cơ chế Khen thưởng và Giải thưởng

## 4.1. Cơ cấu giải thưởng chính thức
* 01 Giải Nhất: Cúp lưu niệm, Giấy khen Giám đốc PTIT và tiền thưởng 10.000.000 VNĐ.
* 02 Giải Nhì: Giấy khen và tiền thưởng 6.000.000 VNĐ/giải.
* 03 Giải Ba: Giấy khen và tiền thưởng 3.000.000 VNĐ/giải.
* 05 Giải Khuyến khích: Tiền thưởng 1.000.000 VNĐ/giải.

## 4.2. Quyền lợi và kinh phí hỗ trợ công bố khoa học
Các đề tài đạt giải Nhất và Nhì được hỗ trợ toàn bộ chi phí đăng ký bài báo trên các tạp chí khoa học quốc tế uy tín thuộc danh mục Scopus/ISI.

---

# 5. Quy chế Làm việc và Đảm bảo Kỹ thuật

## 5.1. Quy định đối với báo cáo viên và tác giả
* Mỗi tác giả có tối đa 15 phút trình bày trực tiếp bằng tiếng Việt hoặc tiếng Anh và 05 phút hỏi đáp cùng hội đồng.
* Slide trình bày phải sử dụng ngôn ngữ chuẩn, không chứa nội dung phản cảm hoặc vi phạm sở hữu trí tuệ.

## 5.2. Phương án dự phòng sự cố kỹ thuật phòng họp
Trang bị hệ thống máy chủ dự phòng, kết nối mạng cáp quang kép và nguồn điện UPS đảm bảo không gián đoạn quá trình trình bày của báo cáo viên.

```

---

### 2. Bảng theo dõi Báo cáo viên (Excel - Thành phần B)

Hệ thống bảng tính gồm 2 sheet được thiết kế tuân thủ tuyệt đối chuẩn Database (Không merge ô tiêu đề, không có hàng/cột trống ngắt quãng, dòng tổng đặt ngoài bảng).

#### Sheet 1: `DanhMucPhong` (Bảng tra cứu nguồn)

| Mã Phòng | Tên Phòng / Hội Trường | Chủ Trì / Phản Biện Chính | Thiết Bị Hỗ Trợ | Đơn Giá Hỗ Trợ (VNĐ) |
| --- | --- | --- | --- | --- |
| P201 | Hội trường lớn A | PGS.TS Nguyễn Văn A | Máy chiếu 4K, Âm thanh vòm | 1.500.000 |
| P202 | Phòng họp chuyên đề B | TS. Trần Thị B | Màn hình LED 85 inch | 1.200.000 |
| P301 | Phòng hội thảo C | PGS.TS Lê Văn C | Hệ thống Hybrid trực tuyến | 1.000.000 |

#### Sheet 2: `DanhSachBaoCao` (Bảng quản lý chính)

*Công thức tại cột **Phòng** (Cột E):* `=IFERROR(VLOOKUP(C2, DanhMucPhong!$A$2:$E$10, 2, FALSE), "Chưa phân bổ")`

*Công thức tại cột **Thù lao** (Cột F):* `=IFERROR(VLOOKUP(C2, DanhMucPhong!$A$2:$E$10, 5, FALSE), 0)`

| Mã TG (A) | Tên Báo Cáo Viên (B) | Mã Phòng (C) | Tên Báo Cáo (D) | Phòng (E) | Thù Lao (VNĐ) (F) |
| --- | --- | --- | --- | --- | --- |
| TG01 | Nguyễn Văn An | P201 | Nghiên cứu ứng dụng IoT trong nông nghiệp | Hội trường lớn A | 1.500.000 |
| TG02 | Trần Thị Bình | P202 | Thuật toán tối ưu hóa mạng 6G | Phòng họp chuyên đề B | 1.200.000 |
| TG03 | Lê Minh Châu | P301 | Phân tích cảm xúc văn bản tiếng Việt | Phòng hội thảo C | 1.000.000 |
| TG04 | Phạm Hoàng Nam | P201 | Bảo mật chuỗi khối trong tài chính số | Hội trường lớn A | 1.500.000 |

* **Ví dụ Lọc đa điều kiện với AutoFilter (`Ctrl + Shift + L`):**
* *Điều kiện 1:* Lọc cột **Mã Phòng** (Cột C) chọn `P201`.
* *Điều kiện 2:* Lọc cột **Thù Lao** (Cột F) chọn *Number Filters -> Greater Than* `1.000.000`.
* *Kết quả:* Hệ thống lọc ra chính xác các báo cáo viên trình bày tại phòng P201 có mức thù lao lớn hơn 1 triệu đồng.



---

### 3. Kịch bản Slide Khai mạc (PowerPoint - Thành phần C)

*Tuân thủ tuyệt đối Quy tắc 10/20/30 (tối đa 8 slides, cỡ chữ tối thiểu 30pt, Assertion Titles).*

* **Slide 1:**
* **Tiêu đề (Assertion Title):** `Hội Nghị Khoa Học Sinh Viên PTIT 2027: Khởi Nguồn Sáng Tạo Số`
* **Nội dung:** Chào mừng 500 đại biểu và báo cáo viên đến với sự kiện học thuật lớn nhất năm của học viện.
* **Cỡ chữ:** Tiêu đề 40pt, Nội dung 32pt.


* **Slide 2:**
* **Tiêu đề (Assertion Title):** `Quy Mô Tổ Chức Đạt Mức Kỷ Lục Với Hơn 150 Đề Tài Nghiên Cứu`
* **Nội dung:** Số lượng bài báo tăng 35% so với năm 2026, quy tụ sinh viên từ 4 cơ sở đào tạo trọng điểm.
* **Cỡ chữ:** Tiêu đề 38pt, Nội dung 32pt.


* **Slide 3:**
* **Tiêu đề (Assertion Title):** `Khối Lượng Đề Tài Tập Trung Mạnh Mẽ Vào Lĩnh Vực Trí Tuệ Nhân Tạo`
* **Loại biểu đồ:** **Donut Chart** (Biểu đồ tỷ trọng).
* **Nội dung hiển thị tỷ lệ:** CNTT & AI (45%), Viễn thông (30%), Điện tử - Tự động hóa (15%), Quản trị số (10%).
* **Cỡ chữ:** Tiêu đề 38pt, Chú thích biểu đồ 30pt.


* **Slide 4:**
* **Tiêu đề (Assertion Title):** `Sự Tăng Trưởng Vượt Bậc Của Các Khoa Về Số Lượng Bài Báo Đạt Giải`
* **Loại biểu đồ:** **Bar Chart** (Biểu đồ cột so sánh).
* **Nội dung hiển thị:** So sánh số lượng đề tài đạt giải giữa Khoa CNTT, Khoa Viễn thông và Khoa Điện tử qua các năm 2025 - 2027.
* **Cỡ chữ:** Tiêu đề 38pt, Trục tọa độ 30pt.


* **Slide 5:**
* **Tiêu đề (Assertion Title):** `Hội Đồng Giám Khảo Gồm Các Chuyên Gia Đầu Ngành Và Doanh Nghiệp`
* **Nội dung:** 100% thành viên hội đồng có học hàm PGS, TS và đại diện R&D từ các tập đoàn công nghệ lớn.
* **Cỡ chữ:** Tiêu đề 38pt, Nội dung 32pt.


* **Slide 6:**
* **Tiêu đề (Assertion Title):** `Tổng Ngân Sách Giải Thưởng Và Hỗ Trợ Công Bố Đạt 50 Triệu Đồng`
* **Nội dung:** Tạo động lực mạnh mẽ để sinh viên đẩy mạnh công bố quốc tế.
* **Cỡ chữ:** Tiêu đề 38pt, Nội dung 32pt.


* **Slide 7:**
* **Tiêu đề (Assertion Title):** `Lộ Trình Tiếp Theo: Hướng Tới Giải Thưởng Khoa Học Công Nghệ Toàn Quốc`
* **Nội dung:** Các đề tài xuất sắc nhất tại hội nghị sẽ được lựa chọn tham gia Giải thưởng "Nức lòng sinh viên nghiên cứu khoa học" cấp Bộ.
* **Cỡ chữ:** Tiêu đề 38pt, Nội dung 32pt.


* **Slide 8 (Call-to-Action):**
* **Tiêu đề (Assertion Title):** `Thông Báo Quan Trọng: Kính Mời Đại Biểu Di chuyển Về Các Phòng Chuyên Đề`
* **Nội dung Call-to-Action:**
1. Phiên song song bắt đầu lúc **09:30 AM** tại các phòng P201, P202, P301.
2. Quét mã QR trên màn hình để tải kỷ yếu điện tử đầy đủ.


* **Cỡ chữ:** Tiêu đề 40pt, Nội dung hành động 34pt.



---

## PHẦN 3: GIẢI QUYẾT CÁC TRƯỜNG HỢP KIỂM THỬ (EDGE CASES)

### 1. Xử lý Edge Case 1 (Word - Thay đổi hàng loạt màu Heading 1 qua 40 trang)

* **Tình huống:** Trưởng ban yêu cầu đổi toàn bộ màu của tất cả các tiêu đề Heading 1 từ màu đỏ sang màu xanh navy xuyên suốt tài liệu dài 40 trang.
* **So sánh thao tác:**
* **Nếu dùng Định dạng thủ công (Manual Formatting):** Trưởng ban hoặc nhân sự phải thủ công cuộn qua từng trang, tìm từng dòng tiêu đề Heading 1, bôi đen và chọn đổi màu. Với tài liệu 40 trang chứa khoảng 15 - 20 chương/mục lớn rải rác, thao tác này đòi hỏi **hơn 60 đến 80 click chuột và thao tác thủ công**, rất dễ bỏ sót các tiêu đề ở giữa các trang, dẫn đến sự thiếu đồng bộ.
* **Nếu dùng Heading Styles chuẩn:** Chỉ cần **đúng 1 thao tác duy nhất**: Nhấp chuột phải vào mẫu style **Heading 1** trên thanh Ribbon (Tab *Home -> Styles*), chọn **Modify**, chọn màu *Navy* và nhấn **OK**. Toàn bộ Heading 1 trên toàn bộ 40 trang lập tức đổi màu đồng loạt trong vòng chưa đầy **2 giây**.



---

### 2. Xử lý Edge Case 2 (Excel - 5 báo cáo viên mời ngoài chưa có trong `DanhMucPhong`)

* **Tình huống:** Khi dùng hàm `VLOOKUP` tra cứu thông tin phòng và thù lao cho 5 báo cáo viên mời ngoài (chưa được cập nhật trong sheet `DanhMucPhong`), hàm thông thường sẽ trả về lỗi `#N/A` làm mất thẩm mỹ bảng dữ liệu.
* **Giải pháp kỹ thuật (Không sửa sheet nguồn):**
* Sử dụng hàm **`IFERROR()`** bao bọc bên ngoài hàm `VLOOKUP`.
* **Công thức áp dụng tại ô E2:**
`=IFERROR(VLOOKUP(C2, DanhMucPhong!$A$2:$E$10, 2, FALSE), "Chưa phân bổ")`
* **Công thức áp dụng tại ô F2:**
`=IFERROR(VLOOKUP(C2, DanhMucPhong!$A$2:$E$10, 5, FALSE), 0)`
* **Kết quả:** Thay vì hiển thị lỗi `#N/A` thô kệch, hệ thống sẽ tự động hiển thị chữ *"Chưa phân bổ"* cho cột Phòng và số `0` cho cột Thù lao đối với 5 báo cáo viên đặc biệt này, giúp bảng tính luôn sạch sẽ và sẵn sàng cho việc lọc dữ liệu.



---

### 3. Xử lý Edge Case 3 (Slide - Bảng số liệu chi tiết vi phạm Rule 10/20/30)

* **Tình huống:** Ban giám khảo yêu cầu phải cung cấp đầy đủ bảng số liệu chi tiết (gồm 50 hàng dữ liệu kinh phí và phân bổ đề tài), nhưng việc đưa toàn bộ bảng này lên Slide sẽ vi phạm nghiêm trọng Quy tắc 10/20/30 (chữ quá nhỏ, gây quá tải thông tin, biến slide thành "Death by PowerPoint").
* **Giải pháp xử lý chuyên nghiệp:**
* **Trên Slide chính (Trình chiếu):** Chỉ hiển thị biểu đồ tổng quan (Bar/Donut Chart) với các số liệu tóm tắt lớn, rõ ràng (Cỡ chữ ≥ 30pt) nhằm phục vụ 500 khán giả tại hội trường.
* **Giải pháp kỹ thuật phụ (Appendix / QR Code):**
1. Đính kèm một mã QR code lớn ở góc phải của Slide trình bày tóm tắt với tiêu đề: *"Quét mã để xem chi tiết Bảng số liệu toàn văn"*.
2. Mã QR dẫn trực tiếp đến file Excel hoặc Google Docs chứa đầy đủ 50 hàng dữ liệu trên Google Drive của ban tổ chức.


* **Lợi ích:** Đảm bảo slide chính cực kỳ tinh gọn, súc tích, tuân thủ tuyệt đối quy tắc trình chiếu hiện đại, đồng thời cung cấp đầy đủ thông tin chi tiết ngay lập tức cho những giám khảo có nhu cầu tra cứu sâu.