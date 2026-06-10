# AI Chat CLI với RAG & LM Studio

Đây là một ứng dụng Command-Line Interface (CLI) siêu tốc độ, cho phép bạn chat trực tiếp với các file mã nguồn, tài liệu (PDF, Word, Excel, PowerPoint, Text) chạy hoàn toàn trên máy tính cá nhân (Local) thông qua **LM Studio**.

Hệ thống được tích hợp công nghệ **RAG (Retrieval-Augmented Generation)** dựa trên ChromaDB để tự động tìm kiếm ngữ cảnh, cho phép AI trả lời các câu hỏi dựa trên nội dung thực tế từ tài liệu của bạn với độ chính xác cao.

---

## 🌟 Tính năng nổi bật
* **Hỗ trợ đa định dạng file:** `.pdf`, `.docx`, `.xlsx`, `.pptx`, `.ipynb`, `.py`, `.js`, `.java`, `.txt`, `.md`, và hơn thế nữa.
* **Smart RAG (Vector Search):** Tự động cắt nhỏ tài liệu, dán nhãn, lưu vào Database và tìm kiếm theo ngữ nghĩa cực kỳ chính xác. Nhận diện file tự động qua cách bạn chat (Fuzzy Match).
* **Tối ưu tốc độ chạm nóc (Performance Tweaks):** 
  * Tận dụng **Prefix Caching** của LM Studio (Prefill siêu nhanh).
  * HTTP **Keep-Alive** giúp bỏ qua độ trễ bắt tay mạng (TCP Handshake).
* **Bypass An Toàn (Uncensored):** Persona mặc định là "Elite Cybersecurity Expert", cho phép trả lời sâu về các vấn đề bảo mật (SQLi, XSS, v.v.) mà không bị từ chối phục vụ.
* **Hỗ trợ Model Suy Luận (Reasoning):** Tự động phát hiện và in luồng suy nghĩ (`reasoning_content`) của các model thế hệ mới.
* **Chuyển đổi Model Thời Gian Thực:**
  * Lệnh `fast`: Chuyển sang mô hình 4B (Hỏi đáp nhanh như chớp).
  * Lệnh `normal`: Chuyển sang mô hình 12B (Lập luận sâu, phân tích logic).

---

## 🚀 Cài đặt & Chuẩn bị

### 1. Cài đặt môi trường (Chỉ cần 1 click)
Yêu cầu máy tính đã cài đặt sẵn **Python 3.10+**. (Lưu ý: Nhớ tick vào nút "Add Python to PATH" lúc cài đặt Python).

Thay vì phải gõ lệnh thủ công, bạn chỉ cần:
1. Click đúp vào file **`install.bat`**.
2. Ngồi đợi 1-2 phút. Phần mềm sẽ tự động tạo Môi trường ảo (Virtual Environment) và tải toàn bộ các thư viện cần thiết một cách gọn gàng, không rác máy.

### 2. Cài đặt LM Studio
1. Tải và cài đặt [LM Studio](https://lmstudio.ai/).
2. Trong thanh tìm kiếm của LM Studio, tải về 2 mô hình sau (tùy vào phần cứng của bạn, có thể thay đổi trong file `chat.py`):
   * `google/gemma-4-e4b` (Dành cho lệnh `fast`).
   * `google/gemma-4-12b-qat` (Dành cho lệnh `normal`).
3. Chuyển sang tab **Developer** (biểu tượng code `</>`) ở cột bên trái LM Studio.
4. Chọn **Local Server** -> Bật (Turn On) Server.
   * *Đảm bảo địa chỉ là: `http://127.0.0.1:1234`*.
   * Trong mục Hardware Settings bên phải, kéo **GPU Offload** lên mức cao nhất có thể, bật **Flash Attention** để chạy nhanh hơn.

---

## 💻 Cách sử dụng

Sau khi khởi động LM Studio Local Server, bạn click đúp vào file `start.bat` (hoặc chạy `python chat.py` từ Terminal) để khởi động.

### Giao diện lệnh (Lưu ý: Mọi lệnh gõ trực tiếp vào ô chat)

**1. Nạp tài liệu vào bộ não AI:**
* `load <đường_dẫn_file>`: Thêm 1 file vào DB.
* `load <file1> <file2>`: Thêm nhiều file cùng lúc. Hỗ trợ cả dấu ngoặc kép `"C:\path\to\file.pdf"`.
* `loadfolder <đường_dẫn_folder>`: Thêm toàn bộ các file trong thư mục vào DB.

**2. Đọc và Hỏi:**
* `read <tên_file_gợi_nhớ>`: Lấy TOÀN BỘ nội dung file ép AI đọc và tóm tắt. (Không cần gõ đúng tên file, gõ sai/thiếu nó vẫn tự tìm được file khớp nhất. VD: `read cheatsheet`).
* `ask <tên_file> <câu hỏi>`: Hỏi sâu về 1 file nhất định. Lọc toàn bộ nhiễu từ các file khác.
* *(Hoặc đơn giản là cứ gõ câu hỏi bình thường, hệ thống tự động tìm kiếm thông minh từ tất cả các tài liệu).*

**3. Quản lý AI & Database:**
* `fast`: Đổi sang chế độ phản hồi siêu tốc.
* `normal`: Đổi sang chế độ thông minh, suy luận sâu.
* `cleardb`: Dọn dẹp, xóa sạch toàn bộ bộ não tài liệu.
* `remove <tên_file>`: Xóa một file cụ thể khỏi DB.
* `files`: Xem danh sách tất cả các tài liệu AI đã đọc.
* `history` hoặc `clear`: Xem hoặc xóa lịch sử chat.
* `norag` / `userag`: Tắt/Bật tính năng đọc tài liệu (chỉ muốn chat chay).

---

## 🛠 Tùy chỉnh nâng cao (Dành cho Dev)
Mở file `chat.py`, ở phía trên cùng bạn có thể tùy chỉnh các cấu hình:
* `MODEL_FAST` và `MODEL_NORMAL`: Thay đổi tên Model nếu bạn tải model khác.
* `RAG_TOP_K`: Số lượng đoạn văn trích xuất mỗi lần hỏi (Mặc định: 5).
* `SYSTEM_PROMPT`: Đổi vai trò (Persona) của con AI.
* `HISTORY_TURNS`: Số câu chat giữ lại làm trí nhớ ngắn hạn.

---
**Chúc bạn code và nghiên cứu an toàn, hiệu quả!** 🔐
