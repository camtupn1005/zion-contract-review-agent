# ZION Contract Review Agent

AI agent rà soát và so sánh hai hợp đồng (.docx hoặc .pdf), xuất báo cáo tiếng Việt dạng Word.

## Yêu cầu

- Python 3.10+
- Google Gemini API Key ([lấy tại đây](https://aistudio.google.com/app/apikey))

## Cài đặt

```bash
# 1. Cài thư viện
pip install -r requirements.txt

# 2. Tạo file .env và nhập API key
echo GEMINI_API_KEY=your_api_key_here > .env
```

## Sử dụng

```bash
# Cơ bản (báo cáo tự động đặt tên theo timestamp)
python agent.py hop_dong_1.pdf hop_dong_2.docx

# Chỉ định tên file đầu ra
python agent.py hop_dong_1.pdf hop_dong_2.docx --output BaoCao_ThangXX.docx
```

## Đầu ra

File `.docx` chứa báo cáo rà soát gồm các phần:

1. Thông tin cơ bản hợp đồng
2. Các bên tham gia
3. Đối tượng và phạm vi hợp đồng
4. Giá trị và điều khoản thanh toán
5. Thời hạn hợp đồng
6. Quyền và nghĩa vụ các bên
7. Điều khoản phạt và bồi thường
8. Điều khoản chấm dứt hợp đồng
9. Giải quyết tranh chấp
10. Điểm khác biệt giữa hai hợp đồng (bảng so sánh)
11. Rủi ro pháp lý và khuyến nghị

## Cấu trúc project

```
zion-contract-review-agent/
├── agent.py          # Agent chính
├── config.yaml       # Cấu hình model, output, review sections
├── requirements.txt  # Thư viện Python
├── .env              # API key (không commit lên git)
└── README.md
```

## Cấu hình

Chỉnh `config.yaml` để thay đổi model, nhiệt độ (temperature), hoặc các mục trong báo cáo.

```yaml
llm:
  model: "gemini-2.0-flash"   # hoặc gemini-1.5-pro
  temperature: 0.2            # 0 = chính xác, 1 = sáng tạo hơn
```
