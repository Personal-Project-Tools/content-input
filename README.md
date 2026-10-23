# TTK Content Tool

Tool local để tổng hợp chỉ số `TTK_Content` từ nhiều file Excel TikTok Shop Live/Video.

## Cách chạy

```powershell
python -m pip install -r requirements.txt
python app.py
```

Sau đó mở:

```text
http://127.0.0.1:8765
```

Trên Windows có thể chạy nhanh bằng cách double-click `run.bat`.

## Cách dùng

1. Chọn nhiều file `.xlsx` cùng lúc.
2. Bấm `Tính chỉ số`.
3. Tool sẽ gom theo `Live/Video` và `Tháng/Năm` lấy từ ô `A1`.
4. Bấm `Tải Excel` nếu muốn lưu bảng kết quả.

## Công thức hiện tại

Live:

- `sum(LIVE-attributed GMV)` = `revenue_24h`
- `sum(Views)` = `content_impression`
- `sum(Product Impressions)` = `product_impression`

Video:

- `sum(Video-attributed GMV)` = `revenue_24h`
- `sum(Video Views)` = `content_impression` (hỗ trợ cả cột cũ `VV`)
- `sum(Product Impressions)` = `product_impression`

Nếu file export bị mất tên header ở một vài cột tổng, tool sẽ fallback theo layout mẫu:

- Live: F = revenue, T = content impression, Z = product impression
- Video: V = revenue, G = content impression, M = product impression

## Chạy thử bằng folder

```powershell
python app.py --folder raw_sample
```
