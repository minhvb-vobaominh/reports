# PHS CR Management — Task Tracker & Dashboard

Pipeline chuẩn hóa dữ liệu quản lý tiến độ CR của phòng IT từ Base WeWork export, sinh báo cáo Excel + dashboard HTML + dữ liệu lịch sử cho BI.

```
raw export (.xls/.xlsx)
        │
        ▼
python convert_base_to_tracker.py <raw> template.xlsx <output>
        │
        ├── <output>.xlsx            Báo cáo Excel (Task Tracker, CR Overview, Dashboard Summary + star schema)
        ├── fact_task_snapshot.csv   Lịch sử snapshot theo ngày (append-only) — nguồn cho Power BI/Metabase
        └── dashboard-data.js        Dữ liệu task-level cho index.html
        │
        ▼
index.html  ← mở bằng trình duyệt, tự cập nhật theo lần chạy gần nhất
```

## 1. Yêu cầu môi trường

- Python 3.x với các package: `openpyxl`, `xlrd==1.2.0` (chỉ cần xlrd nếu input là `.xls` cũ)
- Trình duyệt bất kỳ để xem dashboard (cần internet lần đầu để tải Chart.js từ CDN)

```
pip install openpyxl "xlrd==1.2.0"
```

## 2. Quy trình sử dụng hàng tuần

1. **Export dữ liệu từ Base WeWork** (menu Export → Excel). Cả 2 định dạng đều được hỗ trợ:
   - `.xlsx` 36 cột (có Stage, Project Owner, Weekly-Review, Q-KPI, CR 1/CR 2…) — **khuyến nghị**
   - `.xls` 20 cột legacy (thiếu Stage → status sẽ fallback về trạng thái WeWork)
2. **Đặt file export vào thư mục này**, ví dụ `phs-cr-management.report.<date>.xlsx`
3. **Đóng file output cũ nếu đang mở trong Excel** (nếu không sẽ gặp `PermissionError`)
4. **Chạy script**:
   ```
   python convert_base_to_tracker.py phs-cr-management.report.<date>.xlsx template.xlsx phs-cr-management.report.<date>.output.xlsx
   ```
   Không truyền tham số thì mặc định là `task-input.xls` / `task-template.xlsx` / `task-output.xlsx`.
5. **Kiểm tra log**: `Column detection: 35/35 headers matched` — nếu thấp hơn nghĩa là Base đổi tên cột, cần cập nhật `HEADER_MAP` trong script.
6. **Mở kết quả**: file `.output.xlsx` cho báo cáo chi tiết, `index.html` cho dashboard trực quan.

> Chạy lại script **cùng ngày** sẽ thay thế snapshot của ngày đó trong `fact_task_snapshot.csv` (idempotent). Chạy đều mỗi tuần/ngày để tích lũy lịch sử cho trend/burn-up.

## 3. Các sheet trong file output

| Sheet | Nội dung |
|---|---|
| **Task Tracker** | Toàn bộ task đã chuẩn hóa: Request ID, Team, Status, **Stage**, Phase, deadline/completed (datetime thật)… |
| **CR Overview** | Mỗi dòng = 1 CR duy nhất, trạng thái từng team (Mobile / Dev B / MW), trạng thái tổng hợp, blockers + phần Summary tự tính |
| **Dashboard Summary** | Bảng COUNTIF theo team × status (Todo/Doing/Done/Overdue/Blocked/UAT/Rejected) và team × quý |
| **Team Config** | ⚙️ **Cấu hình team & lead email — sửa ở template, script tự đọc** |
| **Naming Convention** | Quy ước đặt tên task |
| **fact_task_snapshot** + `dim_task`/`dim_team`/`dim_stage` + `bridge_*` | Star schema cho BI: metric SLA, cycle/lead time, risk, multi-team; bridge tách các cột multi-value (division, platform, dev unit) |

## 4. Logic trạng thái (Stage quyết định)

Cột **Stage** trong raw phản ánh vòng đời CR và **ưu tiên hơn** trạng thái task WeWork:

| Stage | Status |
|---|---|
| Go-live, Ready for go-live | Done |
| UAT Testing (và stage chứa "uat/sit/testing") | **UAT** |
| FTL/Web B Coding, Design, Discuss Req, FSS Eval CR | Doing |
| Waiting FTL, Pending | Blocked |
| Backlog, Dev To do | Todo |
| Rejected | Rejected (loại khỏi thống kê dashboard) |

- Task WeWork bị `Quá hạn` sẽ **override** mọi stage chưa done → Overdue
- `Hoàn thành muộn` vẫn tính Done nhưng được giữ lại qua cờ `completed_late` trong fact
- Stage trống (file .xls cũ) → fallback về mapping trạng thái tiếng Việt

## 5. Quy ước đặt tên task (để parse đúng)

```
[CR] [2275303] [SS] Tên CR...          ← chuẩn
[CR] [2310394 - 2391899] [KT] ...      ← range ID (lấy ID đầu)
[SDK] [2400512] ...                    ← category SDK
[MKT] ...                              ← internal, không có CR
```

Team được xác định theo thứ tự: section header trong export → người giao việc (LEAD_MAP) → cột Project Owner → kế thừa từ task cha (subtask).

## 6. Dashboard HTML (`index.html`)

- Mở trực tiếp bằng double-click, **cần `dashboard-data.js` nằm cùng thư mục** (script tự sinh)
- Gồm: 6 thẻ KPI (Total/Done/Doing/UAT/Blocked/Overdue), pie phân bổ theo team, bảng workload, donut theo quý (Q3 tự hiện khi có dữ liệu), bar chart so sánh các quý theo team, và phần **Insights tự sinh** (CR quá hạn kèm số ngày trễ, task blocked kèm stage, backlog theo team, CR liên team)
- Không sửa số liệu trong HTML — mọi con số đều lấy từ `dashboard-data.js`

## 7. Tích hợp BI (Power BI / Metabase / Superset)

- Trỏ vào `fact_task_snapshot.csv` (grain: task × report_date) — đủ để vẽ trend, cumulative flow, burn-up theo `q_kpi`
- Join với các sheet `dim_*` / `bridge_*` trong file output để drill-down theo division/platform/dev unit

## 8. Thay đổi cấu hình

| Muốn thay đổi | Sửa ở đâu |
|---|---|
| Thêm/đổi team, lead email | Sheet **Team Config** trong `template.xlsx` |
| Map username người giao việc → team | `LEAD_MAP` trong script |
| Alias cột Project Owner | `PROJECT_OWNER_MAP` trong script |
| Stage mới → status | `STAGE_STATUS_MAP` / `STAGE_KEYWORD_RULES` (stage lạ đã có keyword fallback) |
| Thứ tự lifecycle stage | `STAGE_ORDER` |

## 9. Sự cố thường gặp

| Lỗi | Nguyên nhân & cách xử lý |
|---|---|
| `PermissionError ... output.xlsx` | File output đang mở trong Excel → đóng rồi chạy lại |
| Team = "Unknown" | Người giao việc không có trong LEAD_MAP và Project Owner trống → bổ sung LEAD_MAP hoặc điền Project Owner trong WeWork |
| `Column detection` < tổng số | Base đổi tên cột → cập nhật `HEADER_MAP` |
| Dashboard trắng | Thiếu `dashboard-data.js` (chưa chạy script) hoặc không có mạng để tải Chart.js |
| Request ID trống trong tracker | Tên task sai quy ước `[CR] [id]` → sửa tên task ở WeWork |
