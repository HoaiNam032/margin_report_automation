# 📊 Margin Report Automation

---

## 📌 Tổng quan

**Margin Report Automation** là project tự động hóa báo cáo **tình hình giao dịch ký quỹ**, được xây dựng theo hướng **SQL-driven reporting pipeline** kết hợp Python và Microsoft SQL Server Data Warehouse.

Project tập trung mạnh vào việc **khai thác, tổng hợp và xử lý dữ liệu ngay tại SQL Server**, sau đó Python đảm nhiệm phần orchestration, xử lý DataFrame và xuất báo cáo Excel.

Thay vì lấy toàn bộ dữ liệu thô về Python rồi mới xử lý, project ưu tiên:

```text
SQL Server
    ↓
SQL Query / Data Extraction
    ↓
DataFrame
    ↓
Business Logic
    ↓
Excel Report

```
## 🎯 Mục tiêu

Project được xây dựng để tự động hóa quy trình tạo báo cáo giao dịch ký quỹ:

- Kết nối trực tiếp đến Data Warehouse.
- Tự động xác định ngày giao dịch mới nhất.
- Truy vấn dữ liệu Margin.
- Truy vấn dữ liệu Used Room.
- Truy vấn số lượng chứng khoán niêm yết.
- Truy vấn giá đóng cửa.
- Xử lý Room đặc biệt.
- Kết hợp nhiều nguồn dữ liệu trong Data Warehouse.
- Tính toán dư nợ cho vay giao dịch ký quỹ.
- Tính các tỷ lệ kiểm soát.
- Phân biệt General Room và Special Room.
- Tạo báo cáo Excel theo format nghiệp vụ.
- Có thể chạy tự động thông qua Windows Batch / Task Scheduler.

## ⭐ SQL — Thành phần trọng tâm

Điểm nổi bật của project là **SQL Server được sử dụng như Data Processing Layer**.

Python không chỉ đơn giản đọc một file Excel rồi tính toán. Project có nhánh xử lý trực tiếp từ **Data Warehouse bằng SQL**, trong đó các bảng dữ liệu được truy vấn độc lập rồi đưa vào pipeline để xây dựng báo cáo.

Trong code, project sử dụng **SQLAlchemy** kết hợp `pyodbc` để kết nối tới **Microsoft SQL Server** thông qua:

```text
ODBC Driver 17 for SQL Server
        ↓
     pyodbc
        ↓
   SQLAlchemy
        ↓
SQL Server Data Warehouse
```

# 🗄️ Data Warehouse

Project khai thác dữ liệu từ nhiều bảng thuộc các **Data Warehouse** khác nhau.

## 📌 Các nguồn dữ liệu chính

```text
[DWH-CoSo].[dbo].[vpr0109]
[DWH-CoSo].[dbo].[230007]
[DWH-CoSo].[dbo].[Data_Hop_Infos]
[DWH-CoSo].[dbo].[vpr0108]

[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
```

# 🔗 CTE — Common Table Expression (ví dụ điển hình)

Project sử dụng **CTE (Common Table Expression)** để xử lý **mapping các Special Room** ngay tại SQL layer.

CTE giúp định nghĩa một bảng tạm logic ngay trong câu SQL, sau đó có thể sử dụng bảng này để `JOIN` với các bảng dữ liệu trong Data Warehouse.

### 🧩 Tạo Special Room Mapping

Ví dụ:

```sql
WITH special_map AS (
    SELECT *
    FROM (VALUES
        ('TV4', 'CL08_PHS', '0095'),
        ('GEX', 'CL07_PHS', '0074'),
        ('BAF', 'CL47_PHS', '0216'),
        ('DL1', 'CL59_PHS', '0287')
    ) AS v(
        ticker,
        room_code_09,
        room_code_08
    )
)
```
## 🔗 JOIN dữ liệu từ nhiều nguồn

Pipeline thực hiện **JOIN và mapping dữ liệu từ nhiều bảng thuộc các Data Warehouse khác nhau** để xây dựng một bộ dữ liệu thống nhất phục vụ tính toán báo cáo.

### 🗄️ Data Sources

```text
vpr0109
   │
   ├── Margin Ratio
   ├── Max Price
   └── Room Code
        │
        ├──────────────┐
        │              │
      230007         vpr0108
        │              │
        ├── Used GR    ├── Special Room
        └── Used SR    └── Special Price / Ratio

Data_Hop_Infos
        │
        └── Volume Listed

DuLieuGiaoDichNgay
        │
        └── Closing Price

```
# 🐍 Vai trò của Python

Python đóng vai trò **Orchestration + Business Processing + Reporting** trong toàn bộ pipeline.

Python không trực tiếp thay thế SQL Server trong việc truy vấn dữ liệu mà đảm nhiệm việc **điều phối pipeline, xử lý dữ liệu sau khi extraction và xây dựng báo cáo cuối cùng**.

### 🗄️ Database

Python chịu trách nhiệm kết nối và tương tác với SQL Server:

- Tạo SQLAlchemy Engine.
- Kết nối SQL Server.
- Thực thi SQL Query / Stored Procedure.
- Đọc kết quả SQL vào DataFrame.
- Quản lý luồng dữ liệu giữa Database và Python.

### ⚙️ Data Processing

Sau khi lấy dữ liệu từ Data Warehouse, Python thực hiện các bước xử lý:

- Merge dữ liệu từ nhiều nguồn.
- `GroupBy` và Data Aggregation.
- Chuẩn hóa dữ liệu.
- Xử lý Special Room.
- Mapping dữ liệu giữa các bảng.
- Tính toán các Business Rules phức tạp.
- Tính các chỉ tiêu phục vụ báo cáo.

### 📊 Reporting

Python đảm nhiệm quá trình tạo và format báo cáo:

- Tạo Workbook.
- Tạo Worksheet.
- Format Header.
- Format số liệu.
- Highlight Special Room.
- Tạo dòng tổng.
- Định dạng báo cáo theo Business Requirement.
- Xuất báo cáo Excel.

### 🔄 Python Pipeline

```text
SQL Server Data Warehouse
          ↓
     SQL / Stored Procedure
          ↓
      Data Extraction
          ↓
        Python
          │
          ├── Data Merge
          ├── Data Cleaning
          ├── Data Mapping
          ├── Business Logic
          └── Data Aggregation
          ↓
    Report Generation
          ↓
       Excel Report
```

