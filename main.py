import pandas as pd
from docxtpl import DocxTemplate
import os

# Đọc file Excel (test với 2 hàng) - giữ CMND như string
df = pd.read_excel("text.xls", nrows=2, dtype={"CMND": str, "Số CMND": str})

# Tạo thư mục output
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

for index, row in df.iterrows():
    doc = DocxTemplate("template.docx")
    context = {
        "Ten_KH": row.get("Ten_KH", ""),
        "CMND": str(row.get("CMND", "")),
        "Ngay_Cap": row.get("Ngay_Cap_CMND", ""),
        "Noi_Cap": row.get("Noi_Cap_CMND", ""),
        "To_Truong": row.get("Ten_To", ""),
        "Ma_KH": row.get("Ma_KH", ""),
    }

    doc.render(context)
    
    # Tạo tên file an toàn (giữ ký tự tiếng Việt)
    customer_name = str(row.get("Ten_KH", "unknown"))
    # Loại bỏ chỉ ký tự không hợp lệ trên Windows
    invalid_chars = '<>:"|?*\\'
    safe_name = "".join(c for c in customer_name if c not in invalid_chars and ord(c) >= 32).strip()
    if not safe_name:
        safe_name = f"Document_{index}"
    
    filename = f"{output_dir}/{safe_name}.docx"
    doc.save(filename)

print("✅ Done! Đã tạo file Word hàng loạt từ CSV.")