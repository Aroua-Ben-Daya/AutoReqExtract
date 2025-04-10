import pdfplumber
import pandas as pd
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment

# ------------------------------------------
# File paths
# Uncomment the line you need
#pdf_path = "D:/UNIVERSITE D'AIX MARSEILLE/Day 1 03_03_2025/X2R1/X2R1-T5.3-D-SIE-102-20_-_D5.1_-_Moving_Block_System_Requirements.pdf"
#pdf_path = "D:/UNIVERSITE D'AIX MARSEILLE/Day 1 03_03_2025/X2R3/X2R3-T4_3-D-SMD-008-19_-_D4.2Part3-SystemSpecification.pdf"
#pdf_path = "D:/UNIVERSITE D'AIX MARSEILLE/Day 1 03_03_2025/X2R5/X2R5-T4_2-D-SMD-003-23_-_D41Part3SystemSpecification.pdf"
pdf_path = "D:/UNIVERSITE D'AIX MARSEILLE/Day 1 03_03_2025/X2R5/X2R5-T4_2-D-SMD-003-23_-_D41Part3SystemSpecification.pdf"
# ------------------------------------------

pdf_filename = os.path.basename(pdf_path).replace(".pdf", "")
output_excel = f"C:/Users/aroua/Desktop/ReqEvolutionTracker/{pdf_filename}_Result.xlsx"

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Extracted Requirements"

# Headers
required_columns = ["Topic", "Requirement ID", "Description", "Traceability"]
for col_idx, header in enumerate(required_columns):
    ws.cell(row=1, column=col_idx + 1, value=header)

# Regex patterns
req_pattern = re.compile(r"(REQ-[A-Za-z0-9]+-\d+|\bREQ-[A-Za-z0-9]+)\s*(\[[^\]]+\])?")
traceability_pattern = re.compile(r"\[(X2R\d+ D\d+\.\d+: REQ-[A-Za-z0-9-]+)\]")
footer_pattern = re.compile(r"(GA\s*\d+\s*)?Page\s+\d+\s+of\s+\d+", re.IGNORECASE)

def extract_description(lines, start_idx):
    """Extracts the description of a requirement from lines."""
    description = []
    for i in range(start_idx, len(lines)):
        if "Rationale:" in lines[i] or "Guidance:" in lines[i]:
            break
        clean_line = footer_pattern.sub("", lines[i]).strip()
        if clean_line:
            description.append(clean_line)
    return "\n".join(description).strip()

def extract_requirements(pdf_path):
    """Parses a PDF and extracts all requirements."""
    requirements = []
    current_topic = "Unknown"
    last_traceability = "[Not Provided]"

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(layout=True)
            if not text:
                continue

            lines = text.split("\n")
            for idx, line in enumerate(lines):
                line = footer_pattern.sub("", line).strip()

                # Detect topic
                normal_topic = re.match(r"^\s*(\d+\.\d+)\s+([A-Za-z][A-Za-z0-9 \-/]+)$", line)
                compact_topic = re.match(r"^\s*(\d+\.\d+)([A-Z][A-Za-z0-9]+)\s*(.*)$", line)

                if normal_topic:
                    section_number = normal_topic.group(1)
                    topic_name = normal_topic.group(2).strip()
                    current_topic = f"{section_number} - {topic_name}"
                    print(f"✅ Detected Topic (normal): {current_topic}")

                elif compact_topic:
                    section_number = compact_topic.group(1)
                    section_code = compact_topic.group(2)
                    rest = compact_topic.group(3).strip()
                    current_topic = f"{section_number} - {section_code} {rest}".strip()
                    print(f"✅ Detected Topic (compact): {current_topic}")

                # Detect traceability
                traceability_match = traceability_pattern.search(line)
                if traceability_match:
                    last_traceability = traceability_match.group(1).strip()
                    print(f"🔵 Found Traceability: {last_traceability}")
                elif "[New]" in line:
                    last_traceability = "New"
                    print("🔵 Found Traceability: New")

                # Detect requirement ID
                req_match = req_pattern.search(line)
                if req_match:
                    req_id = req_match.group(1)
                    traceability = req_match.group(2).strip("[]") if req_match.group(2) else last_traceability
                    description = extract_description(lines, idx + 1)
                    print(f"📌 Storing: {req_id} | {current_topic} | {traceability}")
                    requirements.append((current_topic, req_id, description, traceability))
    return requirements

# Extract and clean
data = extract_requirements(pdf_path)
# Convert to DataFrame
extracted_df = pd.DataFrame(data, columns=required_columns)
# Supprimer les doublons
extracted_df.drop_duplicates(subset=["Requirement ID"], keep="first", inplace=True)
# Write to Excel
for row_idx, row in extracted_df.iterrows():
    for col_idx, col_name in enumerate(required_columns):
        ws.cell(row=row_idx + 2, column=col_idx + 1, value=row[col_name])

# Formatting
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(required_columns)):
    for cell in row:
        cell.alignment = Alignment(wrap_text=True)

# Row height (description)
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3):
    for cell in row:
        ws.row_dimensions[cell.row].height = None

# Column width
for col in ws.columns:
    max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = min(max_length + 5, 50)

# Save
wb.save(output_excel)
print(f"✅ Extraction completed! {len(extracted_df)} requirements saved to: {output_excel}")
