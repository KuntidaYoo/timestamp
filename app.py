import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from shutil import copyfile
from tempfile import TemporaryDirectory

import streamlit as st

# --- CONFIG ---
LATE_COL_DAY = 6          # column G in day-by-day file (0-based index)
HEADER_ROW = 2            # green header row in template (dates + "สาย(นาที)")
ID_COL = 2                # column B in template (employee ID)
FIRST_DATE_COL = 4        # column D in template (first date column)


def load_day_file(path: str) -> pd.DataFrame:
    # read xls with xlrd, xlsx with openpyxl (pandas auto-detects)
    df = pd.read_excel(path, header=None)

    # keep only rows where first cell starts with ≥5 digits (employee rows)
    first_col = df[0].astype(str).str.strip()
    mask = first_col.str.match(r"^\d{5,}")
    df_emp = df[mask].copy()

    rename_map = {
        0: "emp_id",
        1: "name",
        4: "time_in",
        5: "time_out"
    }

    if 6 in df_emp.columns:
        rename_map[6] = "late"

    if 9 in df_emp.columns:  rename_map[9]  = "no_in"
    if 10 in df_emp.columns: rename_map[10] = "no_out"
    if 11 in df_emp.columns: rename_map[11] = "absent"
    if 12 in df_emp.columns: rename_map[12] = "sick"
    if 13 in df_emp.columns: rename_map[13] = "personal"
    if 14 in df_emp.columns: rename_map[14] = "vacation"
    if 15 in df_emp.columns: rename_map[15] = "ordination"

    df_emp = df_emp.rename(columns=rename_map)

    df_emp["emp_id"] = df_emp["emp_id"].astype(str).str.strip()
    df_emp["name"]   = df_emp["name"].astype(str).str.strip()

    cols = ["emp_id", "name", "time_in", "time_out"]
    for extra in ["late","no_in","no_out","absent","sick","personal","vacation","ordination"]:
        if extra in df_emp.columns:
            cols.append(extra)

    return df_emp[cols]



def get_template_date_key_from_filename(path: str) -> str | None:
    """
    From filename like '24.12.68.A.xls' build the same string
    used by template headers, e.g. '2568-12-24 00:00:00'.
    """
    base = os.path.basename(path)  # '24.12.68.A.xls'
    parts = base.split(".")
    if len(parts) < 3:
        return None
    try:
        d = int(parts[0])
        m = int(parts[1])
        yy = int(parts[2])          # 68 -> 2568 (B.E.)
        year_be = 2500 + yy
        key = f"{year_be:04d}-{m:02d}-{d:02d} 00:00:00"
        return key
    except ValueError:
        return None


def fill_template_from_days(template_path: str, day_files: list[str], output_path: str):
    """Copy template.xlsx to output_path and fill in times, late minutes, and reasons."""
    # delete old output file if it exists
    if os.path.exists(output_path):
        os.remove(output_path)

    # copy original template so we don't touch it
    copyfile(template_path, output_path)

    wb = load_workbook(output_path)
    ws = wb.active   # first sheet

    # 1) build map of date headers: HEADER_ROW, columns D..end
    date_col_map = {}      # header_value_as_str -> column index
    for c in range(FIRST_DATE_COL, ws.max_column + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if v is None:
            continue
        text = str(v).strip()
        if text:
            date_col_map[text] = c

    print("Date headers found in template:")
    for k, v in date_col_map.items():
        print(f"  {k} -> col {v}")

    if not date_col_map:
        raise RuntimeError("ไม่พบหัวคอลัมน์วันที่ในแถว header")

    # 2) find "สาย (นาที)" column
    late_col_idx = None
    for c in range(FIRST_DATE_COL, ws.max_column + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if v is None:
            continue
        text = str(v)
        if "สาย" in text or "late" in text:
            late_col_idx = c
            break

    print("Late column index in template:", late_col_idx)

    # 3) index employees by ID only: id -> row index
    index_by_id = {}
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        emp_id_val = ws.cell(row=r, column=ID_COL).value
        if emp_id_val is None:
            continue
        emp_id = str(emp_id_val).strip()
        if emp_id:
            index_by_id[emp_id] = r

    print(f"Number of employees in template (unique IDs): {len(index_by_id)}")

    total_matches = 0

    # 4) read every Day-by-Day file and write values
    for path in day_files:
        print(f"\nProcessing day file: {path}")
        df = load_day_file(path)

        # determine which date column this file belongs to
        date_key = get_template_date_key_from_filename(path)
        print(f"  Date key from filename: {date_key}")
        if date_key not in date_col_map:
            print("  -> This date key not found in template headers, skipping this file.")
            continue
        col_idx = date_col_map[date_key]
        print(f"  -> Will write into column index: {col_idx}")

        day_ids = sorted(df["emp_id"].unique())
        intersection = set(day_ids) & set(index_by_id.keys())
        print(f"  Employees in this file: {len(day_ids)}")
        print(f"  IDs that appear in BOTH day file and template: {len(intersection)}")

        has_late = "late" in df.columns

        for _, row in df.iterrows():
            emp_id = str(row["emp_id"]).strip()
            if not emp_id:
                continue

            row_idx = index_by_id.get(emp_id)
            if row_idx is None:
                continue  # ID not in template

            time_in = row["time_in"]
            cell = ws.cell(row=row_idx, column=col_idx)

            if pd.notna(time_in):
                # normal case: has time-in
                cell.value = str(time_in)
                total_matches += 1
            else:
                # NO time-in -> highlight + write reason(s) inside this date cell
                yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                cell.fill = yellow

                label_map = {
                    "no_in":      "ไม่ตอกเข้า",
                    "no_out":     "ไม่ตอกออก",
                    "absent":     "ขาดงาน",
                    "sick":       "ลาป่วย",
                    "personal":   "ลากิจ",
                    "vacation":   "พักร้อน",
                    "ordination": "บวชคลอด",
                }

                pieces = []
                for key, label in label_map.items():
                    if key not in row.index:
                        continue
                    val = row[key]
                    if pd.isna(val):
                        continue
                    try:
                        num = float(val)
                    except (TypeError, ValueError):
                        continue
                    if num == 0:
                        continue

                    if abs(num - round(num)) < 1e-6:
                        num_str = str(int(round(num)))
                    else:
                        num_str = str(num)

                    if num_str == "1":
                        pieces.append(label)
                    else:
                        pieces.append(f"{label}({num_str})")

                if pieces:
                    cell.value = " ".join(pieces)
                else:
                    cell.value = None

            # 4b) accumulate late minutes into "สาย (นาที)" column
            if has_late and late_col_idx is not None and "late" in row.index:
                late_val = row["late"]
                if pd.notna(late_val):
                    try:
                        add_val = float(late_val)
                    except (TypeError, ValueError):
                        add_val = 0.0

                    if add_val != 0.0:
                        existing = ws.cell(row=row_idx, column=late_col_idx).value
                        try:
                            base = float(existing) if existing is not None else 0.0
                        except (TypeError, ValueError):
                            base = 0.0
                        ws.cell(row=row_idx, column=late_col_idx).value = base + add_val

    wb.save(output_path)
    print("\nSaved filled template to:", output_path)
    print(f"Total cells filled (time_in writes): {total_matches}")


# ===================== STREAMLIT APP =====================

st.set_page_config(page_title="Time Day-by-Day Info", layout="wide")

st.title("สรุปการมาทำงาน")

# --- TEMPLATE DOWNLOAD SECTION ---
st.subheader("📄 Template ต้องใช้รูปแบบนี้เท่านั้น")
with open("template-public.xlsx", "rb") as f:
    st.download_button(
        label="📥 ดาวโหลด",
        data=f,
        file_name="template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")   # divider line

# --- instructions ---
st.markdown(
    """
    **ขั้นตอนการใช้งาน**

    1. อัปโหลดไฟล์ **template** (สรุปการทำงานประจำงวด) – ต้องเป็น `.xlsx`  
    2. อัปโหลดไฟล์ **Day-by-Day** หลายไฟล์ได้ (ต้องเป็น `.xlsx` เช่น `24.12.68.A.xlsx`, `24.12.68.G.xlsx`)  
    3. กด **Generate** เพื่อสร้างไฟล์สรุป `template_filled.xlsx` แล้วดาวน์โหลด
    """
)


template_file = st.file_uploader("อัปโหลด template.xlsx", type=["xlsx"])
day_files_uploaded = st.file_uploader(
    "อัปโหลดไฟล์ข้อมูลสแกนนิ้ว อย่างน้อย 1 ไฟล์ (.xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

generate = st.button("Generate filled template")

if generate:
    if template_file is None:
        st.error("กรุณาอัปโหลด template.xlsx ก่อน")
    elif not day_files_uploaded:
        st.error("กรุณาอัปโหลดไฟล์ Day-by-Day อย่างน้อย 1 ไฟล์")
    else:
        with TemporaryDirectory() as tmpdir:
            # save template
            template_path = os.path.join(tmpdir, "template.xlsx")
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())

            # save each day file to temp folder
            day_paths = []
            for uploaded in day_files_uploaded:
                fname = uploaded.name
                day_path = os.path.join(tmpdir, fname)
                with open(day_path, "wb") as f:
                    f.write(uploaded.getbuffer())
                day_paths.append(day_path)

            output_path = os.path.join(tmpdir, "template_filled.xlsx")

            try:
                fill_template_from_days(template_path, day_paths, output_path)

                with open(output_path, "rb") as f:
                    data = f.read()

                st.success("สร้างไฟล์ template_filled.xlsx เรียบร้อยแล้ว 🎉")
                st.download_button(
                    label="⬇️ Download template_filled.xlsx",
                    data=data,
                    file_name="template_filled.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดระหว่างการประมวลผล: {e}")
