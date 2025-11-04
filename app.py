import streamlit as st
import os
import win32com.client
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from copy import deepcopy

# --- Stage 1: Convert PPT to PPTX ---
def convert_ppt_to_pptx(input_folder):
    ppt_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.ppt')]
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    for ppt_file in ppt_files:
        full_path = os.path.join(input_folder, ppt_file)
        pptx_path = os.path.join(input_folder, ppt_file[:-4] + '.pptx')
        st.info(f"Converting {ppt_file} to {pptx_path}...")
        presentation = powerpoint.Presentations.Open(full_path, WithWindow=False)
        presentation.SaveAs(pptx_path, FileFormat=24)  # 24 = pptx
        presentation.Close()
    powerpoint.Quit()
    st.success("All .ppt files converted to .pptx.")

# --- Stage 2: Extract revision data to Excel ---
REV_HEADERS = [
    "RELEASE NUMBER", "REV LTR",
    "REVISION DESCRIPTION", "BY", "DATE", "APPD"
]

def is_revision_table(table):
    header = [cell.text.strip().upper() for cell in table.rows[0].cells]
    return header == REV_HEADERS

def find_balloons_and_texts_recursive(shapes):
    balloon_shapes = []
    text_shapes = []
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            try:
                if sh.auto_shape_type in [9, 40, 56, 57]:
                    balloon_shapes.append(sh)
            except:
                pass
        if sh.has_text_frame and sh.text.strip():
            text_shapes.append(sh)
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            bs, ts = find_balloons_and_texts_recursive(sh.shapes)
            balloon_shapes.extend(bs)
            text_shapes.extend(ts)
    return balloon_shapes, text_shapes

def get_balloon_letters_flexible(slide):
    balloon_shapes, text_shapes = find_balloons_and_texts_recursive(slide.shapes)
    balloon_letters = []
    for balloon in balloon_shapes:
        bx, by, bw, bh = balloon.left, balloon.top, balloon.width, balloon.height
        b_center = (bx + bw / 2, by + bh / 2)
        nearest_letter, nearest_dist = "", float("inf")
        for txt in text_shapes:
            tx, ty, tw, th = txt.left, txt.top, txt.width, txt.height
            t_center = (tx + tw / 2, ty + th / 2)
            dist = ((b_center[0] - t_center[0]) ** 2 + (b_center[1] - t_center[1]) ** 2) ** 0.5
            text = txt.text.strip()
            if len(text) == 1 and dist < min(bw, bh):
                if dist < nearest_dist:
                    nearest_letter = text
                    nearest_dist = dist
        balloon_letters.append(nearest_letter)
    return balloon_letters

def extract_revision_data_multisheet(input_folder, excel_path):
    per_drawing = {}
    for ppt_file in sorted(os.listdir(input_folder)):
        if ppt_file.lower().endswith(".pptx"):
            ppt_path = os.path.join(input_folder, ppt_file)
            prs = Presentation(ppt_path)
            drawing_name = os.path.splitext(ppt_file)[0]
            revision_rows = []
            balloon_letters = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_table and is_revision_table(shape.table):
                        table = shape.table
                        for i in range(1, len(table.rows)):
                            row_data = [cell.text.strip() for cell in table.rows[i].cells]
                            revision_rows.append(row_data)
                        balloon_letters = get_balloon_letters_flexible(slide)
                        while len(balloon_letters) < len(revision_rows):
                            balloon_letters.append("")
                        balloon_letters = balloon_letters[:len(revision_rows)]
            sheet_rows = []
            for row, balloon in zip(revision_rows, balloon_letters):
                sheet_rows.append(row + [balloon])
            if sheet_rows:
                columns = REV_HEADERS + ["Balloon Text"]
                per_drawing[drawing_name] = pd.DataFrame(sheet_rows, columns=columns)
    with pd.ExcelWriter(excel_path) as writer:
        for name, df in per_drawing.items():
            sheet_name = str(name)[:31]
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    st.success(f"Extraction complete. Data saved in: {excel_path}")

# --- Stage 3: Edit PPTs from updated Excel ---
TABLE_FONT = "Arial Narrow"
TABLE_SIZE = Pt(7)
BALLOON_FONT = "Arial"
BALLOON_SIZE = Pt(11)
REVISION_HEADERS = [
    "RELEASE NUMBER", "REV LTR", "REVISION DESCRIPTION", "BY", "DATE", "APPD"
]

def is_revision_table_edit(headers):
    normalized = [h.upper().replace(" ", "") for h in headers]
    required = [h.upper().replace(" ", "") for h in REVISION_HEADERS]
    return all(h in normalized for h in required)

def clear_table_rows(table):
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)

def add_revision_rows(table, revision_data):
    for rev in revision_data:
        header_xml = table.rows[0]._tr
        new_xml = deepcopy(header_xml)
        table._tbl.append(new_xml)
        new_row = table.rows[len(table.rows) - 1]
        for i, val in enumerate(rev):
            if i < len(new_row.cells):
                cell = new_row.cells[i]
                cell.text = str(val) if val is not None else ""
                para = cell.text_frame.paragraphs[0]
                para.font.name = TABLE_FONT
                para.font.size = TABLE_SIZE

def update_table_and_balloon_for_all(multisheet_excel, ppt_folder, output_folder):
    xl = pd.ExcelFile(multisheet_excel)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for ppt_file in sorted(os.listdir(ppt_folder)):
        if ppt_file.lower().endswith(".pptx"):
            ppt_name = os.path.splitext(ppt_file)[0]
            try:
                if ppt_name not in xl.sheet_names:
                    st.warning(f"Sheet for {ppt_file} not found, skipping.")
                    continue
                df = xl.parse(ppt_name).dropna(how='all')
                if df.empty:
                    st.warning(f"Sheet {ppt_name} is empty, skipping.")
                    continue
                revision_data = df[REVISION_HEADERS].values.tolist()
                balloon_values = df["Balloon Text"].dropna()
                balloon_letter = str(balloon_values.iloc[-1]) if not balloon_values.empty else ""
                pptx_path = os.path.join(ppt_folder, ppt_file)
                output_path = os.path.join(output_folder, ppt_file)
                prs = Presentation(pptx_path)
                revision_done = False
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_table:
                            headers = [cell.text.strip().upper() for cell in shape.table.rows[0].cells]
                            if is_revision_table_edit(headers) and not revision_done:
                                clear_table_rows(shape.table)
                                add_revision_rows(shape.table, revision_data)
                                revision_done = True
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame and shape.text.strip():
                            txt = shape.text.strip()
                            if (len(txt) == 1 and txt.isalpha()) or (len(txt) == 2 and txt[0].isalpha() and txt[1] == '.'):
                                if balloon_letter:
                                    shape.text = str(balloon_letter)
                                    para = shape.text_frame.paragraphs[0]
                                    para.font.name = BALLOON_FONT
                                    para.font.size = BALLOON_SIZE
                prs.save(output_path)
                st.success(f"Updated: {ppt_file}")
            except Exception as e:
                st.error(f"Error updating {ppt_file}: {e}")

# --- Stage 4: Add bullet point to PPTX ---
def add_bullet_point_to_pptx(input_folder, output_folder, new_text_line):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(input_folder):
        if not filename.endswith(".pptx"):
            continue
        file_path = os.path.join(input_folder, filename)
        prs = Presentation(file_path)
        modified = False
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                text_frame = shape.text_frame
                paragraphs = [p for p in text_frame.paragraphs if p.text.strip() != ""]
                if len(paragraphs) < 2:
                    continue
                last_paragraph = paragraphs[-1]
                font_name = None
                font_size = None
                if last_paragraph.runs:
                    font_name = last_paragraph.runs[0].font.name
                    font_size = last_paragraph.runs[0].font.size
                blank_para = text_frame.add_paragraph()
                blank_para.text = " "
                new_para = text_frame.add_paragraph()
                new_para.text = f"{len(paragraphs) + 1}. {new_text_line}"
                new_para.level = last_paragraph.level
                if new_para.runs:
                    run = new_para.runs[0]
                    run.font.name = font_name if font_name else "Arial"
                    run.font.size = font_size
                modified = True
        if modified:
            output_path = os.path.join(output_folder, f"updated_{filename}")
            prs.save(output_path)
            st.success(f"Updated file saved: {output_path}")
        else:
            st.warning(f"No bullet text found in: {filename}")

# ================ Streamlit UI ================

st.title("PowerPoint Engineering Drawings Automation")

st.sidebar.header("Select Stage to Run")
stage = st.sidebar.radio(
    "Automation Stages",
    ("1: Convert PPT to PPTX",
     "2: Extract Revision Data to Excel",
     "3: Edit PPTs from Excel",
     "4: Add Bullet Point to PPTX")
)

if stage == "1: Convert PPT to PPTX":
    st.header("Stage 1: Convert PPT to PPTX")
    input_folder = st.text_input("Input Folder Path (containing .ppt files)")
    if st.button("Run Conversion"):
        if input_folder and os.path.exists(input_folder):
            convert_ppt_to_pptx(input_folder)
        else:
            st.error("Please provide a valid input folder path.")

elif stage == "2: Extract Revision Data to Excel":
    st.header("Stage 2: Extract Revision Data to Excel")
    input_folder = st.text_input("Input Folder Path (containing .pptx files)")
    excel_path = st.text_input("Save Excel File To (full path with .xlsx)")
    if st.button("Run Extraction"):
        if input_folder and os.path.exists(input_folder) and excel_path:
            extract_revision_data_multisheet(input_folder, excel_path)
        else:
            st.error("Please provide valid folder and excel save path.")

elif stage == "3: Edit PPTs from Excel":
    st.header("Stage 3: Edit PPTs from Updated Excel")
    excel_path = st.text_input("Excel File Path (with revision data)")
    ppt_folder = st.text_input("PPTX Folder Path (to edit)")
    output_folder = st.text_input("Output Folder Path (to save edited PPTX files)")
    if st.button("Run Editing"):
        if all([excel_path, ppt_folder, output_folder]) and os.path.exists(excel_path) and os.path.exists(ppt_folder):
            update_table_and_balloon_for_all(excel_path, ppt_folder, output_folder)
        else:
            st.error("Please provide valid Excel, input and output folder paths.")

elif stage == "4: Add Bullet Point to PPTX":
    st.header("Stage 4: Add Bullet Point to PPTX")
    input_folder = st.text_input("Input PPTX Folder Path")
    output_folder = st.text_input("Output Folder Path")
    new_text = st.text_area("New Bullet Text Line")
    if st.button("Add Bullet Point"):
        if input_folder and output_folder and new_text:
            add_bullet_point_to_pptx(input_folder, output_folder, new_text)
        else:
            st.error("Please provide valid input/output folders and bullet text.")
