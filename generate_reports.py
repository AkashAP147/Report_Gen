import pandas as pd
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

def extract_subjects(excel_file):
    try:
        tables = pd.read_html(excel_file)
        if not tables: return []
        df = tables[0]
        if 'S No' not in df.columns and 'USN' not in df.columns:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        df = df.dropna(how='all')
        if 'Subject Code' in df.columns:
            return df['Subject Code'].dropna().unique().tolist()
        return []
    except Exception as e:
        print(f"Extraction error: {e}")
        return []

def assign_blocks(df, max_per_block=32, pref_left=None, pref_right=None):
    subject_students = {subj: group.to_dict('records') for subj, group in df.groupby('Subject Code')}
    arranged_students = []
    
    block_labels = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX']
    block_idx = 0
    
    while subject_students:
        block_no_str = f"Block-{block_labels[block_idx] if block_idx < len(block_labels) else block_idx + 1}"
        
        # Pick largest subject to start the left side, prioritizing preferred subjects
        def get_sort_key(k):
            boost = 0
            if pref_left and k == pref_left: boost = 1000000
            elif pref_right and k == pref_right: boost = 500000
            return len(subject_students[k]) + boost
            
        sorted_subjs = sorted(subject_students.keys(), key=get_sort_key, reverse=True)
        subj_left = sorted_subjs[0]
        
        # Pick second largest subject to start the right side
        subj_right = sorted_subjs[1] if len(sorted_subjs) > 1 else None
        
        single_subj_turn = 'left'
        
        # Fill row by row to ensure even depth on both sides
        for i in range(16):
            if not subject_students:
                break
                
            desk_odd = i * 2 + 1
            desk_even = i * 2 + 2
            
            # If only one subject remains globally, zig-zag it to balance the columns evenly
            if len(subject_students) == 1:
                subj = list(subject_students.keys())[0]
                if single_subj_turn == 'left':
                    student = subject_students[subj].pop(0)
                    student['Block No'] = block_no_str
                    student['Desk No'] = desk_odd
                    arranged_students.append(student)
                    single_subj_turn = 'right'
                else:
                    student = subject_students[subj].pop(0)
                    student['Block No'] = block_no_str
                    student['Desk No'] = desk_even
                    arranged_students.append(student)
                    single_subj_turn = 'left'
                    
                if not subject_students[subj]:
                    del subject_students[subj]
                continue
            
            # Fill Odd Desk (Left Column)
            if subj_left not in subject_students:
                available = [s for s in subject_students.keys() if s != subj_right]
                if available:
                    available.sort(key=lambda k: len(subject_students[k]), reverse=True)
                    subj_left = available[0]
                else:
                    subj_left = None
                    
            if subj_left is not None:
                student = subject_students[subj_left].pop(0)
                student['Block No'] = block_no_str
                student['Desk No'] = desk_odd
                arranged_students.append(student)
                
                if not subject_students[subj_left]:
                    del subject_students[subj_left]
                    
            # Fill Even Desk (Right Column)
            if subj_right not in subject_students:
                available = [s for s in subject_students.keys() if s != subj_left]
                if available:
                    available.sort(key=lambda k: len(subject_students[k]), reverse=True)
                    subj_right = available[0]
                else:
                    subj_right = None
                    
            if subj_right is not None:
                student = subject_students[subj_right].pop(0)
                student['Block No'] = block_no_str
                student['Desk No'] = desk_even
                arranged_students.append(student)
                
                if not subject_students[subj_right]:
                    del subject_students[subj_right]
                    
        block_idx += 1
        
    return pd.DataFrame(arranged_students)

def parse_excel(filepath, pref_left=None, pref_right=None):
    # Read the html file
    dfs = pd.read_html(filepath)
    
    # The first table is likely the header
    # But usually pd.read_html just gets the big table.
    # Let's read the raw html to get header info
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        html_content = f.read()
        
    import re
    college_code = 'VS'
    date = '2026-09-03'
    session = '09:30:00'
    
    if 'CollegeCode -' in html_content:
        m = re.search(r'CollegeCode - ([\w]+)', html_content)
        if m: college_code = m.group(1).strip()
    if 'Date -' in html_content:
        m = re.search(r'Date - ([\d\-]+)', html_content)
        if m: date = m.group(1).strip()
    if 'Session -' in html_content:
        m = re.search(r'Session - ([\d:]+)', html_content)
        if m: session = m.group(1).strip()

    # The actual data table is usually the largest one or the only one parsed
    df_data = dfs[0]
    # Clean up columns if it's the main table
    if 'S No' in df_data.columns or 'USN' in df_data.columns:
        pass # Already has headers
    else:
        # If headers are in row 0
        df_data.columns = df_data.iloc[0]
        df_data = df_data[1:].reset_index(drop=True)
        
    df_data = assign_blocks(df_data, pref_left=pref_left, pref_right=pref_right)
        
    return {
        'college_code': college_code,
        'date': date,
        'session': session,
        'data': df_data
    }

def create_block_list(data_dict, output_dir):
    df = data_dict['data']
    date = data_dict['date']
    session = data_dict['session']
    
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    doc = docx.Document()
    
    for block_no, group in df.groupby('Block No', sort=False):
        # Header formatting
        p_top = doc.add_paragraph()
        p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_top = p_top.add_run("THEORY EXAMINATIONS JUNE/JULY-2026")
        run_top.bold = True
        run_top.underline = True
        
        subjects = group['Subject Code'].unique()
        subj_str = ", ".join(map(str, subjects))
        
        doc.add_paragraph(f"Subject:\nSubject Code: {subj_str} \t\t\t\t Semester:")
        doc.add_paragraph(f"Date: {date} \t\t\t\t\t Time: {session}")
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{block_no.upper().replace('-', ' - ')}")
        run.bold = True
        run.underline = True
        
        table = doc.add_table(rows=17, cols=4)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        for idx, text in enumerate(['USN', 'DESK NO.', 'USN', 'DESK NO.']):
            hdr_cells[idx].text = text
            hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            hdr_cells[idx].paragraphs[0].runs[0].bold = True
            
        # Pre-fill desk numbers
        for i in range(1, 17):
            row = table.rows[i]
            # Left desk (Odd)
            desk_left = i * 2 - 1
            row.cells[1].text = str(desk_left)
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[1].paragraphs[0].runs[0].bold = True
            
            # Right desk (Even)
            desk_right = i * 2
            row.cells[3].text = str(desk_right)
            row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[3].paragraphs[0].runs[0].bold = True
            
        # Map USNs by desk
        desk_to_usn = {row['Desk No']: row['USN'] for _, row in group.iterrows()}
        
        for i in range(1, 17):
            row = table.rows[i]
            desk_left = i * 2 - 1
            desk_right = i * 2
            
            if desk_left in desk_to_usn:
                row.cells[0].text = str(desk_to_usn[desk_left])
                row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if desk_right in desk_to_usn:
                row.cells[2].text = str(desk_to_usn[desk_right])
                row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        doc.add_page_break()
        
    output_path = os.path.join(output_dir, 'Block_List_Generated.docx')
    doc.save(output_path)
    print(f"Generated: {output_path}")

def create_dispatch_format_i(data_dict, output_dir):
    df = data_dict['data']
    date = data_dict['date']
    session = data_dict['session']
    
    # Format Date
    date_parts = date.split('-')
    if len(date_parts) == 3:
        formatted_date = f"{date_parts[2]}.{date_parts[1]}.{date_parts[0]}"
    else:
        formatted_date = date
        
    # Format Time
    if "09:" in session or "10:" in session or "11:" in session or "AM" in session.upper():
        formatted_time = "09.30 AM TO 12.30 PM"
    else:
        formatted_time = "02.00 PM TO 05.00 PM"
    
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
    import re
    
    doc = docx.Document()
    
    for subject_code, group in df.groupby('Subject Code'):
        # Set Landscape
        section = doc.sections[-1]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        # We need a table to place logo on the left and text in the center
        header_table = doc.add_table(rows=1, cols=2)
        header_table.columns[0].width = Inches(1.5)
        header_table.columns[1].width = Inches(7.5)
        
        # Add Logo
        cell_left = header_table.cell(0, 0)
        p_logo = cell_left.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_path = r"d:\COMMERCIAL\DOC_COV\vtu_logo.png"
        if os.path.exists(logo_path):
            run_logo = p_logo.add_run()
            run_logo.add_picture(logo_path, width=Inches(1.0))
            
        # Add Header Text
        cell_right = header_table.cell(0, 1)
        p_text = cell_right.paragraphs[0]
        p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p_text.add_run("Book Dispatch Format - I\n")
        run1.bold = True
        run2 = p_text.add_run("VISVESVARAYA TECHNOLOGICAL UNIVERSITY, BELGAUM\nBELGAUM REGION\nB.E- V Semester Examination JUNE/JULY-2026")
        run2.bold = True
        
        doc.add_paragraph() # Spacing
        
        # Extract branch from USN
        usn = str(group['USN'].iloc[0]).upper()
        # Typical format 2GI21EC001 -> matches 'EC'
        m = re.search(r'\d[A-Z]{2}\d{2}([A-Z]{2,3})', usn)
        if m:
            enclosure = m.group(1)
        else:
            m2 = re.search(r'[A-Za-z]+', str(subject_code))
            enclosure = m2.group(0) if m2 else "ALL"
            
        doc.add_paragraph("Exam Centre: - V.S.M's S.R.K Institute of Technology, Nipani")
        doc.add_paragraph(f"Enclosures: - {enclosure}\nAnswer Script Details:-")
        
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Date of Exam'
        hdr_cells[1].text = 'Exam Timing'
        hdr_cells[2].text = 'Subject Code'
        hdr_cells[3].text = 'No. of Papers'
        
        for cell in hdr_cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].bold = True
        
        row_cells = table.add_row().cells
        row_cells[0].text = formatted_date
        row_cells[1].text = formatted_time
        row_cells[2].text = str(subject_code)
        row_cells[3].text = f"{enclosure}-{len(group):02d}"
        
        for cell in row_cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].bold = True
        
        doc.add_paragraph("\n2. Form -A   \u2714    \t\t3. Form- B   \u2714  \t\t4. Question Paper   \u2714\n")
        
        doc.add_paragraph("To,\n       Dr. Sattagouda M Patil \n       Chief Co-ordinator \n       VTU, Digitization Centre  \"Jnana Sangama\"\n       VTU, Belagavi : 590018\n\nFrom,\n       Chief Superintendent\n")
        
        doc.add_page_break()
        
    output_path = os.path.join(output_dir, 'Dispatch_Format_I_Generated.docx')
    doc.save(output_path)
    print(f"Generated: {output_path}")

def create_dispatch_format_ii(data_dict, output_dir):
    df = data_dict['data']
    date = data_dict['date']
    session = data_dict['session']
    
    session_str = "MOR" if ("09:" in session or "10:" in session or "11:" in session or "AM" in session.upper()) else "AFT"
    date_exam_str = f"{date}\n{session_str}"
    
    doc = docx.Document()
    
    # Headers
    p = doc.add_paragraph("Answer Book Dispatch Format - II\n")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Visvesvaraya Technological University, Belagavi\nBELAGAVI REGION\nANSWER BOOK BUNDLES ACKNOWLEDGEMENT\nJUNE/JULY-2026 Examinations\n")
    run.bold = True
    
    p2 = doc.add_paragraph(f"                       Date: {date}\n")
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_paragraph("Exam Centre: VSM INSTITUTE OF TECHNOLOGY, NIPANI.\n")
    doc.add_paragraph("Valuation Centre: Prof. Sattagouda M Patil ,Chief Co-ordinator, VTU, Digitization Centre \"Jnana Sangama\" VTU, Belagavi-18\n")
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Sl. No.'
    hdr_cells[1].text = 'Date of Exam'
    hdr_cells[2].text = 'Subject Code'
    hdr_cells[3].text = 'Subject Title'
    hdr_cells[4].text = 'No. of Ans. Scripts'
    hdr_cells[5].text = 'No. of Bundles'
    
    sl_no = 1
    total_scripts = 0
    total_bundles = 0
    
    import re
    
    for subject_code, group in df.groupby('Subject Code'):
        subject_name = group['Subject Name'].iloc[0]
        scripts_count = len(group)
        bundles_count = (scripts_count // 20) + 1 # Assuming 20 scripts per bundle
        
        m = re.search(r'[A-Za-z]+', str(subject_code))
        enclosure = m.group(0) if m else "ALL"
        
        row_cells = table.add_row().cells
        row_cells[0].text = f"{sl_no:02d}"
        row_cells[1].text = date_exam_str
        row_cells[2].text = str(subject_code)
        row_cells[3].text = str(subject_name)
        row_cells[4].text = f"{enclosure}-{scripts_count:02d}"
        row_cells[5].text = f"{bundles_count:02d}"
        
        total_scripts += scripts_count
        total_bundles += bundles_count
        sl_no += 1
        
    row_cells = table.add_row().cells
    row_cells[0].text = 'TOTAL'
    row_cells[5].text = f"{total_bundles:02d}"
    
    # Merge Date of Exam cells if there are multiple subjects
    if len(table.rows) > 2:
        start_cell = table.cell(1, 1)
        end_cell = table.cell(len(table.rows) - 2, 1)
        start_cell.merge(end_cell)
        start_cell.text = date_exam_str
        start_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # Set alignment for cells
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Helper to convert numbers to words
    def num2words(n):
        words = {1: 'ONE', 2: 'TWO', 3: 'THREE', 4: 'FOUR', 5: 'FIVE', 6: 'SIX', 7: 'SEVEN', 8: 'EIGHT', 9: 'NINE', 10: 'TEN',
                 11: 'ELEVEN', 12: 'TWELVE', 13: 'THIRTEEN', 14: 'FOURTEEN', 15: 'FIFTEEN', 16: 'SIXTEEN', 17: 'SEVENTEEN',
                 18: 'EIGHTEEN', 19: 'NINETEEN', 20: 'TWENTY', 30: 'THIRTY', 40: 'FORTY', 50: 'FIFTY', 60: 'SIXTY', 70: 'SEVENTY', 80: 'EIGHTY', 90: 'NINETY'}
        if n in words: return words[n]
        if n < 100:
            tens, units = divmod(n, 10)
            return words[tens*10] + (" " + words[units] if units else "")
        return str(n)
        
    doc.add_paragraph(f"\nReceived ({num2words(total_bundles)}) Sealed Answer book bundles.\n\n\n")
    
    doc.add_paragraph("Chief Superintendent \t\t\t\t\t\t Member, Collection Team\n(Signature with date & Seal) \t\t\t\t\t (Signature with date)")
    
    output_path = os.path.join(output_dir, 'Dispatch_Format_II_Generated.docx')
    doc.save(output_path)
    print(f"Generated: {output_path}")

def create_consolidated_list(data_dict, output_dir):
    df = data_dict['data']
    date = data_dict['date']
    session = data_dict['session']
    
    from docx.shared import Cm, Inches, Pt
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    
    doc = docx.Document()
    
    # Set Narrow Margins to fit the table better
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    # Add Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("VSM INSTITUTE OF TECHNOLOGY, NIPANI\nVTU THEORY EXAMINATION JUNE/JULY-2026\nConsolidated block wise information")
    run.bold = True
    
    # Add Time and Date
    total_candidates = len(df)
    p2 = doc.add_paragraph(f"TIME :- {session} \t\t\t\t DATE: {date} \t\t\t\t TOTAL: {total_candidates:02d}")
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.allow_autofit = False
    
    hdr_cells = table.rows[0].cells
    headers = ['Block No', 'Subject Code', 'Subject Name', 'USN Numbers', 'Total No. Of Candidates', 'Invigilator\nSign']
    for idx, text in enumerate(headers):
        hdr_cells[idx].text = text
        for p_hdr in hdr_cells[idx].paragraphs:
            p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run_hdr in p_hdr.runs:
                run_hdr.bold = True
                run_hdr.font.size = Pt(10)
    
    # Adjusted widths to fit inside narrow margins (approx 19 cm total width available)
    widths = [Cm(1.5), Cm(2.0), Cm(2.8), Cm(9.5), Cm(1.7), Cm(1.5)]
    
    for block_no, block_group in df.groupby('Block No', sort=False):
        start_row_idx = len(table.rows)
        
        for subject_code, group in block_group.groupby('Subject Code'):
            subject_name = group['Subject Name'].iloc[0]
            usns = group['USN'].tolist()
            
            # Group into chunks of 4 and join with newlines (NO SPACES AFTER COMMAS)
            usns_chunks = [usns[i:i + 4] for i in range(0, len(usns), 4)]
            usns_str = "\n".join([",".join(map(str, chunk)) for chunk in usns_chunks])
            
            row_cells = table.add_row().cells
            row_cells[0].text = block_no
            row_cells[0].paragraphs[0].runs[0].bold = True
            
            row_cells[1].text = str(subject_code)
            row_cells[1].paragraphs[0].runs[0].bold = True
            
            row_cells[2].text = str(subject_name)
            row_cells[2].paragraphs[0].runs[0].bold = True
            
            row_cells[3].text = usns_str
            row_cells[4].text = f"{len(usns)}"
            row_cells[5].text = ''
            
            # Set font size for all cells in this row to 10pt
            for cell in row_cells:
                for p_cell in cell.paragraphs:
                    for run_cell in p_cell.runs:
                        run_cell.font.size = Pt(10)
            
        # Merge block_no and invigilator sign cells if multiple subjects
        end_row_idx = len(table.rows) - 1
        if end_row_idx > start_row_idx:
            # Merge Block No (col 0)
            table.cell(start_row_idx, 0).merge(table.cell(end_row_idx, 0))
            table.cell(start_row_idx, 0).text = block_no
            table.cell(start_row_idx, 0).paragraphs[0].runs[0].bold = True
            
            # Merge Invigilator Sign (col 5)
            table.cell(start_row_idx, 5).merge(table.cell(end_row_idx, 5))
            
    # Apply widths, alignments, and vertical centering
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
            row.cells[idx].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    doc.add_paragraph("\n\n\nDeputy Chief Supdt. \t\t\t\t\t\t\t Chief Supdt.")
    
    output_path = os.path.join(output_dir, 'Consolidated_List_Generated.docx')
    doc.save(output_path)
    print(f"Generated: {output_path}")

def run_generation(excel_file, output_dir, pref_left=None, pref_right=None):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    print(f"Reading {excel_file}...")
    try:
        data = parse_excel(excel_file, pref_left, pref_right)
        print("Data loaded successfully. Generating reports...")
        
        create_block_list(data, output_dir)
        create_dispatch_format_i(data, output_dir)
        create_dispatch_format_ii(data, output_dir)
        create_consolidated_list(data, output_dir)
        
        print(f"All reports generated successfully in {output_dir}")
        return True
    except Exception as e:
        print(f"Error occurred: {e}")
        return False

if __name__ == '__main__':
    base_dir = r"d:\COMMERCIAL\DOC_COV"
    excel_file = os.path.join(base_dir, "VS_2026-09-03_09_30_00_Session.xls")
    output_dir = os.path.join(base_dir, "Generated_Reports")
    run_generation(excel_file, output_dir)
