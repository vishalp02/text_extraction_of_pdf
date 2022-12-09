import fitz
import pdfplumber
import pandas as pd
import re


pdf_path = 'OPC 30050 - UA Companion Specification for Packml 1.01.pdf'
pdf = fitz.open(pdf_path)

toc = pdf.get_toc()

#finding from-to page number of the section
for j in range(len(toc)):
    
    if re.search('^6.3 ', toc[j][1]):
        page_start = toc[j][2]
        
    if re.search('^6.4 ', toc[j][1]):
        page_end = toc[j][2]
         

def not_within_bboxes(obj):

    def obj_in_bbox(_bbox):
        v_mid = (obj['top'] + obj['bottom']) / 2
        h_mid = (obj['x0'] + obj['x1']) / 2
        x0, top, x1, bottom = _bbox
        return (h_mid >= x0) and (h_mid < x1) and (v_mid >= top) and (v_mid < bottom)

    return not any(obj_in_bbox(__bbox) for __bbox in bboxes)


pdf_data = []

with pdfplumber.open(pdf_path) as pdf:
    for i in range(page_start-1, page_end-1):
        page = pdf.pages[i]

        bboxes = [
            table.bbox
            for table in page.find_tables(
                table_settings={
                    'vertical_strategy': 'lines',
                    'horizontal_strategy': 'lines',
                    'explicit_vertical_lines': page.curves + page.edges,
                    'explicit_horizontal_lines': page.curves + page.edges,
                }
            )
        ]

        page_data = re.sub(r'OPC 30050: PackML  \d\d  V 1.01 ', '', page.filter(not_within_bboxes).extract_text())
        page_df = pd.DataFrame({'Sentences': re.split(r'\.\W', page_data)})
        page_df['Sentences'] += '.'
        page_df['Sentences'] = page_df['Sentences'].str.replace(r'\n', ' ')
        pdf_data.append(page_df)


sentences_df = pd.concat(pdf_data)
sentences_df.to_excel('Extracted_sentences.xlsx', index=False)
