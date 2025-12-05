import re
from pathlib import Path
import pdfplumber
import pandas as pd

PDF_PATH = Path("/Users/amritkumar/Downloads/RBI-Report.pdf")
OUTPUT_XLSX = Path("/Users/amritkumar/Downloads/RBI-Report.xlsx")

#Chek for OCR

ocr_available = True
try:
    from PIL import Image
    import pytesseract
except Exception as e:
    ocr_available = False
 
def pdf_to_text_pages(pdf_path):
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for i , page in enumerate(pdf.pages , start = 1):
            text = page.extract_text()
            if text and text.strip():
                pages_text.append((i,text))

            else:
                if ocr_available:
                    im = page.to_image(resolution=300).original
                    ocr_text = pytesseract.image_to_string(im)
                    pages_text.append((i,ocr_text))
                else:
                    pages_text.append((i,""))
    return pages_text

def find_sheet_and_date(full_text):
    sheet = " "
    eff_date = " "

    sheet_pattern = [
        r'Circular No\.?\s*([A-Za-z0-9\/\-\.\(\) ]+)',
        r'No\.?\s*([A-Za-z0-9\/\-\.\(\) ]{3,30})',
        r'\bFile No\.?\s*[:\-]?\s*([A-Za-z0-9\/\-\.\(\) ]+)',
        r'\bDO No\.?\s*[:\-]?\s*([A-Za-z0-9\/\-\.\(\) ]+)'
    ]

    for pat in sheet_pattern:
        match = re.search(pat, full_text, re.IGNORECASE)
        if match:
            sheet = match.group(1).strip()
            break

    eff_patterns = [
        r'Effective\s+from\s+([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})',
        r'Effective\s+Date\s*[:\-]\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})',
        r'with effect from\s+([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})',
        r'(\b[0-9]{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+[0-9]{4}\b)'
    ]

    for pat in eff_patterns:
        match = re.search(pat , full_text , re.IGNORECASE)

        if match:
            eff_date = match.group(1).strip()
            break
    return sheet , eff_date

def extract_main_heading(full_text):

    lines = full_text.splitlines()

    # keep only first 15 lines for heading search
    lines = [ln.strip() for ln in lines[:15]]

    heading_lines = []
    for ln in lines:
        if ln.strip() == "":
            break  # stop at first blank line
        # ignore RBI address, logos, page numbers etc., if needed add conditions here
        heading_lines.append(ln.strip())

    # join lines into a single heading
    heading = " ".join(heading_lines)
    return heading

def split_into_paragraphs(full_text):
    text = full_text.replace('\r\n', '\n').replace('\r' , '\n')

    para_regex = re.compile(r'(?m)^\s*(\d{1,3}\.)\s*') 

    starts = []
    for match in para_regex.finditer(text):
        starts.append((match.start() , match.group(1).rstrip('.')))

    paragraphs = []
    if not starts:
        # no numbered paras — split by double newline into logical paragraphs
        blocks = [b.strip() for b in re.split(r'\n\s*\n', text) if b.strip()]
        for idx, b in enumerate(blocks, start=1):
            paragraphs.append({
                'para_no': f'{idx}',
                'text': b
            })
        return paragraphs

    for i, (pos, num) in enumerate(starts):
        start_pos = pos
        end_pos = starts[i+1][0] if i+1 < len(starts) else len(text)
        block = text[start_pos:end_pos].strip()
        # remove the leading "44." from block text
        block_text = re.sub(r'^\s*\d{1,3}\.\s*', '', block, count=1)
        paragraphs.append({
            'para_no': num,
            'text': block_text
        })
    return paragraphs

def detect_parent_child(paragraphs):
    for p in paragraphs:
        p['parent'] = ''
    return paragraphs


def assign_heading_candidates(full_text , paragraphs):
    headings = []
    for p in paragraphs:
        # find the exact paragraph occurrence in the full_text (best effort)
        snippet = p['text'][:60].strip()
        if not snippet:
            p['heading'] = ''
            continue
        # escape regex meta chars
        esc = re.escape(snippet)
        m = re.search(esc, full_text)
        heading = ""
        if m:
            # find preceding 200 chars and split into lines
            start = max(0, m.start()-400)
            context = full_text[start:m.start()]
            # look for short line right before paragraph (last non-empty line)
            lines = [ln.strip() for ln in context.splitlines() if ln.strip()]
            if lines:
                cand = lines[-1]
                # heuristics: if candidate is short (<=10 words) and doesn't end with '.' it's likely a header
                if len(cand.split()) <= 10 and not cand.endswith('.'):
                    heading = cand
            if not heading:
                # fallback: build heading from paragraph's first meaningful words (remove section numbers like "44.")
                words = re.findall(r'\w+', p['text'])
                heading = " ".join(words[:8]) + ("..." if len(words)>8 else "")

        else:
            # fallback: first 8 words
            words = re.findall(r'\w+', p['text'])
            heading = " ".join(words[:8]) + ("..." if len(words)>8 else "")
        p['heading'] = heading
    return paragraphs

# Run extraction
pages = pdf_to_text_pages(PDF_PATH)
full_text = "\n".join([t for (_, t) in pages])
sheet, eff_date = find_sheet_and_date(full_text)
main_heading = extract_main_heading(full_text)
paras = split_into_paragraphs(full_text)
paras = detect_parent_child(paras)
paras = assign_heading_candidates(full_text, paras)

# Build DataFrame
rows = []
for i, p in enumerate(paras, start=1):
    rows.append({
        'Seq': i,
        'FileName': PDF_PATH.name,
        'SheetNumber': sheet,
        'EffectiveDate': eff_date,
        'MainHeading': main_heading,
        'ParaNumber': p.get('para_no',''),
        'ParentPara': p.get('parent',''),
        'Heading': p.get('heading',''),
        'ParagraphText': p.get('text','')
    })

df = pd.DataFrame(rows, columns=['Seq','FileName','SheetNumber','EffectiveDate','ParaNumber','ParentPara','Heading','ParagraphText'])

# Save to Excel
df.to_excel(OUTPUT_XLSX, index=False)

# Display first 40 rows for learning purposes
print("Extracted RBI Paragraphs (preview):")
print(df.head(40).to_string())

print(f"Saved full extraction to: {OUTPUT_XLSX}")
