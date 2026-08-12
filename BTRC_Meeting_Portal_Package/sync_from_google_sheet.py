import pandas as pd
import json
import re
import os
import sys
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

# User's target Google Sheet ID & GID
SHEET_ID = "1kXlVrNZ9wEGPTrW5ClvGcoi037Oad8rk"
SHEET_GID = "832412813"

# Export URLs
EXPORT_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={SHEET_GID}"
PUB_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/pub?output=csv&gid={SHEET_GID}"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
json_path = os.path.join(BASE_DIR, 'meetings_db.json')
html_path = os.path.join(BASE_DIR, 'index.html')
local_excel_path = os.path.join(BASE_DIR, 'extracted_agendas_live.xlsx')

print("🌐 1/4 Fetching live data from Google Sheet...")

df = None
for url in [EXPORT_CSV_URL, PUB_CSV_URL]:
    try:
        df = pd.read_csv(url)
        print(f"  ✓ Successfully fetched Google Sheet live from: {url}")
        break
    except Exception as e:
        pass

if df is None:
    print("  ⚠️ Notice: Google Sheet requires 'Publish to Web' or 'Anyone with link can view' permission.")
    print("  Falling back to local Excel database...")
    df = pd.read_excel(local_excel_path)

df = df.fillna('')
meetings_dict = {}

for idx, row in df.iterrows():
    m_no = str(row.get('মিটিং নম্বর', '')).strip()
    m_date = str(row.get('মিটিংয়ের তারিখ', '')).strip()
    a_no = str(row.get('এজেন্ডা নম্বর', '')).strip()
    subject = str(row.get('বিষয়', '')).strip()
    decision = str(row.get('সিদ্ধান্ত', '')).strip()
    fine = str(row.get('প্রশাসনিক জরিমানা (টাকা)', '')).strip()
    impl = str(row.get('বাস্তবায়নকারী বিভাগ', '')).strip()
    status = str(row.get('মামলার অবস্থা', '')).strip()

    if not m_no:
        continue

    if m_no not in meetings_dict:
        meetings_dict[m_no] = {
            "meeting_number": m_no,
            "meeting_date": m_date,
            "agendas": []
        }

    meetings_dict[m_no]["agendas"].append({
        "agenda_no": a_no if a_no else str(len(meetings_dict[m_no]["agendas"]) + 1),
        "subject": subject,
        "decision": decision,
        "fine_amount": fine,
        "details": {
            "presentation_summary": f"বিষয়: {subject}",
            "tables": [],
            "implementation": impl if impl else "এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট",
            "assigned_inspector": "",
            "case_status": status if status else "চলমান"
        }
    })

def meeting_sort_key(m):
    num_str = re.sub(r'\D', '', m["meeting_number"])
    return int(num_str) if num_str.isdigit() else 999

sorted_meetings = sorted(list(meetings_dict.values()), key=meeting_sort_key)

coordination_mom = []
commissioner_instructions = []
if os.path.exists(json_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            old_db = json.load(f)
            coordination_mom = old_db.get('coordination_mom', [])
            commissioner_instructions = old_db.get('commissioner_instructions', [])
    except Exception:
        pass

master_db_object = {
    "meetings": sorted_meetings,
    "coordination_mom": coordination_mom,
    "commissioner_instructions": commissioner_instructions
}

print("💾 2/4 Saving JSON Database (meetings_db.json)...")
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(master_db_object, f, ensure_ascii=False, indent=2)

print("🌐 3/4 Updating Web Portal (index.html)...")
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

json_str = json.dumps(master_db_object, ensure_ascii=False, indent=4)

start_idx = html.find('const masterDb = {')
end_idx = html.find('const banglaDigits =', start_idx)
if start_idx != -1 and end_idx != -1:
    substring = html[start_idx:end_idx]
    last_brace_idx = substring.rfind('};')
    new_text = f'const masterDb = {json_str};'
    new_html = html[:start_idx] + new_text + html[start_idx + last_brace_idx + 2:]

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new_html)

# Update local excel backup
try:
    df.to_excel(local_excel_path, index=False)
except Exception as e:
    print(f"  Notice: Could not update Excel backup ({e})")

print("📝 4/4 Updating MS Word Documents (cm word copy)...")
word_dir = os.path.join(BASE_DIR, 'cm word copy')
os.makedirs(word_dir, exist_ok=True)
try:
    for f_name in os.listdir(word_dir):
        if f_name.endswith('.docx'):
            try:
                os.remove(os.path.join(word_dir, f_name))
            except Exception:
                pass

    for meeting in sorted_meetings:
        meeting_no = meeting.get('meeting_number', 'Unknown')
        safe_filename = re.sub(r'[\\/*?:"<>|]', '_', meeting_no)
        filename = os.path.join(word_dir, f"{safe_filename}.docx")
        
        doc = Document()
        doc.add_heading(f"কমিশন সভা: {meeting_no}", 0)
        doc.add_paragraph(f"তারিখ: {meeting.get('meeting_date', '')}")
        
        for a in meeting.get('agendas', []):
            agenda_no = a.get('agenda_no', '')
            doc.add_heading(f"এজেন্ডা {agenda_no}", level=1)
            
            doc.add_heading("বিষয় (Subject):", level=2)
            doc.add_paragraph(a.get('subject', ''))
            
            doc.add_heading("সিদ্ধান্ত (Decision):", level=2)
            doc.add_paragraph(a.get('decision', ''))
            
            if a.get('fine_amount'):
                doc.add_paragraph(f"প্রশাসনিক জরিমানা: {a.get('fine_amount')} টাকা")
                
            doc.add_paragraph("-" * 40)
            
        doc.save(filename)
except Exception as err:
    print(f"  Notice: Word document generation skipped ({err})")

print("🎉 Synchronization Complete!")

