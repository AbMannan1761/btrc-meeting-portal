import pandas as pd
import json
import os
import sys
import re
import xml.etree.ElementTree as ET
import zipfile

# Input file paths
excel_overall_path = r"c:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\Summery_Over all.xlsx"
excel_sharmin_path = r"c:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\CD-PS-15042026V1 from sharmin mam.xlsx"
word_300_path = r"C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\doc_converted.docx"
word_302_path = r"c:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\Agenda_302.docx"
word_307_path = r"c:\Users\mannan\OneDrive - Bangladesh Telecommunication Regulatory Commission\1. E & I\Meeting\307\list of 307 draft.docx"
mom_path = r"C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\mom_cm_sm.docx"
instruction_path = r"C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\cm_sm_instruction.docx"

output_json_path = r"C:\Users\mannan\.gemini\antigravity\brain\d4fb72b8-9e63-42ba-8680-37ffcb1867fa\scratch\meetings_db.json"

# Date formatter helper
def format_meeting_date(cm_no, raw_date):
    raw_date_str = str(raw_date).strip()
    if raw_date_str == "nan" or not raw_date_str:
        pass
    
    # Pre-defined accurate mapping based on chronological analysis and Excel date swap correction
    date_mapping = {
        "278": "৩০ অক্টোবর, ২০২৩",
        "280": "২৯ জানুয়ারি, ২০২৪",
        "281": "২২ ফেব্রুয়ারি, ২০২৪",
        "282": "২৫ মার্চ, ২০২৪",
        "283": "২৫ এপ্রিল, ২০২৪",
        "284": "২৬ মে, ২০২৪",
        "285": "২৭ জুন, ২০২৪",
        "286": "১৩ আগস্ট, ২০২৪",
        "287": "০৬ অক্টোবর, ২০২৪",
        "288": "২০ নভেম্বর, ২০২৪",
        "289": "০১ ডিসেম্বর, ২০২৪",
        "290": "২৯ ডিসেম্বর, ২০২৪",
        "291": "২৭ জানুয়ারি, ২০২৫",
        "292": "২৫ ফেব্রুয়ারি, ২০২৫",
        "293": "২৪ মার্চ, ২০২৫",
        "294": "২১ এপ্রিল, ২০২৫",
        "295": "১৯ মে, ২০২৫",
        "296": "৩০ জুন, ২০২৫",
        "297": "২১ জুলাই, ২০২৫",
        "298": "২৫ আগস্ট, ২০২৫",
        "300": "২৭ অক্টোবর, ২০২৫",
        "301": "১৮ নভেম্বর, ২০২৫",
        "302": "২৪ ডিসেম্বর, ২০২৫",
        "303": "২৪ ডিসেম্বর, ২০২৫",
        "304": "০৮ ফেব্রুয়ারি, ২০২৬",
        "305": "৩০ মার্চ, ২০২৬",
        "306": "২৬ এপ্রিল, ২০২৬",
        "307": "২৪ মে, ২০২৬"
    }
    
    clean_cm = str(cm_no).split('.')[0].replace('তম', '').strip()
    if clean_cm in date_mapping:
        return date_mapping[clean_cm]
    
    return raw_date_str

# Fine extractor helper that works for both English and Bengali digits
def extract_fine(text):
    if not text or str(text).lower() == "nan":
        return ""
    text_str = str(text).strip()
    # Pattern to match numbers (Bengali or English) followed by slashes, hyphens or "টাকা"
    match = re.search(r'([0-9\u09e6-\u09ef,]+)\s*(?:/-|/|টাকা|প্রশাসনিক জরিমানা)', text_str)
    if match:
        val = match.group(1).strip().strip(',')
        if len(val.replace(',', '')) > 2:
            return val
    # Fallback pattern matching long digit sequences
    match_fallback = re.search(r'([0-9\u09e6-\u09ef,]{4,})', text_str)
    if match_fallback:
        val = match_fallback.group(1).strip().strip(',')
        return val
    return ""

# Start building the database
meetings = {}

# 1. Parse Summery_Over all.xlsx
print("Parsing Summery_Over all.xlsx...")
df_overall = pd.read_excel(excel_overall_path, sheet_name='Sheet1', skiprows=1)
df_overall.columns = [
    'sl_no', 'organization', 'category', 'inspection_date', 
    'report_submission_date', 'cm_date', 'cm_no', 'agenda_no', 
    'decision', 'need_info_from', 'remarks'
]
df_overall = df_overall.dropna(subset=['cm_no'])

for idx, row in df_overall.iterrows():
    raw_cm = str(row['cm_no']).strip()
    cm_num = raw_cm.split('.')[0].replace('তম', '').strip()
    
    if not cm_num.isdigit():
        continue
        
    cm_key = f"{cm_num}তম"
    cm_date = format_meeting_date(cm_num, row['cm_date'])
    
    if cm_key not in meetings:
        meetings[cm_key] = {
            "meeting_number": cm_key,
            "meeting_date": cm_date,
            "agendas": []
        }
        
    agenda_no = str(row['agenda_no']).strip().replace('.0', '')
    if agenda_no == "nan" or not agenda_no:
        agenda_no = "বিবিধ"
        
    org = str(row['organization']).strip()
    if org == "nan":
        org = "অন্যান্য বিষয়"
        
    category = str(row['category']).strip()
    if category == "nan":
        category = ""
        
    decision = str(row['decision']).strip()
    if decision == "nan" or not decision:
        decision = "আলোচনার জন্য উপস্থাপন করা হয়নি বা কোনো সিদ্ধান্ত লিপিবব্ধ করা হয়নি।"
        
    need_info = str(row['need_info_from']).strip()
    remarks = str(row['remarks']).strip()
    
    impl = "এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট"
    if need_info != "nan" and need_info:
        impl += f" এবং {need_info}"
    if remarks != "nan" and remarks:
        impl += f" (মন্তব্য: {remarks})"
        
    inspection_date = str(row['inspection_date']).strip()
    submission_date = str(row['report_submission_date']).strip()
    
    bg_text = f"প্রতিষ্ঠান: {org}\n"
    if category:
        bg_text += f"লাইসেন্সের ধরণ: {category}\n"
    if inspection_date != "nan" and inspection_date:
        bg_text += f"পরিদর্শনের তারিখ: {inspection_date}\n"
    if submission_date != "nan" and submission_date:
        bg_text += f"প্রতিবেদন জমাদানের তারিখ: {submission_date}\n"
    bg_text += "প্রেক্ষাপট: নিয়মিত পরিদর্শন ও লাইসেন্সিং শর্তাবলী পরিপালন তদারকির অংশ হিসেবে অভিযান/তদন্ত পরিচালনা করা হয়।"
    
    details_dict = {
        "presentation_summary": bg_text,
        "tables": [],
        "implementation": impl,
        "assigned_inspector": "",
        "case_status": "নিষ্পন্ন" if "জরিমানা আরোপ" in decision else "চলমান"
    }
    
    fine_amount = extract_fine(decision)
        
    meetings[cm_key]["agendas"].append({
        "agenda_no": agenda_no,
        "subject": f"{org} ({category}) এর পরিদর্শন ও আইনানুগ ব্যবস্থা গ্রহণ সংক্রান্ত" if category else f"{org} এর বিষয়াবলী",
        "decision": decision,
        "fine_amount": fine_amount,
        "details": details_dict
    })

# 2. Integrate CD-PS-15042026V1 from sharmin mam.xlsx (merging fine amounts, department claims)
print("Parsing CD-PS-15042026V1 from sharmin mam.xlsx...")
xl_sharmin = pd.ExcelFile(excel_sharmin_path)

def find_matching_agenda(cm_key, org_name):
    if cm_key not in meetings:
        return None
    org_clean = org_name.lower().split(' ')[0].replace('-', '').replace('@', '').replace('.', '').strip()
    if not org_clean or len(org_clean) < 3:
        return None
    for a in meetings[cm_key]["agendas"]:
        sub_lower = a["subject"].lower()
        dec_lower = a["decision"].lower()
        if org_clean in sub_lower or org_clean in dec_lower:
            return a
    return None

for sheet in xl_sharmin.sheet_names:
    df_s = xl_sharmin.parse(sheet)
    print(f"Integrating Sheet: {sheet}")
    
    if sheet == 'Administrative-Fine':
        df_s.columns = [str(c).strip() for c in df_s.iloc[1]]
        df_s = df_s.iloc[2:]
        df_s['Commission Meeting'] = df_s['Commission Meeting'].ffill()
        for idx, row in df_s.iterrows():
            cm_num = str(row.get('Commission Meeting', '')).strip().split('.')[0]
            org = str(row.get('Name of organization', '')).strip()
            fine = str(row.get('Fine amount', '')).strip()
            status = str(row.get('Payment Status', '')).strip()
            remark = str(row.get('Remark', '')).strip()
            lic_type = str(row.get('Type of Lisence', '')).strip()
            
            if not cm_num or not org or org == "nan" or cm_num == "nan":
                continue
                
            cm_key = f"{cm_num}তম"
            if cm_key not in meetings:
                meetings[cm_key] = {"meeting_number": cm_key, "meeting_date": format_meeting_date(cm_num, ""), "agendas": []}
                
            a = find_matching_agenda(cm_key, org)
            if a:
                if fine and fine != "nan":
                    a["fine_amount"] = fine
                if status and status != "nan":
                    a["details"]["case_status"] = status
                    a["decision"] += f"\n[পেমেন্ট স্ট্যাটাস: {status}]"
                if remark and remark != "nan":
                    a["decision"] += f"\n[মন্তব্য: {remark}]"
            else:
                meetings[cm_key]["agendas"].append({
                    "agenda_no": "সংযুক্ত",
                    "subject": f"{org} ({lic_type if lic_type != 'nan' else 'ISP'}) এর প্রশাসনিক জরিমানা",
                    "decision": f"প্রশাসনিক জরিমানা: {fine if fine != 'nan' else '০'}/- টাকা।\nপেমেন্ট স্ট্যাটাস: {status if status != 'nan' else 'চলমান'}\nমন্তব্য: {remark if remark != 'nan' else '-'}",
                    "fine_amount": fine if fine != "nan" else "",
                    "details": {
                        "presentation_summary": f"প্রতিষ্ঠান: {org}\nলাইসেন্স ধরণ: {lic_type}",
                        "tables": [],
                        "implementation": "অর্থ, হিসাব ও রাজস্ব বিভাগ এবং ইএন্ডআই ডিরেক্টরেট",
                        "assigned_inspector": "",
                        "case_status": status if status != "nan" else "চলমান"
                    }
                })

    elif sheet == 'Revenue-Sharing':
        df_s.columns = [str(c).strip() for c in df_s.iloc[1]]
        df_s = df_s.iloc[2:]
        df_s['Commission Meeting'] = df_s['Commission Meeting'].ffill()
        for idx, row in df_s.iterrows():
            cm_num = str(row.get('Commission Meeting', '')).strip().split('.')[0]
            org = str(row.get('Name of organization', '')).strip()
            claim = str(row.get('Fine amount', '')).strip()
            remark = str(row.get('Remarks', '')).strip()
            lic_type = str(row.get('Type of Lisence', '')).strip()
            
            if not cm_num or not org or org == "nan" or cm_num == "nan":
                continue
                
            cm_key = f"{cm_num}তম"
            if cm_key not in meetings:
                meetings[cm_key] = {"meeting_number": cm_key, "meeting_date": format_meeting_date(cm_num, ""), "agendas": []}
                
            a = find_matching_agenda(cm_key, org)
            if a:
                if claim and claim != "nan":
                    a["revenue_sharing_claim"] = claim
                    a["decision"] += f"\n[রেভিনিউ শেয়ারিং বকেয়া দাবি: {claim}/- টাকা]"
                if remark and remark != "nan":
                    a["decision"] += f"\n[মন্তব্য: {remark}]"
            else:
                meetings[cm_key]["agendas"].append({
                    "agenda_no": "সংযুক্ত",
                    "subject": f"{org} ({lic_type if lic_type != 'nan' else 'Aggregator'}) এর রেভিনিউ শেয়ারিং ও বকেয়া দাবি",
                    "decision": f"রেভিনিউ শেয়ারিং বকেয়া দাবি: {claim if claim != 'nan' else '০'}/- টাকা।\nমন্তব্য: {remark if remark != 'nan' else '-'}",
                    "fine_amount": "",
                    "revenue_sharing_claim": claim if claim != 'nan' else "",
                    "details": {
                        "presentation_summary": f"প্রতিষ্ঠান: {org}\nলাইসেন্স ধরণ: {lic_type}",
                        "tables": [],
                        "implementation": "Finance, Revenue & E&I Division",
                        "assigned_inspector": "",
                        "case_status": "চলমান"
                    }
                })

    elif sheet == 'Legal':
        df_s.columns = [str(c).strip() for c in df_s.iloc[1]]
        df_s = df_s.iloc[2:]
        df_s['Commission Meeting'] = df_s['Commission Meeting'].ffill()
        for idx, row in df_s.iterrows():
            cm_num = str(row.get('Commission Meeting', '')).strip().split('.')[0]
            org = str(row.get('Name of organization', '')).strip()
            claim_type = str(row.get('Type of Claim', '')).strip()
            lic_type = str(row.get('Type of Lisence', '')).strip()
            
            if not cm_num or not org or org == "nan" or cm_num == "nan":
                continue
                
            cm_key = f"{cm_num}তম"
            if cm_key not in meetings:
                meetings[cm_key] = {"meeting_number": cm_key, "meeting_date": format_meeting_date(cm_num, ""), "agendas": []}
                
            a = find_matching_agenda(cm_key, org)
            if a:
                a["decision"] += f"\n[আইনি/লাইসেন্স সংক্রান্ত দাবি: {claim_type}]"
            else:
                meetings[cm_key]["agendas"].append({
                    "agenda_no": "সংযুক্ত",
                    "subject": f"{org} ({lic_type}) এর লাইসেন্স নবায়ন/শেয়ার হস্তান্তর সংক্রান্ত বিষয়",
                    "decision": f"লাইসেন্স/আইনি সংক্রান্ত ব্যত্যয়: {claim_type}",
                    "fine_amount": "",
                    "details": {
                        "presentation_summary": f"প্রতিষ্ঠান: {org}\nলাইসেন্স ধরণ: {lic_type}\nদাবি: {claim_type}",
                        "tables": [],
                        "implementation": "লিগ্যাল অ্যান্ড লাইসেন্সিং বিভাগ এবং ইএন্ডআই ডিরেক্টরেট",
                        "assigned_inspector": "",
                        "case_status": "চলমান"
                    }
                })

    elif sheet == 'Clause 65(4)':
        df_s.columns = [str(c).strip() for c in df_s.iloc[1]]
        df_s = df_s.iloc[2:]
        df_s['Commission Meeting'] = df_s['Commission Meeting'].ffill()
        for idx, row in df_s.iterrows():
            cm_num = str(row.get('Commission Meeting', '')).strip().split('.')[0]
            org = str(row.get('Name of organization', '')).strip()
            claim_type = str(row.get('Type of Claim', '')).strip()
            lic_type = str(row.get('Type of Lisence', '')).strip()
            fine = str(row.get('Fine amount', '')).strip()
            inspector = str(row.get('Assigned Inspector', '')).strip()
            status = str(row.get('Status', '')).strip()
            
            if not cm_num or not org or org == "nan" or cm_num == "nan":
                continue
                
            cm_key = f"{cm_num}তম"
            if cm_key not in meetings:
                meetings[cm_key] = {"meeting_number": cm_key, "meeting_date": format_meeting_date(cm_num, ""), "agendas": []}
                
            a = find_matching_agenda(cm_key, org)
            if a:
                if fine and fine != "nan" and not a.get("fine_amount"):
                    a["fine_amount"] = fine
                if inspector and inspector != "nan":
                    a["details"]["assigned_inspector"] = inspector
                if status and status != "nan":
                    a["details"]["case_status"] = status
                a["decision"] += f"\n[ধারা ৬৫(৪) জরিমানা ও কেইস স্ট্যাটাস: {status if status != 'nan' else 'চলমান'}]"
            else:
                meetings[cm_key]["agendas"].append({
                    "agenda_no": "সংযুক্ত",
                    "subject": f"{org} ({lic_type}) এর ধারা ৬৫(৪) ব্যত্যয়",
                    "decision": f"জরিমানা: {fine}/- টাকা।\nঅবস্থা: {status if status != 'nan' else 'চলমান'}\nপরিদর্শক: {inspector if inspector != 'nan' else '-'}",
                    "fine_amount": fine if fine != "nan" else "",
                    "details": {
                        "presentation_summary": f"প্রতিষ্ঠান: {org}\nলাইসেন্স ধরণ: {lic_type}\nঅভিযোগের ধরণ: {claim_type}",
                        "tables": [],
                        "implementation": "ইএন্ডআই ডিরেক্টরেট",
                        "assigned_inspector": inspector if inspector != "nan" else "",
                        "case_status": status if status != "nan" else "চলমান"
                    }
                })

    elif sheet == 'E&I':
        df_s.columns = [str(c).strip() for c in df_s.columns]
        df_s['Commission Meeting'] = df_s['Commission Meeting'].ffill()
        for idx, row in df_s.iterrows():
            cm_num = str(row.get('Commission Meeting', '')).strip().split('.')[0]
            org = str(row.get('Name of organization', '')).strip()
            claim = str(row.get('Type of Claim', '')).strip()
            lic = str(row.get('Type of Lisence', '')).strip()
            fine = str(row.get('Fine amount', '')).strip()
            dept = str(row.get('Department', '')).strip()
            action = str(row.get('Unnamed: 6', '')).strip()
            
            if not cm_num or not org or org == "nan" or cm_num == "nan":
                continue
                
            cm_key = f"{cm_num}তম"
            if cm_key not in meetings:
                meetings[cm_key] = {"meeting_number": cm_key, "meeting_date": format_meeting_date(cm_num, ""), "agendas": []}
                
            a = find_matching_agenda(cm_key, org)
            if a:
                if action and action != "nan":
                    a["decision"] += f"\n[এনফোর্সমেন্ট অ্যাকশন: {action}]"
                if fine and fine != "nan" and not a.get("fine_amount"):
                    a["fine_amount"] = fine
            else:
                meetings[cm_key]["agendas"].append({
                    "agenda_no": "সংযুক্ত",
                    "subject": f"{org} ({lic}) এর ইএন্ডআই এনফোর্সমেন্ট অ্যাকশন",
                    "decision": f"এনফোর্সমেন্ট অ্যাকশন: {action if action != 'nan' else 'পর্যালোচনাধীন'}\nদাবি: {claim if claim != 'nan' else '-'}",
                    "fine_amount": fine if fine != "nan" else "",
                    "details": {
                        "presentation_summary": f"প্রতিষ্ঠান: {org}\nলাইসেন্স ধরণ: {lic}\nদাবি: {claim}",
                        "tables": [],
                        "implementation": f"{dept if dept != 'nan' else 'ইএন্ডআই ডিরেক্টরেট'}",
                        "assigned_inspector": "",
                        "case_status": "চলমান"
                    }
                })

# 3. Inject detailed Word Documents (Meetings 300, 302, 307)

# 3.1 Meeting 300 (Siam Online BD)
print("Injecting detailed 300th meeting data...")
meetings["300তম"] = {
    "meeting_number": "৩০০তম",
    "meeting_date": "২৭ অক্টোবর, ২০২৫",
    "agendas": [
        {
            "agenda_no": "০১",
            "subject": "আইএসপি লাইসেন্সধারী প্রতিষ্ঠান Siam Online BD কর্তৃক কমিশনের জারিকৃত অপারেশনাল নির্দেশনাসহ আইএসপি গাইডলাইন ও লাইসেন্স এর শর্ত ভঙ্গ করায় প্রয়োজনীয় আইনানুগ ব্যবস্থা গ্রহণ প্রসঙ্গে।",
            "decision": "উত্তরা (পশ্চিম) উপজেলা/থানা আইএসপি লাইসেন্সধারী প্রতিষ্ঠান Siam Online BD কর্তৃক অনুমোদন ব্যতীত অফিসের ঠিকানা পরিবর্তন করা, DIS-এ PoP ঘোষণা না করা, বিলিং সিস্টেম না থাকা এবং উত্তরা (পশ্চিম) উপজেলা/থানা আইএসপি লাইসেন্স হওয়া সত্ত্বেও উত্তরা (পূর্ব) ও দক্ষিণখান থানায় ২৫০টি সংযোগ প্রদান করার কারণে আইএসপি গাইডলাইন ও লাইসেন্সের শর্ত এবং কমিশনের নির্দেশনা ভঙ্গ করার অপরাধ সংঘটিত হওয়ায় বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ আইন, ২০০১ এর ধারা ৬৫(২) ও ৬৫(৩) এর বিধান মোতাবেক প্রতিষ্ঠানটির উপর ৭৫,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করে আদায় সাপেক্ষে ৮০% ব্যান্ডউইডথ ক্যাপিং অপসারণের সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "৭৫,০০০",
            "details": {
                "presentation_summary": "কমিশনে দাখিলকৃত অভিযোগের ভিত্তিতে ও নিয়মিত পরিদর্শনের অংশ হিসেবে কমিশনের এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট এর পরিদর্শকদল Siam Online BD নামক উত্তরা (পশ্চিম) উপজেলা/থানা আইএসপি লাইসেন্সধারী প্রতিষ্ঠানের অফিস/স্থাপনা গত ১১/০৮/২০২৫ তারিখে সরেজমিন পরিদর্শন করে এবং পরিদর্শনে প্রাপ্ত ব্যত্যয়ের উপর কারণ দর্শানো নোটিশ জারি করা হয়। প্রতিষ্ঠানটি তাদের ব্যত্যয়সমূহ স্বীকার করে ক্ষমা চেয়ে আবেদন করেছে।",
                "tables": [
                    {
                        "headers": ["ব্যত্যয়ের ক্ষেত্র", "বিবরণ", "গাইডলাইন/আইনি বিধি"],
                        "rows": [
                            ["অফিস স্থানান্তর", "বিটিআরসি'র অনুমোদন ব্যতিত লাইসেন্সের অফিস ঠিকানা পরিবর্তন করা।", "লাইসেন্সের শর্ত লঙ্ঘন"],
                            ["অবৈধ PoP স্থাপন", "১৮/সি, রোড নং- ৭/বি, সেক্টর- ৩, উত্তরা (পশ্চিম) এ অবৈধ নেটওর্য়াকিং ডিভাইস ও PoP স্থাপন।", "কমিশনের PoP সংক্রান্ত নির্দেশনা"],
                            ["বিলিং সিস্টেম", "Siam Online BD এর কোন বিলিং সিস্টেম নেই।", "অপারেশনাল নির্দেশনাবলী"],
                            ["এলাকা বহির্ভূত সংযোগ", "উত্তরা (পশ্চিম) থানা লাইসেন্স হওয়া সত্ত্বেও উত্তরা (পূর্ব) এবং দক্ষিণখান থানায় ২৫০টি সংযোগ প্রদান।", "আইএসপি গাইডলাইন ক্লজ ৫.৪"]
                        ]
                    },
                    {
                        "headers": ["প্রতিষ্ঠান সম্পর্কিত তথ্য", "পরিসংখ্যান / মান"],
                        "rows": [
                            ["মোট গ্রাহক সংখ্যা", "১,২০০ জন"],
                            ["আপস্ট্রিম ব্যান্ডউইডথ", "৪০৯ এমবিপিএস"],
                            ["আপস্ট্রিম প্রদানকারী", "Startrek Telecom Ltd."],
                            ["পরিদর্শনে গৃহীত ব্যবস্থা", "৮০% ব্যান্ডউইডথ ক্যাপিং ও অবৈধ PoP মালামাল জব্দ"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ, হিসাব ও রাজস্ব বিভাগ।",
                "assigned_inspector": "পরিদর্শক দল (E&I)",
                "case_status": "জরিমানা আদায় সাপেক্ষে নিষ্পন্ন"
            }
        }
    ]
}

# 3.2 Meeting 302 (Mobile Operator Coverage and ISP Always On Network)
print("Injecting detailed 302nd meeting data...")
meetings["302তম"] = {
    "meeting_number": "৩০২তম",
    "meeting_date": "২৪ ডিসেম্বর, ২০২৫",
    "agendas": [
        {
            "agenda_no": "১৬",
            "subject": "এনফোর্সমেন্ট অ্যান্ড ইন্সপেকশন ডিরেক্টরেট-এর জন্য মাইক্রোকার ভাড়ার প্রস্তাবনা: আলোচনা ও সিদ্ধান্ত গ্রহণ প্রসঙ্গে।",
            "decision": "অভিযান ব্যয়ের অব্যবহৃত বার্ষিক বরাদ্দকৃত ৩৫ লক্ষ টাকার সীমার মধ্যে থেকে E&I ডিরেক্টরেটের অভিযান ও দাপ্তরিক পরিদর্শনের লজিস্টিক চাহিদা মেটানোর জন্য মাইক্রোকার ভাড়ার প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "কমিশনের এনফোর্সমেন্ট অ্যান্ড ইন্সপেকশন ডিরেক্টরেট (E&I) দেশের বিভিন্ন স্থানে সংবেদনশীল ও আইন-শৃঙ্খলা বাহিনী সংশ্লিষ্ট অভিযান পরিচালনা করে থাকে। পর্যাপ্ত নিজস্ব যানের অভাবে অভিযানে তাৎক্ষণিক গতিশীলতা বজায় রাখা সম্ভব হচ্ছিল না। E&I এর বার্ষিক অভিযান বাজেট ৩৫ লক্ষ টাকার বিপরীতে বিগত অর্থবছরগুলোতে তুলনামূলক অনেক কম ব্যয় (যেমন ২০২৩-২৪ এ ১৯.৬৮ লক্ষ এবং ২০২৪-২৫ এ মাত্র ২.৫৩ লক্ষ টাকা) হওয়ায় পর্যাপ্ত আর্থিক সংকুলান রয়েছে।",
                "tables": [
                    {
                        "headers": ["অর্থবছর", "বরাদ্দকৃত বাজেট (টাকা)", "প্রকৃত ব্যয় (টাকা)", "অবশিষ্ট অব্যবহৃত বাজেট"],
                        "rows": [
                            ["২০২২-২৩", "৩৫,০০,০০০/-", "৪,২৪,৯১৮/-", "৩০,৭৫,০৮২/-"],
                            ["২০২৩-২৪", "৩৫,০০,০০০/-", "১৯,৬৮,৪৪৩/-", "১৫,৩১,৫৫৭/-"],
                            ["২০২৪-২৫ (চলতি)", "৩৫,০০,০০০/-", "২,৫৩,৯১০/-", "৩২,৪৬,০ ৯০/-"]
                        ]
                    }
                ],
                "implementation": "প্রশাসন বিভাগ ও অর্থ বিভাগ।",
                "assigned_inspector": "",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৭",
            "subject": "Always On Network Ltd. (ISP) এর লাইসেন্স শর্তভঙ্গ এবং অবৈধ নেটওয়ার্ক ও ট্যারিফ লংঘন প্রসঙ্গে আইনানুগ ব্যবস্থা গ্রহণ।",
            "decision": "ক) ২,০০,০০,০০০/- টাকা (দুই কোটি) DIS এ সঠিক তথ্য প্রদান না করার জন্য প্রশাসনিক জরিমানা আরোপ।\nখ) অনুমোদন ব্যাতীত লাইসেন্স ব্যান্ডের ৫ গিগাহার্জ ব্যান্ডের ২০ মেগাহার্জ ব্যবহার এবং NoC হতে PoP পর্যন্ত NTTN ব্যতীত নিজস্ব ব্যবস্থাপনায় অভার হেড ক্যাবল ব্যবহারের জন্য কারণ দর্শানোর পত্র প্রেরণ।\nগ) কমিশনের অনুমতি ব্যতীত শেয়ার ট্র্যান্সফার এর জন্য ৫.৫% শেয়ার ট্রান্সফার ফি সহ ১৫% বিলম্ব ফি প্রদানের জন্য পত্র প্রেরণ।\nঘ) পরবর্তী নির্দেশনা না দেয়া পর্যন্ত ৫০% ক্যাপিং এবং জরিমানা সমূহ আদায় সাপেক্ষে ক্যাপিং অপসরণ।",
            "fine_amount": "২,০০,০০,০০০",
            "details": {
                "presentation_summary": "পরিদর্শনকালে দেখা যায় যে, প্রতিষ্ঠানটি কমিশনের DIS পোর্টালে সম্পূর্ণ মিথ্যা ও অসত্য তথ্য পরিবেশন করেছে। এছাড়াও লাইসেন্সিং শর্তাবলী লংঘন করে নিজস্ব ওভারহেড ক্যাবল টানা এবং অনুমতি বিহীন স্পেকট্রাম ব্যান্ডউইডথ ব্যবহারের প্রমান পাওয়া গিয়েছে।",
                "tables": [],
                "implementation": "ইএন্ডআই ডিরেক্টরেট, অর্থ বিভাগ ও স্পেকট্রাম বিভাগ।",
                "assigned_inspector": "DD, Golam Sorwar",
                "case_status": "ক্যাপিং ও কারণ দর্শানো"
            }
        },
        {
            "agenda_no": "১৮",
            "subject": "Velocity Networks Ltd. (IIG) এর লাইসেন্স শর্ত লঙ্ঘন ও অবৈধ লিংক স্থাপন প্রসঙ্গে আইনানুগ ব্যবস্থা গ্রহণ।",
            "decision": "কমিশনকে অবহিত না করে অবৈধভাবে আইপিএলসি (IPLC) সার্কিট স্থাপনের কার্যক্রম গ্রহণ, লাইসেন্সের ২১.২ এবং ২৯ এর শর্ত লঙ্ঘন, নিজস্ব বিলিং সিস্টেম না থাকা এবং গাজীপুর হতে NoC পর্যন্ত NTTN ছাড়া অভারহেড ফাইবার ক্যাবল স্থাপনের জন্য প্রতিষ্ঠানটির উপর ৫০,০০,০০০/- (পঞ্চাশ লক্ষ) টাকা প্রশাসনিক জরিমানা আরোপ করা হলো।",
            "fine_amount": "৫০,০০,০০০",
            "details": {
                "presentation_summary": "এনফোর্সমেন্ট দলের পরিদর্শনে Velocity Networks Ltd. এর গাজীপুর লিংক এবং বিলিং ড্যাশবোর্ডে চরম ব্যত্যয় পাওয়া যায়। লাইসেন্সের শর্ত ভঙ্গ করে বিটিআরসিকে না জানিয়ে সার্কিট টানা হয়েছে।",
                "tables": [],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ, হিসাব ও রাজস্ব বিভাগ।",
                "assigned_inspector": "DD, Taifur Rahman",
                "case_status": "জরিমানা আরোপিত"
            }
        },
        {
            "agenda_no": "১৯",
            "subject": "Udayan Online Ltd. (ISP) এর লাইসেন্সিং শর্ত পরিপালন সংক্রান্ত ব্যত্যয় প্রসঙ্গে।",
            "decision": "প্রতিষ্ঠানটির পরিদর্শনে প্রাপ্ত ব্যত্যয়ের প্রেক্ষিতে কেন লাইসেন্স বাতিলসহ আইনানুগ ব্যবস্থা নেওয়া হবে না, তা জানতে চেয়ে কারণ দর্শানোর (Show Cause) নোটিশ প্রেরণের সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "E&I পরিদর্শক দলের পরিদর্শনে উদয়ন অনলাইন লিমিটেড এর PoP ও অফিস ঠিকানা এবং গ্রাহক তালিকায় অসঙ্গতি পাওয়া যায়।",
                "tables": [],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও এলএল বিভাগ।",
                "assigned_inspector": "",
                "case_status": "কারণ দর্শানোর নোটিশ"
            }
        },
        {
            "agenda_no": "QoS-০১",
            "subject": "মোবাইল ফোন অপারেটরদের নেটওয়ার্ক কাভারেজ প্রাপ্যতা, রিচার্জ ব্যালেন্স ও ট্যারিফ বাস্তবায়ন সংক্রান্ত বিষয়ে টাঙ্গাইল, বগুড়া, গাইবান্ধা, সিলেট, কুমিল্লা, চট্টগ্রাম, ময়মনসিংহ ও শেরপুর জেলাসমূহে পরিচালিত পরিদর্শনে প্রাপ্ত ফলাফলসমূহ অধিকতর পর্যালোচনা প্রসঙ্গে।",
            "decision": "মোবাইল ফোন অপারেটরদের নেটওয়ার্ক ও গ্রাহক সেবা পরিদর্শনের প্রাপ্ত তথ্যসমূহ কমিশন কর্তৃক বিশদভাবে পর্যালোচনা করা হলো এবং রিচার্জ ব্যালেন্স ও ট্যারিফ বাস্তবায়ন পরিবীক্ষণ জোরদার করার লক্ষ্যে অপারেটরদের জন্য স্পেকট্রাম ও ট্যারিফ নির্দেশনাবলী সংশোধনের সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "মোবাইল অপারেটরদের সেবার মানোন্নয়ন ও গ্রাহক স্বার্থ নিশ্চিত করার লক্ষ্যে বিটিআরসির এনফোর্সমেন্ট এন্ড ইন্সপেকশন ডিরেক্টরেট, ইঞ্জিনিয়ারিং এন্ড অপারেশন্স বিভাগ, স্পেকট্রাম বিভাগ এবং সিস্টেমস এন্ড সার্ভিসস বিভাগের যৌথ টিম ও ৪টি মোবাইল অপারেটরের প্রতিনিধিদের সমন্বয়ে টাঙ্গাইল, বগুড়া, সিলেট, মৌলভীবাজার, কুমিল্লা, চট্টগ্রাম, ময়মনসিংহ, শেরপুরসহ ১২টি জেলায় সরেজমিনে ড্রাইভ টেস্ট ও ফিল্ড ভিজিট সম্পন্ন হয়।",
                "tables": [
                    {
                        "headers": ["পরিদর্শনের তারিখ", "পরিদর্শনকৃত জেলার নাম", "অফিস আদেশ নম্বর", "যৌথ পরিদর্শন টিম"],
                        "rows": [
                            ["০৭-১০ জুলাই, ২০২৫", "টাঙ্গাইল, বগুড়া ও গাইবান্ধা", "১৪.৩২.০০০০.০০০.৪০০.২৫.০০০৭.২২.১১৬৭", "ইএন্ডআই, ইও, স্পেকট্রাম, এসএস বিভাগ ও অপারেটর প্রতিনিধি"],
                            ["২০-২৪ জুলাই, ২০২৫", "হবিগঞ্জ, মৌলভীবাজার ও সিলেট", "১৪.৩২.০০০০.০০০.৪০০.২৫.০০০৭.২২.১১৬৭", "ইএন্ডআই, ইও, স্পেকট্রাম, এসএস বিভাগ ও অপারেটর প্রতিনিধি"],
                            ["২৮-৩১ জুলাই, ২০২৫", "কুমিল্লা, নোয়াখালী এবং চট্টগ্রাম", "১৪.৩২.০০০০.০০০.৪০০.২৫.০০০৭.২২.১২৫৪", "ইএন্ডআই, ইও, স্পেকট্রাম, এসএস বিভাগ ও অপারেটর প্রতিনিধি"],
                            ["১১-১৪ অগাস্ট, ২০২৫", "ময়মনসিংহ, জামালপুর ও শেরপুর", "১৪.৩২.০০০০.০০০.৪০০.২৫.০০০৭.২২.১২৫৪", "ইএন্ডআই, ইও, স্পেকট্রাম, এসএস বিভাগ ও অপারেটর প্রতিনিধি"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই, স্পেকট্রাম, সিস্টেমস এন্ড সার্ভিসেস ডিরেক্টরেট এবং মোবাইল অপারেটরবৃন্দ।",
                "assigned_inspector": "যৌথ পরিদর্শন দল",
                "case_status": "নীতিমালা সংশোধন"
            }
        }
    ]
}

# 3.2.5 Meeting 306 (Level-3, Idea Tec, Asia Pacific, Sajid Trading, HRC, Pan M Tech, Delta Software, France Bangla, QoS, Global Voice, Earth Telecommunication, Wintel, A2P SMS)
print("Injecting detailed 306th meeting data...")
meetings["306তম"] = {
    "meeting_number": "৩০৬তম",
    "meeting_date": "২৬ এপ্রিল, ২০২৬",
    "agendas": [
        {
            "agenda_no": "০১",
            "subject": "লেভেল-৩ ক্যারিয়ার লিমিটেড (Level-3 Carrier Limited) এর আইআইজি (IIG) লাইসেন্সের বকেয়া রাজস্ব আদায় এবং লাইসেন্স নবায়ন অনুমোদন প্রসঙ্গে।",
            "decision": "লেভেল-৩ ক্যারিয়ার লিমিটেডকে তাদের আইআইজি লাইসেন্স নবায়নের অনুমতি প্রদান করা হলো, তবে শর্ত থাকে যে আগামী তিন মাসের মধ্যে বকেয়া রাজস্বের ৫০% পরিশোধ করতে হবে এবং অবশিষ্ট ৫০% পরবর্তী ১২টি সমান মাসিক কিস্তিতে প্রদান করতে হবে।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "লেভেল-৩ ক্যারিয়ার লিমিটেড (IIG) এর লাইসেন্স নবায়ন এবং রাজস্ব বকেয়া সংক্রান্ত আবেদন বিবেচনা করা হয়। প্রতিষ্ঠানটির মোট বকেয়া রাজস্বের পরিমাণ পর্যালোচনা করে কিস্তি সুবিধা প্রদানের সিদ্ধান্ত গ্রহণ করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "বকেয়া রাজস্বের পরিমাণ", "কমিশনের শর্ত", "লাইসেন্স নবায়নের স্থিতি"],
                        "rows": [
                            ["IIG", "১৫,০০,০০,০০০/- টাকা", "৩ মাসের মধ্যে ৫০% ব্যাংক গ্যারান্টি প্রদান এবং অবশিষ্ট ১২ কিস্তিতে পরিশোধ", "শর্তসাপেক্ষ নবায়ন অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড ও স্পেকট্রাম বিভাগ",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০২",
            "subject": "আইডিয়া টেক লিমিটেড (Idea Tec Limited) এর শেয়ার স্থানান্তর অনুমোদন এবং বিলম্ব ফি মওকুফ সংক্রান্ত আবেদন পর্যালোচনা।",
            "decision": "আইডিয়া টেক লিমিটেড (Nationwide ISP) এর শেয়ার স্থানান্তর অনুমোদন করা হলো। তবে শেয়ার স্থানান্তরে বিলম্বের জন্য প্রযোজ্য বিলম্ব ফি বাবদ ৫০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো এবং তা পরবর্তী ৩০ দিনের মধ্যে পরিশোধের নির্দেশ দেওয়া হলো।",
            "fine_amount": "৫০,০০০",
            "details": {
                "presentation_summary": "আইডিয়া টেক লিমিটেড এর শেয়ার স্থানান্তর সংক্রান্ত নথি এবং আরজেএসসি (RJSC) এর তথ্য পর্যালোচনা করা হয়। লাইসেন্সের শর্ত অনুযায়ী অগ্রিম অনুমতি ব্যতীত শেয়ার স্থানান্তরের কারণে জরিমানা আরোপ করা হয়।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "শেয়ার স্থানান্তরের পরিমাণ", "আরোপিত জরিমানা", "পরিশোধের সময়সীমা"],
                        "rows": [
                            ["Idea Tec Limited", "Nationwide ISP", "১০০% শেয়ার স্থানান্তর", "৫০,০০০/- টাকা", "৩০ দিন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "জরিমানা আরোপ"
            }
        },
        {
            "agenda_no": "০৩",
            "subject": "এশিয়া প্যাসিফিক কমিউনিকেশন (Asia Pacific Communication) এর শেয়ার স্থানান্তর অনুমোদন এবং বিলম্ব ফি মওকুফ সংক্রান্ত আবেদন পর্যালোচনা।",
            "decision": "এশিয়া প্যাসিফিক কমিউনিকেশন (Nationwide ISP) এর শেয়ার স্থানান্তর অনুমোদন করা হলো। তবে লাইসেন্সের শর্ত ভঙ্গ করে অনুমতি ব্যতীত শেয়ার স্থানান্তরের জন্য প্রযোজ্য বিলম্ব ফি বাবদ ৭৫,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো এবং তা পরবর্তী ৩০ দিনের মধ্যে পরিশোধের নির্দেশ দেওয়া হলো।",
            "fine_amount": "৭৫,০০০",
            "details": {
                "presentation_summary": "এশিয়া প্যাসিফিক কমিউনিকেশন লিমিটেড এর শেয়ার স্থানান্তর ও লাইসেন্স বিধি ভঙ্গের বিষয়টি পর্যালোচনা করা হয়। অগ্রিম অনুমতি ছাড়া মালিকানা পরিবর্তনের জন্য জরিমানা আরোপপূর্বক শেয়ার স্থানান্তর বৈধ করা হলো।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "শেয়ার স্থানান্তরের ধরণ", "আরোপিত জরিমানা", "পরিশোধের সময়সীমা"],
                        "rows": [
                            ["Asia Pacific Communication Ltd", "Nationwide ISP", "অংশীদারী শেয়ার স্থানান্তর", "৭৫,০০০/- টাকা", "৩০ দিন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "জরিমানা আরোপ"
            }
        },
        {
            "agenda_no": "০৪",
            "subject": "মেসার্স সাজিদ ট্রেডিং (M/S Sajid Trading) এর ন্যাশনওয়াইড আইএসপি লাইসেন্স সাজিদ ট্রেডিং লিমিটেড এ রূপান্তর এবং বিধি লঙ্ঘনজনিত ব্যত্যয় পর্যালোচনা।",
            "decision": "মেসার্স সাজিদ ট্রেডিং এর লাইসেন্সটি সাজিদ ট্রেডিং লিমিটেড নামে হস্তান্তরের আবেদন শর্তসাপেক্ষে মঞ্জুর করা হলো। লাইসেন্সের শর্ত লঙ্ঘন করে অগ্রিম অনুমতি ব্যতীত রূপান্তরের জন্য ১,০০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো।",
            "fine_amount": "১,০০,০০০",
            "details": {
                "presentation_summary": "সাজিদ ট্রেডিং (Nationwide ISP) এর লাইসেন্স হস্তান্তর এবং আরজেএসসি (RJSC) এর নথিপত্র পর্যালোচনা করা হয়। লাইসেন্সধারী প্রতিষ্ঠানটি পূর্ব অনুমতি ছাড়া লিমিটেড কোম্পানিতে রূপান্তরিত হওয়ায় এই ব্যবস্থা নেওয়া হয়।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "ব্যত্যয়ের ধরণ", "প্রশাসনিক জরিমানা", "স্থিতি"],
                        "rows": [
                            ["M/S Sajid Trading", "Nationwide ISP", "অনুমতি ব্যতীত লিমিটেড কোম্পানিতে রূপান্তর ও হস্তান্তর", "১,০০,০০০/- টাকা", "জরিমানা পরিশোধ সাপেক্ষে হস্তান্তর"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "জরিমানা আরোপ"
            }
        },
        {
            "agenda_no": "০৫",
            "subject": "এইচআরসি টেকনোলজিস লিমিটেড (HRC Technologies Limited) এর ন্যাশনওয়াইড আইএসপি এবং আইপিটিএসপি (IPTSP) লাইসেন্সের বকেয়া রাজস্ব আদায় এবং নবায়ন পর্যালোচনা।",
            "decision": "এইচআরসি টেকনোলজিস লিমিটেডকে তাদের লাইসেন্স নবায়নের অনুমতি প্রদান করা হলো, তবে শর্ত থাকে যে বকেয়া রাজস্বের প্রযোজ্য ভ্যাট ও লেট ফিসহ বকেয়া আগামী ৩ মাসের মধ্যে পরিশোধ করতে হবে।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "এইচআরসি টেকনোলজিস লিমিটেড এর দীর্ঘদিনের বকেয়া রাজস্ব এবং নবায়নের ফাইল পর্যালোচনা করা হয়। প্রতিষ্ঠানটির আবেদনের প্রেক্ষিতে কিস্তি ও ভ্যাট সংক্রান্ত বকেয়া মিটমাট করার শর্তারোপ করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "লাইসেন্স ইস্যুর তারিখ", "বকেয়া স্থিতি", "মন্তব্য"],
                        "rows": [
                            ["Nationwide ISP & IPTSP", "০৩-০৯-২০০৯", "বকেয়া প্রযোজ্য", "ভ্যাট ও লেট ফিসহ বকেয়া পরিশোধ সাপেক্ষে নবায়ন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ বিভাগ।",
                "assigned_inspector": "অর্থ বিভাগ ও ইএন্ডআই",
                "case_status": "শর্তসাপেক্ষ নবায়ন"
            }
        },
        {
            "agenda_no": "০৬",
            "subject": "প্যান এম টেক লিমিটেড (Pan M Tech Ltd.) এর ন্যাশনওয়াইড আইএসপি লাইসেন্সের শেয়ার স্থানান্তর অনুমোদন প্রসঙ্গে।",
            "decision": "প্যান এম টেক লিমিটেড এর শেয়ার স্থানান্তর অনুমোদন করা হলো এবং বিলম্বিত আবেদনের জন্য ৫০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো।",
            "fine_amount": "৫০,০০০",
            "details": {
                "presentation_summary": "প্যান এম টেক লিমিটেড এর মালিকানা পরিবর্তন এবং আরজেএসসি (RJSC) নথি পর্যালোচনা করা হয়। বিধি লঙ্ঘনের কারণে জরিমানা সাপেক্ষে মালিকানা বদল মঞ্জুর করা হলো।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "আরোপিত জরিমানা", "পরিশোধের সময়সীমা"],
                        "rows": [
                            ["Pan M Tech Ltd.", "Nationwide ISP", "৫০,০০০/- টাকা", "৩০ দিন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "জরিমানা আরোপ"
            }
        },
        {
            "agenda_no": "০৭",
            "subject": "ডেল্টা সফটওয়্যার অ্যান্ড কমিউনিকেশন লিমিটেড (Delta Software and Communication Limited) এর শেয়ার স্থানান্তর অনুমোদন এবং বিলম্ব ফি মওকুফ সংক্রান্ত আবেদন পর্যালোচনা।",
            "decision": "ডেল্টা সফটওয়্যার অ্যান্ড কমিউনিকেশন লিমিটেড (Nationwide ISP) এর শেয়ার স্থানান্তর অনুমোদন করা হলো। বিলম্বিত আবেদনের প্রেক্ষিতে ৫০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো।",
            "fine_amount": "৫০,০০০",
            "details": {
                "presentation_summary": "ডেল্টা সফটওয়্যার এর শেয়ার স্থানান্তর নথিপত্র এবং পূর্ব অনুমতি ছাড়া মালিকানা পরিবর্তনের ব্যত্যয় পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "আরোপিত জরিমানা", "স্থিতি"],
                        "rows": [
                            ["Delta Software & Communication Ltd", "Nationwide ISP", "৫০,০০০/- টাকা", "জরিমানা পরিশোধের নির্দেশ"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "জরিমানা আরোপ"
            }
        },
        {
            "agenda_no": "০৮",
            "subject": "ফ্রান্স বাংলা আইটিসি টেকনোলজি প্রা. লিমিটেড (France Bangla ITC Technology Pvt. Ltd.) এর ভিটিএস (VTS) লাইসেন্সের সেবা মান ও কার্যক্রম পর্যালোচনা।",
            "decision": "প্রতিষ্ঠানটির সেবা মান বৃদ্ধির নির্দেশ প্রদান করা হলো এবং গ্রাহকদের সঠিক জিপিএস ট্র্যাকিং ও নিরাপত্তা সেবা নিশ্চিত করার জন্য কঠোর সতর্কবাণী জারি করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "ফ্রান্স বাংলা আইটিসি টেকনোলজি এর ভিটিএস লাইসেন্সের আওতায় প্রদত্ত সেবা এবং লাইসেন্সিং বিধিমালা পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "পরিদর্শন প্রতিবেদন", "সিদ্ধান্ত"],
                        "rows": [
                            ["VTS", "গ্রাহক সেবার মান উন্নত করা প্রয়োজন", "সতর্কীকরণ"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট।",
                "assigned_inspector": "এনফোর্সমেন্ট টিম",
                "case_status": "সতর্কীকরণ"
            }
        },
        {
            "agenda_no": "০৯",
            "subject": "মোবাইল ফোন অপারেটরদের সেবার মান (QoS) পরিবীক্ষণের লক্ষ্যে দেশব্যাপী পরিচালিত ড্রাইভ টেস্টের ফলাফল উপস্থাপন ও পর্যালোচনা।",
            "decision": "ড্রাইভ টেস্টে প্রাপ্ত অপারেটরদের কল ড্রপ, নেটওয়ার্ক কভারেজ ও দুর্বল সিগন্যাল সংক্রান্ত ব্যত্যয়ের জন্য গ্রামীণফোন, রবি, বাংলালিংক ও টেলিটককে সেবার মান উন্নত করার এবং কল ড্রপের ক্ষেত্রে গ্রাহকদের ক্ষতিপূরণ প্রদানের কঠোর নির্দেশ প্রদান করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "বিটিআরসির ড্রাইভ টেস্ট মনিটরিং টিমের পরিদর্শন ও ড্রাইভ টেস্টে প্রাপ্ত কল ড্রপ ও থ্রুপুট ডাটা উপস্থাপন করা হয়।",
                "tables": [
                    {
                        "headers": ["অপারেটর", "পরিদর্শনকৃত জেলা", "কল ড্রপ হার (KPI)", "সিদ্ধান্ত"],
                        "rows": [
                            ["জিপি, রবি, বাংলালিংক, টেলিটক", "সারাদেশের প্রধান সড়ক ও রেললাইন", "নির্ধারিত সীমার চেয়ে বেশি কল ড্রপ", "ক্ষতিপূরণ ও সেবা মানোন্নয়নের নির্দেশ"]
                        ]
                    }
                ],
                "implementation": "ইঞ্জিনিয়ারিং অ্যান্ড অপারেশন্স বিভাগ এবং ইএন্ডআই।",
                "assigned_inspector": "যৌথ ড্রাইভ টেস্ট টিম",
                "case_status": "কার্যক্রম চলমান"
            }
        },
        {
            "agenda_no": "১০",
            "subject": "গ্লোবাল ভয়েস টেলিকম লিমিটেড (Global Voice Telecom Ltd.) এর আইজিডব্লিউ (IGW) ব্যান্ডউইথ ব্যবহারের সীমা (Bandwidth Capping) নির্ধারণ প্রসঙ্গে।",
            "decision": "গ্লোবাল ভয়েস টেলিকম লিমিটেড এর বকেয়া রাজস্ব পরিশোধের অগ্রগতি সন্তোষজনক না হওয়ায় তাদের ব্যান্ডউইথ ক্যাপ বহাল রাখার এবং বকেয়া পরিশোধ সাপেক্ষে ক্যাপ আংশিক শিথিল করার সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "গ্লোবাল ভয়েস টেলিকম (IGW) এর বকেয়া ও রেভিনিউ শেয়ারিং সংক্রান্ত ডাটা এবং ক্যাপ মনিটরিং ডাটা পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "বকেয়া স্থিতি", "ক্যাপের পরিমাণ", "মন্তব্য"],
                        "rows": [
                            ["IGW", "বকেয়া রাজস্ব বিদ্যমান", "ব্যান্ডউইথ ক্যাপ সক্রিয়", "পরিশোধের ওপর ভিত্তি করে শিথিলযোগ্য"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট এবং অর্থ বিভাগ।",
                "assigned_inspector": "অর্থ ও স্পেকট্রাম বিভাগ",
                "case_status": "ক্যাপ বহাল"
            }
        },
        {
            "agenda_no": "১১",
            "subject": "আর্থ টেলিকমিউনিকেশন লিমিটেড (Earth Telecommunication Ltd.) এর আইআইজি লাইসেন্সের ব্যান্ডউইথ ক্যাপ এবং এনটিটিএন সংযোগ সংক্রান্ত বকেয়া পর্যালোচনা।",
            "decision": "আর্থ টেলিকমিউনিকেশনের বকেয়া রাজস্বের বিপরীতে ১৫ কোটি টাকার জরিমানা ও কিস্তি সুবিধা সংক্রান্ত রিভিউ আবেদন নাকচ করা হলো এবং বকেয়া কিস্তি পরিশোধের নির্দেশ প্রদান করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "আর্থ টেলিকমিউনিকেশন (IIG) এর বকেয়া এবং এনটিটিএন সংযোগ সেবা সংক্রান্ত বিরোধ ও রিভিউ আবেদন পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "বকেয়া রাজস্ব", "রিভিউ সিদ্ধান্ত", "মন্তব্য"],
                        "rows": [
                            ["IIG", "১৫,০০,০০,০০০/- টাকা", "রিভিউ আবেদন নাকচ", "বকেয়া দ্রুত পরিশোধের নির্দেশ"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড ও অর্থ বিভাগ",
                "case_status": "আবেদন নাকচ"
            }
        },
        {
            "agenda_no": "১২",
            "subject": "লেভেল-৩ ক্যারিয়ার লিমিটেড (Level-3 Carrier Ltd.) এর আইসিএক্স (ICX) ব্যান্ডউইথ ক্যাপ এবং রেভিনিউ শেয়ারিং বকেয়া পর্যালোচনা।",
            "decision": "লেভেল-৩ ক্যারিয়ার লিমিটেড (ICX) এর রেভিনিউ শেয়ারিং সংক্রান্ত বিরোধ ও ১৫ কোটি টাকার বকেয়া রিভিউ প্রক্রিয়া ত্বরান্বিত করার এবং বকেয়া আদায়ের কার্যক্রম জোরদার করার সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "লেভেল-৩ ক্যারিয়ার এর আইসিএক্স ও আইআইজি উভয় লাইসেন্সের বকেয়া রাজস্ব ও রেভিনিউ শেয়ারিং সংক্রান্ত জটিলতা এবং আদালতে চলমান কেইসসমূহের স্থিতি পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সের ধরণ", "বকেয়া/জরিমানা", "স্থিতি", "মন্তব্য"],
                        "rows": [
                            ["ICX", "১৫,০০,০০,০০০/- টাকা", "রিভিউ চলমান", "আইনি নিষ্পত্তি সাপেক্ষে কার্যক্রম"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "লিগ্যাল ও ইএন্ডআই",
                "case_status": "রিভিউ চলমান"
            }
        },
        {
            "agenda_no": "১৩",
            "subject": "এ২পি এসএমএস (A2P SMS) লাইসেন্সধারী প্রতিষ্ঠানসমূহের পারস্পরিক বিরোধ, রেভিনিউ শেয়ারিং ব্যত্যয় এবং আরজেএসসি (RJSC) এর সিডিউল-এক্স (Schedule X) সংক্রান্ত ব্যত্যয় পর্যালোচনা।",
            "decision": "এ২পি এসএমএস লাইসেন্সধারী প্রতিষ্ঠান (BIT & BYTE SOLUTION, TMSS ICT Ltd, Comjagat Technologies) এর লাইসেন্স নবায়ন শর্তসাপেক্ষে অনুমোদিত করা হলো এবং প্রত্যককে ৫০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করা হলো। উইন্টেল (Wintel) এবং রিভ/সফটওয়্যার শপ (REVE / Software Shop) এর Gross Revenue এবং আরজেএসসি এর Schedule X সংক্রান্ত নিয়ম ভঙ্গের জন্য সতর্কবাণী ও জরিমানা আরোপ করা হলো।",
            "fine_amount": "৫০,০০০",
            "details": {
                "presentation_summary": "এ২পি এসএমএস গেটওয়ে অপারেটরদের রেভিনিউ শেয়ারিং ডাটা এবং আরজেএসসি এর মালিকানা বদল ও সিডিউল-এক্স দাখিল সংক্রান্ত লাইসেন্স বিধির শর্তাবলী পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["অপারেটরের নাম", "ব্যত্যয়ের ধরণ", "আরোপিত জরিমানা", "মন্তব্য"],
                        "rows": [
                            ["BIT & BYTE SOLUTION", "রেভিনিউ শেয়ারিং ও মালিকানা তথ্য বিলম্ব", "৫০,০০০/- টাকা", "৩০ দিনের মধ্যে পরিশোধের নির্দেশ"],
                            ["TMSS ICT Ltd", "লাইসেন্স বিধি লঙ্ঘন", "৫০,০০০/- টাকা", "৩০ দিনের মধ্যে পরিশোধের নির্দেশ"],
                            ["Comjagat Technologies", "আরজেএসসি তথ্য বিলম্ব", "৫০,০০০/- টাকা", "৩০ দিনের মধ্যে পরিশোধের নির্দেশ"],
                            ["Wintel", "Schedule X ব্যত্যয়", "সতর্কীকরণ ও নিয়মিতকরণ", "আরজেএসসি তথ্য দাখিলের নির্দেশ"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও সিস্টেমস এন্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড ও এসএস বিভাগ",
                "case_status": "জরিমানা ও সতর্কীকরণ"
            }
        }
    ]
}

# 3.3 Meeting 307 (Riyad Hossain draft reports and new cases + other departments' agendas)
print("Injecting detailed 307th meeting data...")
meetings["307তম"] = {
    "meeting_number": "৩০৭তম",
    "meeting_date": "২৪ মে, ২০২৬",
    "agendas": [
        {
            "agenda_no": "০১",
            "subject": "ইন্টারনেট সেবাদাতা প্রতিষ্ঠানসমূহের (ISP) পারস্পরিক বিরোধ, নিয়মভঙ্গ ও শুনানি সংক্রান্ত সারসংক্ষেপ উপস্থাপন ও অনুমোদন প্রসঙ্গে।",
            "decision": "ইন্টারনেট সেবাদাতা প্রতিষ্ঠানসমূহের (ISP) পারস্পরিক বিরোধ, নিয়মভঙ্গ ও শুনানির প্রেক্ষিতে সার্কেল নেটওয়ার্ক বনাম ট্রায়াঙ্গেল সার্ভিসেস, এইচ কে নেট বনাম আবুল কালাম অনলাইন, এবং এস এম ইন্টারনেট বনাম ফাস্ট ইন্টারনেট এর সমঝোতা চুক্তি ও মুচলেকা অনুমোদন করা হলো এবং উভয় পক্ষকে শর্ত ভঙ্গের ক্ষেত্রে লাইসেন্স বাতিলের বিষয়ে সতর্ক করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "সার্কেল নেটওয়ার্ক, ট্রায়াঙ্গেল সার্ভিসেস, এইচ কে নেট, আবুল কালাম অনলাইন, এবং এস এম ইন্টারনেটের পারস্পরিক তার কাটা, এলাকা বিরোধ ও অন্যান্য ব্যত্যয়ের প্রেক্ষিতে অনুষ্ঠিত শুনানিসমূহের বিবরণ ও সমঝোতা চুক্তি কমিশনে উপস্থাপন করা হয়।",
                "tables": [
                    {
                        "headers": ["অভিযোগকারী ও অভিযুক্ত", "বিরোধের বিষয়", "সিদ্ধান্ত ও অগ্রগতি"],
                        "rows": [
                            ["সার্কেল নেটওয়ার্ক বনাম আই নক্স নেট", "আদাবর এলাকায় তার কাটা ও সাবোটাজ", "আই নক্স দুঃখ প্রকাশ করেছে; এলাকা সীমানা থাকবে না, গ্রাহক পছন্দমত সেবা নেবে। সতর্ক করা হলো।"],
                            ["সার্কেল নেটওয়ার্ক বনাম ট্রায়াঙ্গেল সার্ভিসেস", "রংপুরে তার কাটা ও হুমকি প্রদান", "ট্রায়াঙ্গেল দুঃখ প্রকাশ করেছে; দোষী কর্মীদের বিরুদ্ধে ব্যবস্থা গ্রহণের নির্দেশ দেওয়া হয়েছে।"],
                            ["এইচ কে নেট বনাম আবুল কালাম অনলাইন", "রূপগঞ্জে ফাইবার ক্যাবল কাটা ও এলাকা বিরোধ", "উভয় পক্ষ সমঝোতা স্বাক্ষর করেছে এবং ক্ষতিপূরণ প্রদানপূর্বক অভিযোগ প্রত্যাহার করেছে।"],
                            ["এস এম ইন্টারনেট বনাম ফাস্ট ইন্টারনেট", "চাঁদপুরে অবৈধ রিসেলার ও এলাকা লঙ্ঘন", "উভয় পক্ষ চাঁদপুর সদর ও মতলব সীমানা মেনে চলার মুচলেকা প্রদান করেছে।"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "E&I শুনানী বোর্ড",
                "case_status": "নিষ্পত্তি ও সতর্কীকরণ"
            }
        },
        {
            "agenda_no": "০২",
            "subject": "Getco Telecommunication Ltd (ICX), Donia Cable Vision (ISP), এবং Electro Soft (ISP) সহ ১৯টি ইন্টারনেট সেবাদাতা ও আইসিএক্স প্রতিষ্ঠানের পরিদর্শন ও তদন্ত প্রতিবেদন উপস্থাপন প্রসঙ্গে।",
            "decision": "ব্যত্যয়সমূহের জন্য Getco Telecommunication, Electro Soft, এবং Dhaka Fiber Net এর কারণ দর্শানোর নোটিশের জবাব পর্যালোচনা করে পরবর্তী পদক্ষেপ গ্রহণ এবং Donia Cable Vision এর পরিদর্শন প্রতিবেদন ৩০৬তম কমিশন সভার সিদ্ধান্তের প্রেক্ষিতে সমন্বয় করার নির্দেশ প্রদান করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "উপ-সহকারী পরিচালক রিয়াদ হোসেনের পরিদর্শন দল কর্তৃক দাখিলকৃত প্রতিবেদনের আলোকে Getco Telecommunication (ICX), Donia Cable Vision (ISP), M/S Blueberry & Rapid Network (ISP), ADN Telecom Ltd (ISP), Always On Network Bangladesh, System Solution & Technologies, Dhaka Fiber Net, Excel Technologies Ltd, Sky View Online, এবং Electro Soft এর পরিদর্শন ফাইল ও ও শো-কজের জবাবসমূহ পর্যালোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "ব্যত্যয়ের ধরণ", "বর্তমান অবস্থা"],
                        "rows": [
                            ["Getco Telecommunication (ICX)", "পরিদর্শন প্রতিবেদন ব্যত্যয়", "কারণ দর্শানোর নোটিশ জারি করা হয়েছে, জবাব পাওয়া যায়নি।"],
                            ["M/S Blueberry & Rapid Network (ISP)", "নিয়ম বহির্ভূত কর্মকান্ড", "শো-কজ জবাব এসেছে, ৩০৭তম কমিশন সভায় চূড়ান্ত উপস্থাপনের যোগ্য।"],
                            ["ADN Telecom Ltd (ISP)", "লাইসেন্স বিধি লঙ্ঘন", "শো-কজ জবাব এসেছে, চূড়ান্ত সিদ্ধান্ত গ্রহণের যোগ্য।"],
                            ["Dhaka Fiber Net Ltd", "লাইসেন্স ব্যত্যয়", "শো-কজ জবাব এসেছে, ৩০৭তম commission সভায় উপস্থাপিত।"],
                            ["Electro Soft (ISP)", "পরিদর্শন ব্যত্যয়", "শো-কজ জবাব এসেছে, জবাব পর্যালোচনার সিদ্ধান্ত।"],
                            ["Donia Cable Vission (ISP)", "পরিদর্শন প্রতিবেদন", "৩০৬তম কমিশন সভার সিদ্ধান্তের সাথে সমন্বয়যোগ্য।"],
                            ["Always On Network (ISP)", "শর্ত লঙ্ঘন", "৩০৬তম কমিশন সভায় উপস্থাপিত হয়েছিল।"],
                            ["System Solution & Technologies", "পরিদর্শন ব্যত্যয়", "৩০৬তম কমিশন সভায় উপস্থাপিত হয়েছিল।"],
                            ["Excel Technologies Ltd", "ভেন্ডর লাইসেন্স ব্যত্যয়", "প্রতিবেদন পর্যালোচনার সিদ্ধান্ত।"],
                            ["VoIP খাগড়াছড়ি", "অবৈধ ভিওআইপি", "তদন্ত প্রতিবেদন দাখিল।"],
                            ["Sky View Online Ltd (ISP)", "সংযোগ ও বিলিং ব্যত্যয়", "Donia Cable Vission এর ফাইলের সাথে সমন্বয়।"],
                            ["Mango Teleservices & BTCL", "৬৫(৪) তদন্ত প্রতিবেদন", "তদন্ত প্রতিবেদন পর্যালোচনাধীন।"],
                            ["BD Link & Novocom Limited", "৬৫(৪) তদন্ত প্রতিবেদন", "তদন্ত প্রতিবেদন পর্যালোচনাধীন।"],
                            ["Udayon Online Limited", "পরিদর্শন প্রতিবেদন", "৩০৬তম কমিশন সভায় উপস্থাপিত হয়েছিল, সিদ্ধান্ত চলমান।"],
                            ["Digital One Broadband Service", "পরিদর্শন প্রতিবেদন", "প্রতিবেদন দাখিল সম্পন্ন।"],
                            ["VoIP (মোহাম্মদপুর)", "অবৈধ ভিওআইপি", "অভিযান ও মামলা দায়ের প্রতিবেদন।"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট।",
                "assigned_inspector": "Riyad Hossain (AD, E&I)",
                "case_status": "জবাব পর্যালোচনা"
            }
        },
        {
            "agenda_no": "০৩",
            "subject": "মোবাইলে জব্দকৃত সিম ও অবৈধ ভিওআইপি (VoIP) অপারেশন সংক্রান্ত থানা মামলা ও অগ্রগতি আপডেট।",
            "decision": "NTMC এর সহযোগিতায় কুমিল্লা কোতয়ালী এবং চট্টগ্রামের হালিশহরে পরিচালিত সফল ভিওআইপি অভিযানে জব্দকৃত হাজার হাজার সিম এবং পলাতক আসামীদের বিরুদ্ধে রুজুকৃত মামলার তদন্ত তদারকি জোরদার করার জন্য লিগ্যাল বিভাগকে নির্দেশ প্রদান করা হলো। চট্টগ্রামে হালিশহরের অভিযান ও মামলা দায়েরের তারিখ ১৯/০৫/২০২৬ খ্রিঃ চূড়ান্তভাবে নির্ধারণ করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "কুমিল্লা কোতয়ালী এবং চট্টগ্রামের হালিশহর এলাকায় এনটিএমসি-এর সহযোগিতায় পরিচালিত সরেজমিন ভিওআইপি অভিযানে বিপুল পরিমাণ অবৈধ ভিওআইপি আলামত ও সিম উদ্ধার করা হয় এবং বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ আইন, ২০০১ এর ধারা ৩৫(২)/৫৫(৭)/৫৭(৩)/৭৪ অনুযায়ী থানায় নিয়মিত মামলা দায়ের করা হয়।",
                "tables": [
                    {
                        "headers": ["অভিযানের স্থান", "অভিযানের তারিখ", "জব্দকৃত সিম ও সরঞ্জাম", "আইনি ধারা", "মামলার বিবরণ ও আসামী"],
                        "rows": [
                            ["কোতয়ালী, কুমিল্লা", "২০-২৩ এপ্রিল, ২০২৬", "অবৈধ ভিওআইপি সরঞ্জাম ও প্রায় ৮,০০০ সিম (গ্রামীণফোন, রবি, বাংলালিংক, Teletalk)", "৩৫(২)/৫৫(৭)/৫৭(৩)/৭৪ ধারা", "কোতয়ালী মডেল থানা মামলা (তারিখ ২২/০৪/২০২৬)। পলাতক আসামী: মোঃ আবু সাঈদ চৌধুরী (৪৫)।"],
                            ["হালিশহর, চট্টগ্রাম", "১৭-২০ মে, ২০২৬", "টেলিটক ও রবির অবৈধ সিম ও ভিওআইপি সরঞ্জাম", "৩৫(২)/৫৭(৩)/৭৪ ধারা", "হালিশহর থানা মামলা (তারিখ ১৯/০৫/২০২৬)। পলাতক আসামী: জিয়াবুল হক (৪২)।"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও লিগ্যাল এন্ড লাইসেন্সিং (এলএল) বিভাগ।",
                "assigned_inspector": "E&I VoIP Team & Law Enforcement",
                "case_status": "মামলা চলমান"
            }
        },
        {
            "agenda_no": "০৪",
            "subject": "Direct-to-Cell (D2C) Proof of Concept (PoC) পরিচালনার উদ্দেশ্যে Banglalink কে Mobile Network Code (MNC) ও PLMN বরাদ্দ প্রদান প্রসঙ্গে।",
            "decision": "মেসার্স মাহের ইন্টারন্যাশনাল-এর আবেদনের প্রেক্ষিতে পার্টনার মোবাইল অপারেটর বাংলালিংক ডিজিটাল কমিউনিকেশনস লিমিটেড-কে Direct-to-Cell (D2C) Proof of Concept (PoC) টেস্ট পরিচালনার উদ্দেশ্যে Mobile Network Code (MNC) '08' (PLMN 470 08) সাময়িকভাবে ব্যবহারের অনুমোদন প্রদান করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "নতুন প্রযুক্তির ট্রায়াল হিসেবে Direct-to-Cell (D2C) PoC পরীক্ষার জন্য আবেদন করা হয়। স্পেকট্রাম বিভাগের কারিগরি কমিটি পর্যালোচনা করে বাংলালিংক-কে MNC 08 (PLMN 470 08) সাময়িক বরাদ্দের জন্য commission সভায় প্রস্তাব উত্থাপন করে এবং তা অনুমোদিত হয়।",
                "tables": [
                    {
                        "headers": ["আবেদনকারী", "পার্টনার অপারেটর", "প্রযুক্তির ধরণ", "বরাদ্দকৃত কোড (PLMN)", "স্ট্যাটাস"],
                        "rows": [
                            ["Maher International", "Banglalink Digital Communications Ltd", "Direct-to-Cell (D2C) PoC", "PLMN 470 08 (MNC: 08)", "সাময়িক অনুমোদন"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ ও সিস্টেমস এন্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Spectrum Engineering Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০৫",
            "subject": "\"Honor Choice\" ব্র্যান্ড নামযুক্ত ফিচার ফোন বাংলাদেশে আমদানির উদ্দেশ্যে টাইপ অ্যাপ্রুভাল এবং এনওসি (NOC) প্রদান প্রসঙ্গে।",
            "decision": "Honor Information Technology Co. Limited (Hong Kong) এর মালিকানাধীন \"Honor Choice\" ব্র্যান্ডের ফিচার ফোনসমূহ আন্তর্জাতিক ব্র্যান্ড নাম বিবেচনা করে টাইপ অ্যাপ্রুভাল এবং বাংলাদেশে আমদানির জন্য ছাড়পত্র (NOC) প্রদানের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "অনার চয়েস (Honor Choice) ফিচার ফোন আমদানির উদ্দেশ্যে লোকাল এজেন্ট এক্সট্রিম (Xtreme) এর আবেদনের প্রেক্ষিতে আন্তর্জাতিক ব্র্যান্ড হিসেবে ডিভাইসটির টাইপ অ্যাপ্রুভাল ও আমদানির অনাপত্তি সনদ (NOC) প্রদানের প্রস্তাব বিবেচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["ব্র্যান্ড নাম", "মালিকানাধীন প্রতিষ্ঠান", "আবেদনকারী (লোকাল এজেন্ট)", "ডিভাইসের ধরণ", "সিদ্ধান্ত"],
                        "rows": [
                            ["Honor Choice", "Honor Information Technology Co. Ltd (HK)", "Xtreme Distribution", "ফিচার ফোন (Feature Phone)", "টাইপ অ্যাপ্রুভাল ও NOC অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ।",
                "assigned_inspector": "Spectrum Type Approval Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০৬",
            "subject": "\"OUKITEL\" ব্র্যান্ডের ফিচার ফোন আমদানির উদ্দেশ্যে টাইপ অ্যাপ্রুভাল এবং আমদানির ছাড়পত্র (NOC) প্রদান প্রসঙ্গে।",
            "decision": "Shenzhen YunJi Intelligent Technology Co., Ltd. এর মালিকানাধীন \"OUKITEL\" ব্র্যান্ডের ফিচার ফোনসমূহকে টাইপ অ্যাপ্রুভাল এবং আমদানির উদ্দেশ্যে ছাড়পত্র (NOC) প্রদানের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "ওউকিটেল (OUKITEL) ব্র্যান্ডের ফিচার ফোন আমদানির জন্য আবেদন করা হয়। স্পেকট্রাম বিভাগ আন্তর্জাতিক ব্র্যান্ড রেজিস্ট্রেশন ও কারিগরি মান বিবেচনা করে টাইপ অ্যাপ্রুভালের সুপারিশ করে।",
                "tables": [
                    {
                        "headers": ["ব্র্যান্ড নাম", "প্রস্তুতকারক প্রতিষ্ঠান", "ডিভাইসের ধরণ", "সিদ্ধান্ত"],
                        "rows": [
                            ["OUKITEL", "Shenzhen YunJi Intelligent Technology Co., Ltd.", "ফিচার ফোন", "টাইপ অ্যাপ্রুভাল ও NOC অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ।",
                "assigned_inspector": "Spectrum Type Approval Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০৭",
            "subject": "নন-মোবাইল হ্যান্ডসেট ক্যাটাগরিতে Proton Router, Walkie-talkie এবং Vehicle Tracking Device আমদানির লক্ষ্যে টাইপ অ্যাপ্রুভাল ও ভেন্ডর তালিকাভুক্তি অনুমোদন।",
            "decision": "Telecommunication Service Related Equipment Non-Mobile Handset ক্যাটাগরিতে প্রোটন রাউটার (Proton Router), ওয়াকি-টকি (Walkie-talkie) এবং ভেহিকল ট্র্যাকিং ডিভাইস (Vehicle Tracking Device) আমদানির লক্ষ্যে প্রয়োজনীয় টাইপ অ্যাপ্রুভাল ও ভেন্ডর তালিকাভুক্তির প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "টেলিকম সেবা সংশ্লিষ্ট বিভিন্ন নন-মোবাইল সরঞ্জাম যেমন রাউটার, ওয়াকি-টকি ও ট্র্যাকিং ডিভাইস আমদানির জন্য বিভিন্ন আমদানিকারক ও ভেন্ডর তালিকাভুক্তির আবেদন স্পেকট্রাম বিভাগের কারিগরি কমিটি পর্যালোচনা করে অনুমোদন প্রদান করে।",
                "tables": [
                    {
                        "headers": ["সরঞ্জামের ধরণ", "ব্র্যান্ড/মডেল", "অনুমোদনের ধরণ", "অবস্থা"],
                        "rows": [
                            ["Proton Router", "Proton Series", "Type Approval & Enlistment", "অনুমোদিত"],
                            ["Walkie-talkie", "Various Models", "Type Approval & Enlistment", "অনুমোদিত"],
                            ["Vehicle Tracking Device", "Various Models", "Type Approval & Enlistment", "অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ।",
                "assigned_inspector": "Spectrum Equipment Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০৮",
            "subject": "EGSM (৮৮০-৮৯০ মেগাহার্টজ / ৯২৫-৯৩৫ মেগাহার্টজ) স্পেকট্রাম ব্যান্ডের পুনর্গঠন ও অকশন পলিসি নির্ধারণ প্রসঙ্গে।",
            "decision": "EGSM (900 MHz) ব্যান্ডের স্পেকট্রামের কার্যকারিতা বাড়াতে অকশন পলিসি এবং অপারেটরদের ব্যান্ড রি-অ্যালোকেশনের প্রস্তাব চূড়ান্ত অনুমোদন দেওয়া হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "EGSM ব্যান্ডের স্পেকট্রাম পুনর্গঠনের মাধ্যমে তরঙ্গের সর্বোচ্চ ব্যবহার নিশ্চিত করা এবং অকশনের জন্য স্পেকট্রাম বিভাগ প্রস্তাব পেশ করে।",
                "tables": [
                    {
                        "headers": ["ফ্রিকোয়েন্সি ব্যান্ড", "স্পেকট্রাম রেঞ্জ", "পুনর্গঠনের লক্ষ্য", "অবস্থা"],
                        "rows": [
                            ["EGSM 900 MHz", "880-890 MHz & 925-935 MHz", "তরঙ্গ পুনর্গঠন ও অকশন পলিসি নির্ধারণ", "অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ।",
                "assigned_inspector": "Spectrum Management Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "০৯",
            "subject": "চট্টগ্রাম বন্দর কর্তৃপক্ষ (CPA) এর 'Port Modernization and Digital Connectivity Enhancement' প্রকল্পের আওতায় ৩.৫ গিগাহার্জ ব্যান্ডে ডেডিকেটেড 5G LTE PoC ট্রায়াল পরিচালনা।",
            "decision": "চট্টগ্রাম বন্দর কর্তৃপক্ষকে ৩.৫ গিগাহার্জ ব্যান্ডে (3300-3340 MHz) ৩০ মেগাহার্টজ স্পেকট্রাম ব্যবহারের মাধ্যমে ডেডিকেটেড 5G LTE Stand-Alone PoC পরীক্ষার অনুমোদন প্রদান করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "চট্টগ্রাম বন্দরের আধুনিকীকরণ ও স্মার্ট পোর্ট কার্যক্রম ভাগের জন্য ৩.৫ গিগাহার্জ ব্যান্ডে ডেডিকেটেড নেটওয়ার্ক স্থাপনের ট্রায়াল অনুমোদন করা হয়।",
                "tables": [
                    {
                        "headers": ["আবেদনকারী প্রতিষ্ঠান", "প্রকল্পের নাম", "বরাদ্দকৃত ফ্রিকোয়েন্সি", "ব্যান্ডউইথ", "সিদ্ধান্ত"],
                        "rows": [
                            ["Chittagong Port Authority (CPA)", "Port Modernization & Digital Connectivity", "3.5 GHz Band (3300-3340 MHz)", "30 MHz (Stand-Alone)", "PoC ট্রায়াল অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "স্পেকট্রাম বিভাগ ও সিস্টেমস এন্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Spectrum Engineering Section",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১০",
            "subject": "আইপিটিএসপি (IPTSP) অপারেটরদের লাইসেন্স নবায়ন, বকেয়া রাজস্ব আদায় এবং সিপিলএ (CPLA) আদালতের আদেশ সংক্রান্ত পর্যালোচনা।",
            "decision": "সিভিল পিটিশন ফর লিভ টু আপিল (CPLA) এর প্রেক্ষিতে আইপিটিএসপি অপারেটরদের বকেয়া এবং নবায়নের বিষয়ে আদালতের নির্দেশনা অনুযায়ী ব্যবস্থা গ্রহণের প্রস্তাব অনুমোদিত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "আইপিটিএসপি অপারেটরদের লাইসেন্স নবায়ন ফি এবং বকেয়া রাজস্ব আদায় সম্পর্কিত আইনি জটিলতা নিরসনে সুপ্রিম কোর্টের সিপিলএ আদেশের বাস্তবায়ন নিয়ে আলোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["আইনি বিষয়", "আদালতের ফোরাম", "সংশ্লিষ্ট লাইসেন্স", "অবস্থা"],
                        "rows": [
                            ["Civil Petition for Leave to Appeal (CPLA)", "সুপ্রিম কোর্ট", "IPTSP License", "আদালতের আদেশ অনুযায়ী ব্যবস্থা গ্রহণের সিদ্ধান্ত"]
                        ]
                    }
                ],
                "implementation": "লিগ্যাল বিভাগ ও সিস্টেমস এন্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Legal & Systems Section",
                "case_status": "আইনি নির্দেশনা মোতাবেক চলমান"
            }
        },
        {
            "agenda_no": "১১",
            "subject": "মেসার্স কম্বাইন্ড সফট (Combined Soft) এর স্বত্বাধিকারীর মৃত্যুজনিত কারণে উত্তরাধিকার সূত্রে লাইসেন্স স্থানান্তর অনুমোদন।",
            "decision": "The Majority Act, 1875 এবং উত্তরাধিকার সনদের (Succession Certificate) ভিত্তিতে 'কম্বাইন্ড সফট' (Combined Soft) এর মালিকানা ও লাইসেন্স স্থানান্তরের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "কম্বাইন্ড সফট এর স্বত্বাধিকারীর মৃত্যুর পর তার উত্তরাধিকারীদের আবেদনের প্রেক্ষিতে দ্য মেজরিটি অ্যাক্ট ১৮৭৫ এবং আদালতের সাকসেশন সার্টিফিকেট অনুযায়ী লাইসেন্সটি স্থানান্তরের প্রস্তাব উত্থাপন করা হয়।",
                "tables": [
                    {
                        "headers": ["লাইসেন্সধারী প্রতিষ্ঠান", "পূর্ববর্তী স্বত্বাধিকারী", "আবেদনের ভিত্তি", "সিদ্ধান্ত"],
                        "rows": [
                            ["Combined Soft", "প্রয়াত স্বত্বাধিকারী", "Succession Certificate & Majority Act 1875", "উত্তরাধিকারীদের নামে লাইসেন্স স্থানান্তর অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ ও আইনি বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১২",
            "subject": "স্টার টেক ভিটিএস (Star Tech VTS) এর ভেহিকল ট্র্যাকিং সার্ভিস লাইসেন্সের নবায়ন ও কার্যক্রম পর্যালোচনা।",
            "decision": "স্টার টেক ভিটিএস (Star Tech VTS) এর লাইসেন্স শর্তাবলী পরিপালন এবং বকেয়া ফি পরিশোধ সাপেক্ষে লাইসেন্স নবায়নের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "স্টার টেক ভিটিএস এর লাইসেন্স নবায়নের আবেদন ও তাদের পূর্ববর্তী কার্যক্রম সিস্টেমস এন্ড সার্ভিসেস বিভাগ পর্যালোচনা করে ফি আদায় সাপেক্ষে নবায়নের সুপারিশ করে।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠানের নাম", "লাইসেন্সের ধরণ", "নবায়নের শর্ত", "অবস্থা"],
                        "rows": [
                            ["Star Tech VTS", "Vehicle Tracking Service (VTS)", "বকেয়া ফি ও জরিমানা পরিশোধ", "নবায়ন অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৩",
            "subject": "Frontier Towers Bangladesh Ltd (টাওয়ার শেয়ারিং লাইসেন্সধারী) এর শেয়ার হোল্ডিং কাঠামো পরিবর্তন এবং লাইসেন্সের শর্ত পূরণে রোল-আউট সময়সীমা বৃদ্ধি প্রসঙ্গে।",
            "decision": "Frontier Towers Bangladesh Ltd. (সাবেক AB HighTech Consortium Ltd.) এর শেয়ার কাঠামো পরিবর্তনপূর্বক নতুন বিনিয়োগকারী সমন্বিত শেয়ার হোল্ডিং (Pinnacle BD Holding: 46.8%, Integrated Services: 5.2%, Luminous Equity: 16%, Aroosa Janashakti: 16%, Dunhill Services: 16%) অনুমোদন করা হলো এবং লাইসেন্সের রোল-আউট সময়সীমা বাড়ানোর আবেদন মঞ্জুর করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "টাওয়ার শেয়ারিং লাইসেন্সের শর্তাবলী পরিপালন এবং রোল-আউট লক্ষ্যমাত্রা পূরণে গতিশীলতা আনতে শেয়ার হোল্ডিং কাঠামো পরিবর্তন এবং সময় বাড়ানোর আবেদন জমা দেওয়া হয়। সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ আবেদনটি বিশ্লেষণ করে সুপারিশ করে।",
                "tables": [
                    {
                        "headers": ["শেয়ারহোল্ডার নাম", "পূর্ববর্তী অংশীদারিত্ব (%)", "নতুন অংশীদারিত্ব (%)", "শেয়ার টাইপ / ক্লাস"],
                        "rows": [
                            ["Pinnacle BD Holding Pte. Ltd.", "৬৩.০৪%", "৪৬.৮%", "Class B/C Shares"],
                            ["Integrated Services Limited", "৬.৯৬%", "৫.২%", "Class B/C Shares"],
                            ["Luminous Equity / Management", "১০.০০%", "১৬.০%", "Class A/C Shares"],
                            ["Aroosa Janashakti Limited", "১০.০০%", "১৬.০%", "Class A/C Shares"],
                            ["Dunhill Services Limited", "১০.০০%", "১৬.০%", "Class A/C Shares"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ ও আইনি বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৪",
            "subject": "ইন্টারন্যাশনাল টেরিস্ট্রিয়াল ক্যাবল (ITC) ওআইনি শর্ত পরিপালন নভোকম লিমিটেড (NovoCom Ltd.) এর লাইসেন্স শর্ত পরিপালন ও রাজস্ব বকেয়া সংক্রান্ত আবেদন পর্যালোচনা।",
            "decision": "নভোকম লিমিটেডের আইটিসি লাইসেন্সের শর্ত ভঙ্গ ও বিলম্বিত ফি পরিশোধের বিষয়ে কঠোর সতর্কবার্তা প্রদানপূর্বক বকেয়া আদায়ের রূপরেখা অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "নভোকম লিমিটেডের আইটিসি লাইসেন্স বকেয়া রাজস্ব এবং সরকারের পাওনা আদায়ের উদ্দেশ্যে গঠিত তদন্ত দলের সুপারিশ ও আইনি শর্ত পরিপালন নিয়ে আলোচনা করা হয়।",
                "tables": [
                    {
                        "headers": ["অপারেটর", "লাইসেন্স ধরণ", "রাজস্ব বকেয়া অবস্থা", "সিদ্ধান্ত"],
                        "rows": [
                            ["NovoCom Limited", "ITC", "রাজস্ব বকেয়া রয়েছে", "ফি পরিশোধ ও বকেয়া আদায়ের সময়সূচী নির্ধারণ"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ ও অর্থ বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "তদন্ত রিপোর্ট সাপেক্ষে সিদ্ধান্ত"
            }
        },
        {
            "agenda_no": "১৫",
            "subject": "ন্যাশনওয়াইড টেলিকমিউনিকেশন ট্রান্সমিশন নেটওয়ার্ক (NTTN) অপারেটরদের অপটিক্যাল ফাইবার নেটওয়ার্ক সম্প্রসারণ ও কমপ্লায়েন্স পর্যালোচনা।",
            "decision": "সামিট কমিউনিকেশনস, ফাইবার অ্যাট হোম, বাহন লিমিটেড, পিজিসিবি, এবং বাংলাদেশ রেলওয়ে সহ এনটিটিএন অপারেটরদের পারস্পরিক সংযোগ স্থাপনে সহাবস্থানের রূপরেখা এবং নেটওয়ার্ক সম্প্রসারণের লাইসেন্স শর্তাবলী পরিপালনের রূপরেখা অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "এনটিটিএন লাইসেন্সধারী প্রতিষ্ঠানসমূহের দেশব্যাপী অপটিক্যাল ফাইবার নেটওয়ার্ক সম্প্রসারণ, সরকারের রেলওয়ে ও বিদ্যুৎ গ্রিড অবকাঠামো ব্যবহার এবং লাইসেন্সিং গাইডলাইনের শর্ত পরিপালন সংক্রান্ত বিস্তারিত উপস্থাপন।",
                "tables": [
                    {
                        "headers": ["এনটিটিএন অপারেটর", "অবকাঠামো ধরণ", "সম্প্রসারণের স্থিতি", "সিদ্ধান্ত"],
                        "rows": [
                            ["Summit Communications Ltd.", "ওভারহেড ও আন্ডারগ্রাউন্ড অপটিক ফাইবার", "দেশব্যাপী সম্প্রসারণ", "লাইসেন্স বিধি পরিপালন জোরদার"],
                            ["Fiber@Home Ltd.", "অপটিক ফাইবার নেটওয়ার্ক", "দেশব্যাপী সম্প্রসারণ", "সহাবস্থান ও শেয়ারিং জোরদার"],
                            ["Bahon / PGCB / Railway", "বিদ্যুৎ গ্রিড ও রেলওয়ে অপটিকাল ফাইবার", "বিশেষ অবকাঠামো ব্যবহার", "রাজস্ব অংশীদারিত্ব পরিপালন"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৬",
            "subject": "Cel Telecom Ltd., Novotel Ltd., এবং Mir Telecom Ltd. (IGW অপারেটরসমূহ) এর ট্রানজিশনাল পিরিয়ডে অ্যানুয়াল লাইসেন্স ফি কিস্তিতে পরিশোধের আবেদন পর্যালোচনা।",
            "decision": "Cel Telecom, Novotel, এবং Mir Telecom এর বার্ষিক লাইসেন্স ফি কিস্তিতে পরিশোধের বিশেষ সুবিধা অনুমোদন করা হলো (নভোটেলের জন্য ২৫ লক্ষ টাকা প্রতি কোয়ার্টার এবং মির টেলিকমের জন্য ১২টি সমান মাসিক কিস্তি)।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "নতুন টেলিকমিউনিকেশন নেটওয়ার্ক ও লাইসেন্সিং পলিসি ২০২৫ এবং আইসিএসপি (ICSP) লাইসেন্সিং চালুর প্রেক্ষিতে আইজিডব্লিউ অপারেটরদের অন্তর্বর্তীকালীন সময়ে লাইসেন্স ফি কিস্তিতে দেওয়ার অনুমোদন দেওয়া হয়।",
                "tables": [
                    {
                        "headers": ["প্রতিষ্ঠান", "অনুমোদিত কিস্তির সুবিধা", "পলিসি ভিত্তি", "অবস্থা"],
                        "rows": [
                            ["NovoTel Ltd.", "ত্রৈমাসিক কিস্তি (২৫ লক্ষ টাকা প্রতি কোয়ার্টার)", "ICSP লাইসেন্সিং পলিসি ২০২৫", "অনুমোদিত"],
                            ["Mir Telecom Ltd.", "১২টি সমান মাসিক কিস্তিতে পরিশোধ", "ICSP লাইসেন্সিং পলিসি ২০২৫", "অনুমোদিত"],
                            ["Cel Telecom Ltd.", "শেয়ারহোল্ডিং পরিবর্তনের প্রেক্ষিতে কিস্তি সুবিধা পর্যালোচনা", "ICSP লাইসেন্সিং পলিসি ২০২৫", "শর্তসাপেক্ষ অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "অর্থ, হিসাব ও রাজস্ব বিভাগ এবং ইএন্ডআই ডিরেক্টরেট।",
                "assigned_inspector": "Finance and Revenue Division",
                "case_status": "কিস্তি সুবিধা অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৭",
            "subject": "আলিজা এন্টারপ্রাইজ লিমিটেড (Aliza Enterprise Limited) এর লাইসেন্স ফি ও বকেয়া জরিমানা আদায় প্রসঙ্গে।",
            "decision": "আলিজা এন্টারপ্রাইজ লিমিটেডের বকেয়া প্রশাসনিক জরিমানা ও লাইসেন্স ফি আদায়ের উদ্দেশ্যে ডিমান্ড লেটার জারির সিদ্ধান্ত গৃহীত হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "আলিজা এন্টারপ্রাইজ লিমিটেডের বকেয়া জরিমানা ও লাইসেন্স ফি পরিশোধে ব্যর্থতার প্রেক্ষিতে অর্থ বিভাগ ও আইন বিভাগের সমন্বিত প্রতিবেদনে পাওনা আদায়ের জন্য তাগাদা প্রদানের সিদ্ধান্ত নেওয়া হয়।",
                "tables": [
                    {
                        "headers": ["অপারেটরের নাম", "বকেয়া ধরণ", "বকেয়া টাকার পরিমাণ", "সিদ্ধান্ত"],
                        "rows": [
                            ["Aliza Enterprise Limited", "লাইসেন্স ফি ও জরিমানা", "নির্দিষ্ট পরিমাণ বকেয়া", "ডিমান্ড লেটার জারির নির্দেশ"]
                        ]
                    }
                ],
                "implementation": "অর্থ, হিসাব ও রাজস্ব বিভাগ ও লিগ্যাল বিভাগ।",
                "assigned_inspector": "Finance & Revenue Division",
                "case_status": "ডিমান্ড লেটার ও পাওনা আদায়"
            }
        },
        {
            "agenda_no": "১৮",
            "subject": "ক্লাউডকম কনসালটেন্সি (Cloudcom Consultancy DWC-LLC) কে A2P SMS গেটওয়ে ও TVAS ব্র্যান্ড রেজিস্ট্রেশন অনুমোদন প্রসঙ্গে।",
            "decision": "ক্লাউডকম কনসালটেন্সিকে A2P SMS এগ্রিগেটর ও TVAS রেজিস্ট্রেশন প্রদানের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "ক্লাউডকম কনসালটেন্সি DWC-LLC নামক বিদেশী তথ্যপ্রযুক্তি প্রতিষ্ঠানের আবেদনের প্রেক্ষিতে বাংলাদেশে A2P SMS গেটওয়ে ট্রাফিক পরিচালনা এবং সংশ্লিষ্ট TVAS কার্যক্রমের অনুমোদন দেওয়া হলো।",
                "tables": [
                    {
                        "headers": ["আবেদনকারী", "সেবার ধরণ", "ব্র্যান্ড/রেজিস্ট্রেশন", "সিদ্ধান্ত"],
                        "rows": [
                            ["Cloudcom Consultancy DWC-LLC", "A2P SMS Aggregator & TVAS", "Cloudcom Gateway", "রেজিস্ট্রেশন ও ব্র্যান্ড অনুমোদন"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "১৯",
            "subject": "Meteorological And Related Services (MARS) Limited কে দেশের আবহাওয়া ও দুর্যোগ সম্পর্কিত পূর্বাভাসের বিশেষ সেবার জন্য শর্ট কোড বরাদ্দ প্রদান প্রসঙ্গে।",
            "decision": "Short Code Allocation Procedure, 2010 অনুযায়ী Meteorological And Related Services (MARS) Limited-কে ৫,০০,০০০/- (পাঁচ লক্ষ) টাকা ফিস পরিশোধ সাপেক্ষে আবহাওয়া সেবা প্রদানের জন্য ডিমান্ড লেটার ইস্যু এবং বিশেষ শর্ট কোড বরাদ্দ প্রদানের অনুমোদন দেওয়া হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "জনসাধারণের নিকট সহজে আবহাওয়ার তথ্য ও দুর্যোগের সতর্কতা বার্তা প্রদানের লক্ষ্যে আবহাওয়া অধিদপ্তর অনুমোদিত MARS Limited এর আবেদনের প্রেক্ষিতে এই বিশেষ শর্ট কোড বরাদ্দ প্রদানের সিদ্ধান্ত গ্রহণ করা হয়।",
                "tables": [
                    {
                        "headers": ["আবেদনকারী প্রতিষ্ঠান", "সেবার ধরণ", "বরাদ্দকৃত পলিসি", "প্রযোজ্য ফিস", "অবস্থা"],
                        "rows": [
                            ["Meteorological & Related Services (MARS) Ltd", "আবহাওয়া পূর্বাভাস ও দুর্যোগ সতর্কতা বার্তা", "Short Code Allocation Procedure 2010", "৫,০০,০০০/- টাকা বার্ষিক ফি", "ডিমান্ড লেটার ইস্যু সাপেক্ষে বরাদ্দ"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "২০",
            "subject": "প্রিমিয়াম কানেক্টিভিটি লিমিটেড (Premium Connectivity Limited) কে আইপিটিএসপি সেবার জন্য বিশেষ শর্ট কোড '১৬৬০৫' বরাদ্দ প্রদান।",
            "decision": "Short Code Allocation Procedure, 2020 অনুযায়ী প্রিমিয়াম কানেক্টিভিটি লিমিটেড-কে বার্ষিক ফিস পরিশোধ সাপেক্ষে শর্ট কোড ১৬৬০৫ বরাদ্দের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "আইপিটিএসপি (IPTSP) লাইসেন্সধারী প্রিমিয়াম কানেক্টিভিটি লিমিটেড এর শর্ট কোড বরাদ্দের আবেদন পরীক্ষা-নিরীক্ষা করে শর্ট কোড ১৬৬০৫ প্রদানের সুপারিশ করা হয়।",
                "tables": [
                    {
                        "headers": ["অপারেটর", "সেবার ধরণ", "বরাদ্দকৃত শর্ট কোড", "অবস্থা"],
                        "rows": [
                            ["Premium Connectivity Limited", "IPTSP Call Service", "১৬৬০৫ (16605)", "অনুমোদিত"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ।",
                "assigned_inspector": "Systems & Services Division",
                "case_status": "অনুমোদিত"
            }
        },
        {
            "agenda_no": "২১",
            "subject": "\"Design, Develop, Implementation of Operation & QoS monitoring system of Fixed Broadband Operators for BTRC\" প্রকল্পের দরপত্র মূল্যায়ন ও নোটিফিকেশন অব অ্যাওয়ার্ড (NOA) প্রদান।",
            "decision": "ফিক্সড ব্রডব্যান্ড গ্রাহক সেবার মান (QoS) পরিবীক্ষণের ডেডিকেটেড সিস্টেম তৈরির দরপত্র মূল্যায়নে নির্বাচিত প্রতিষ্ঠান \"Technometrics Ltd\" কে নোটিফিকেশন অব অ্যাওয়ার্ড (NOA) ইস্যু এবং ITES কর ও ভ্যাট অব্যাহতি প্রদানের প্রস্তাব অনুমোদন করা হলো।",
            "fine_amount": "",
            "details": {
                "presentation_summary": "কমিশনের ব্রডব্যান্ড গ্রাহক সেবার গুণগত মান ও কমপ্লায়েন্স তদারকির লক্ষ্যে এই মনিটরিং সিস্টেম দরপত্র আহ্বান করা হয়। কারিগরি ও আর্থিক মূল্যায়নে Technometrics Ltd সর্বোচ্চ স্কোর অর্জন করায় কাজটির চূড়ান্ত অ্যাওয়ার্ড দেওয়া হয়।",
                "tables": [
                    {
                        "headers": ["প্রকল্পের নাম", "নির্বাচিত প্রতিষ্ঠান", "আবেদনসমূহ", "সিদ্ধান্ত", "কর স্থিতি"],
                        "rows": [
                            ["Fixed Broadband QoS Monitoring System", "Technometrics Ltd", "Dream 71, Intech, Mysoft, Saffron, Devnet-UY", "নোটিফিকেশন অব অ্যাওয়ার্ড (NOA) প্রদান", "ITES Exemption (কর ও ভ্যাট অব্যাহতি)"]
                        ]
                    }
                ],
                "implementation": "সিস্টেমস অ্যান্ড সার্ভিসেস বিভাগ ও অর্থ বিভাগ।",
                "assigned_inspector": "QoS Procurement Committee",
                "case_status": "অ্যাওয়ার্ড প্রদান সম্পন্ন"
            }
        }
    ]
}

# 3.4 Override/Update Meeting 293 Agendas with detailed Excel information
if "293তম" in meetings:
    print("Updating 293rd meeting agendas with detailed Excel data...")
    for agenda in meetings["293তম"]["agendas"]:
        subject_lower = agenda["subject"].lower()
        dec_lower = agenda["decision"].lower()
        
        # Digital One Broadband
        if "digital one" in subject_lower or "digital one" in dec_lower:
            agenda["decision"] = "ডিজিটাল ওয়ান ব্রডব্যান্ড ইন্টারনেট সার্ভিস (উপজেলা/থানা আইএসপি) কর্তৃক ০৪টি রিসেলার এর মাধ্যমে প্রায় ৫৫০ জন গ্রাহককে একইসাথে ডিস ও ইন্টারনেট সেবা প্রদান করা, একই কক্ষে আইএসপি ও ডিশ যন্ত্রপাতি স্থাপন করা এবং স্যাটেলাইট চ্যানেল ডিস্ট্রিবিউটর এর সাথে কোনো চুক্তিপত্র ব্যতীত ডিশ ব্যবসা পরিচালনা করার কারণে বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ আইন, ২০০১ এর বিধান মোতাবেক প্রতিষ্ঠানটির উপর ৫০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপ করে আদায় সাপেক্ষে ৮০% ব্যান্ডউইডথ ক্যাপিং অপসারণের সিদ্ধান্ত গৃহীত হলো।"
            agenda["details"] = {
                "presentation_summary": "পরিদর্শনকালে দেখা যায় যে, প্রতিষ্ঠানটি ৪টি রিসেলার নিয়োগ করে তাদের মাধ্যমে ৫৫০ জন গ্রাহককে একইসাথে ডিশ ও ইন্টারনেট সেবা প্রদান করছে। এছাড়া ডিশ এবং আইএসপি নেটওয়ার্ক যন্ত্রপাতি একই কক্ষে স্থাপন করেছে এবং স্যাটেলাইট চ্যানেল ডিস্ট্রিবিউটরদের সাথে কোনো চুক্তি ছাড়া ডিশ ব্যবসা পরিচালনা করছে, যা লাইসেন্স শর্তের পরিপন্থী।",
                "tables": [
                    {
                        "headers": ["ব্যত্যয়ের ক্ষেত্র", "বিবরণ", "লাইসেন্স বিধি/শর্ত"],
                        "rows": [
                            ["রিসেলার ও ডিশ সেবা", "০৪টি রিসেলার এর মাধ্যমে প্রায় ৫৫০ জন গ্রাহককে একইসাথে ডিস এবং ইন্টারনেট সেবা প্রদান করা।", "আইএসপি গাইডলাইন লঙ্ঘন"],
                            ["যৌথ যন্ত্রপাতি স্থাপন", "আইএসপি ও ডিশ যন্ত্রপাতি একই কক্ষে স্থাপন করে সেবা প্রদান।", "অপারেশনাল নির্দেশনাবলী"],
                            ["ডিশ পাইরেসি", "স্যাটেলাইট চ্যানেল ডিস্ট্রিবিউটর এর সাথে কোনো চুক্তিপত্র ব্যতীত পাইরেসির মাধ্যমে ডিশ ব্যবসা পরিচালনা।", "আইনি শর্ত লঙ্ঘন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ, হিসাব ও রাজস্ব বিভাগ।",
                "assigned_inspector": "পরিদর্শক দল (E&I)",
                "case_status": "জরিমানা আদায় সাপেক্ষে নিষ্পন্ন"
            }
            
        # Chouddagram Broadband
        elif "chouddagram" in subject_lower or "chouddagram" in dec_lower:
            agenda["decision"] = "চৌদ্দগ্রাম ব্রডব্যান্ড (উপজেলা/থানা আইএসপি) কর্তৃক রিসেলারের মাধ্যমে ইন্টারনেট সেবা প্রদান করা, “এক দেশ এক রেট” ট্যারিফ যথাযথভাবে বাস্তবায়ন না করা, IP Log সংরক্ষণ না করা, ৬০%-৪০% শেয়ারিংয়ে রিসেলার চালানো, প্যারেন্টাল কন্ট্রোল না রাখা এবং NOC থেকে PoP এ নিজস্ব ওভারহেড ফাইবার স্থাপন করার অপরাধে ৭৫,০০০/- টাকা প্রশাসনিক জরিমানা আরোপের সিদ্ধান্ত গৃহীত হলো।"
            agenda["details"] = {
                "presentation_summary": "পরিদর্শনকালে চৌদ্দগ্রাম ব্রডব্যান্ড-এর একাধিক ব্যত্যয় পরিলক্ষিত হয়। প্রতিষ্ঠানটি রিসেলার নিয়োগ করে ইন্টারনেট সেবা প্রদান করছে, 'এক দেশ এক রেট' ট্যারিফ মেনে চলছে না এবং গ্রাহকদের আইপি লগ সংরক্ষণ করছে না। এছাড়া ৬০%-৪০% শেয়ারিংয়ে রিসেলার পরিচালনা, প্যারেন্টাল কন্ট্রোল না থাকা এবং অনুমতি ব্যতীত নিজস্ব ফাইবার ফেলার মতো ব্যত্যয়সমূহ করেছে।",
                "tables": [
                    {
                        "headers": ["ব্যত্যয়ের ক্ষেত্র", "বিবরণ", "লাইসেন্স বিধি/শর্ত"],
                        "rows": [
                            ["অবৈধ রিসেলার", "রিসেলারের মাধ্যমে ইন্টারনেট সেবা প্রদান করা এবং ৬০%-৪০% শেয়ারে ব্যবসা করা।", "আইএসপি গাইডলাইন লঙ্ঘন"],
                            ["ট্যারিফ অবমাননা", "“এক দেশ এক রেট” যথাযথভাবে বাস্তবায়ন না করা।", "কমিশনের ট্যারিফ নির্দেশনা"],
                            ["আইপি লগ", "রিসেলারের মাধ্যমে সংযুক্ত গ্রাহকের IP Log সংরক্ষণ না করা।", "নিরাপত্তা ও কমপ্লায়েন্স"],
                            ["লজিস্টিকস ও পলিসি", "প্যারেন্টাল কন্ট্রোল না রাখা এবং NOC থেকে PoP এ নিজস্ব ওভারহেড ফাইবার স্থাপন করা।", "লাইসেন্সিং শর্তাবলী"],
                            ["তথ্য গোপন", "কমিশনের DIS পোর্টালে সঠিক ও ফালনাগাদ তথ্য প্রদান না করা।", "রিপোর্টিং শর্ত লঙ্ঘন"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ, হিসাব ও রাজস্ব বিভাগ।",
                "assigned_inspector": "পরিদর্শক দল (E&I)",
                "case_status": "জরিমানা পরিশোধের নির্দেশ"
            }
            
        # Crazy Net
        elif "crazy net" in subject_lower or "crazy net" in dec_lower or "crazy" in subject_lower:
            agenda["decision"] = "ক্রেজি নেট (উপজেলা/থানা আইএসপি) কর্তৃক মারুফ মারিয়া ডটনেট নামে অবৈধ রিসেলার নিয়োগ, ঋষিপাড়া কার্যালয় হতে ডিশ ও ইন্টারনেট যৌথ পরিচালনা, ক্রেজি নেটের নামে বিল সংগ্রহ না করে রিসেলারের নামে বিল তোলা, IP Log এবং SAF ফর্ম যথাযথভাবে সংগ্রহ না করার অপরাধে ১,০০,০০০/- টাকা প্রশাসনিক জরিমানা আরোপের সিদ্ধান্ত গৃহীত হলো।"
            agenda["details"] = {
                "presentation_summary": "ক্রেজি নেট নামক আইএসপিটির পরিদর্শনে দেখা যায় যে, প্রতিষ্ঠানটি 'মারুফ মারিয়া ডটনেট' নামক রিসেলার নিয়োগ করে ইন্টারনেট সেবা দিচ্ছে এবং তাদের ঋষিপাড়া কার্যালয় থেকে ইন্টারনেট ও ডিশ উভয় সেবা একই সাথে পরিচালনা করছে। এছাড়া ক্রেজি নেটের নামে বিল সংগ্রহ না করে রিসেলারের নামে বিল তোলা হচ্ছে, যা লাইসেন্স শর্তের পরিপন্থী।",
                "tables": [
                    {
                        "headers": ["ব্যত্যয়ের ক্ষেত্র", "বিবরণ", "লাইসেন্স বিধি/শর্ত"],
                        "rows": [
                            ["অবৈধ রিসেলার", "মারুফ মারিয়া ডটনেট নামে রিসেলার নিয়োগের মাধ্যমে ইন্টারনেট সেবা প্রদান করা।", "আইএসপি গাইডলাইন লঙ্ঘন"],
                            ["যৌথ অপারেশন", "ঋষিপাড়া, হেমায়েতপুর কার্যালয় হতে একইসাথে ডিস ও ইন্টারনেট সার্ভিস পরিচালনা।", "অপারেশনাল গাইডলাইন"],
                            ["অবৈধ বিলিং", "PoP স্থাপনায় Crazy Net নামে বিল সংগ্রহ না করে রিসেলারের নামে বিল সংগ্রহ করা।", "আর্থিক ও বিলিং শর্ত"],
                            ["লগ ও SAF সংগ্রহ", "গ্রাহকদের IP log এবং SAF ফর্ম যথাযথভাবে সংরক্ষণ ও সংগ্রহ না করা।", "নিরাপত্তা কমপ্লায়েন্স"],
                            ["রিপোর্টিং", "কমিশনের DIS পোর্টালে সঠিক ও বাস্তব তথ্য প্রদান না করা।", "কমিশনের নির্দেশনা"]
                        ]
                    }
                ],
                "implementation": "ইএন্ডআই ডিরেক্টরেট ও অর্থ, হিসাব ও রাজস্ব বিভাগ।",
                "assigned_inspector": "পরিদর্শক দল (E&I)",
                "case_status": "জরিমানা পরিশোধের নির্দেশ"
            }


# 4. Clean and Sort the meetings
clean_meetings = []
for cm_key, m_data in meetings.items():
    # Sort agendas by agenda_no (numeric or text)
    def agenda_sort_key(a):
        no = a["agenda_no"]
        digits = re.findall(r'\d+', no)
        if digits:
            return int(digits[0])
        return 999  # text agendas like 'বিবিধ' at the end
        
    m_data["agendas"] = sorted(m_data["agendas"], key=agenda_sort_key)
    clean_meetings.append(m_data)

# Sort meetings by numeric index in ascending order
def meeting_sort_key(m):
    num_str = m["meeting_number"].replace('তম', '')
    if num_str.isdigit():
        return int(num_str)
    digits = re.findall(r'\d+', num_str)
    if digits:
        return int(digits[0])
    return 999

sorted_meetings = sorted(clean_meetings, key=meeting_sort_key)

# Add coordination meetings & commissioner instructions data
coordination_mom = [
    {
        "title": "এনফোর্সমেন্ট এন্ড ইন্সপেকশন (ইএন্ডআই) ডিরেক্টরেটের কার্যক্রম বিষয়ক সমন্বয় সভার প্রতিবেদন",
        "date": "২৩ মে, ২০২৬",
        "chairperson": "মাননীয় কমিশনার (স্পেকট্রাম বিভাগ)",
        "attendees": "পরিচালক (ইএন্ডআই) এবং অন্যান্য কর্মকর্তাবৃন্দ",
        "topics": [
            {
                "section": "১. বকেয়া রাজস্ব ও প্রশাসনিক জরিমানা আদায়",
                "summary": "বিগত দেড় বছরে মোট ৮৩টি কেইসের বিপরীতে প্রায় ৪৬ কোটি টাকার বকেয়া বা জরিমানার বিষয় চলমান রয়েছে। এর মধ্যে ৩৯.৯৮ কোটি টাকার (৬৫টি কেইস) রিভিউ প্রক্রিয়া চলমান, ৩.৮১ কোটি টাকা রেভিনিউ শেয়ারিং সংক্রান্ত এবং প্রায় ২ কোটি টাকা প্রশাসনিক জরিমানা। লেভেল থ্রি, আর্থ (উভয়ের ১৫ কোটি করে), আইটিসি, সামিট, নভোকম, ফাইবার অ্যাট হোম, ম্যাঙ্গো এবং বিটিসিএল-এর মতো প্রতিষ্ঠানগুলোর জরিমানা ও রিভিউ এর অন্তর্ভুক্ত।",
                "decision": "বকেয়া আদায়ের জন্য অর্থ বিভাগের সাথে সমন্বয়হীনতা দূর করতে হবে। মাননীয় কমিশনার নির্দেশ দিয়েছেন যে, এই ৮৩টি কেইসের প্রতিটি ধাপ (যেমন- পরিদর্শনের তারিখ, কমিশনের সিদ্ধান্ত, আপিল, বর্তমান অবস্থা, দায়িত্বপ্রাপ্ত কর্মকর্তা) ট্র্যাক করার জন্য একটি বিস্তারিত মাস্টার এক্সেল ডেটাবেস তৈরি করতে হবে, যাতে এক নজরে সব কেইসের সর্বশেষ অবস্থা জানা যায়।"
            },
            {
                "section": "২. অবৈধ ভিওআইপি (VoIP) অভিযান ও জব্দকৃত সিম",
                "summary": "এনটিএমসি-এর সহায়তায় গত সপ্তাহে চট্টগ্রামে ২১৬টি টেলিটক সিম এবং এপ্রিলে কুমিল্লায় প্রায় ৮,০০০ সিম (যার মধ্যে রবিরই ৫,৯০০টি) জব্দ করা হয়েছে। এই অভিযানগুলোর প্রেক্ষিতে থানায় মামলা দায়ের করা হয়েছে, জব্দ তালিকা করা হয়েছে এবং অপারেটরদের শুনানির জন্য ডাকা হয়েছে। পরবর্তী আইনি পদক্ষেপের জন্য বিষয়টি লিগাল ও লাইসেন্সিং বিভাগে পাঠানো হয়েছে।",
                "decision": "এই মামলা ও জরিমানার বিষয়গুলো পরবর্তী কমিশন সভায় হালনাগাদ তথ্য হিসেবে উপস্থাপন করে কমিশনকে অবহিত করতে হবে।"
            },
            {
                "section": "৩. জব্দকৃত সিমের অব্যবহৃত ব্যালেন্স (Unused Balance)",
                "summary": "২০১৭ সালের ডিরেক্টিভ অনুযায়ী গ্রামীণফোন, বাংলালিংক, রবি এবং টেলিটকের কাছে রিসাইকেলকৃত সিমের অব্যবহৃত ব্যালেন্স বাবদ প্রায় ১২ কোটি ৪৪ লাখ টাকা পাওনা রয়েছে। মোবাইল অপারেটররা দাবি করছে যে ১২ বছরের তথ্য তাদের কাছে নেই এবং ডেটা ডিলিট হয়ে গেছে।",
                "decision": "কোম্পানি আইন অনুযায়ী তথ্য সংরক্ষণের বাধ্যবাধকতার বিষয়টি লিগাল দৃষ্টিকোণ থেকে যাচাই করে অপারেটরদের কাছ থেকে এই টাকা উদ্ধারের প্রক্রিয়া জোরদার করতে হবে। এছাড়া ২০১৭ সালের পুরোনো ডিরেক্টিভটি হালনাগাদ করার জন্য এসএস (SS) বিভাগের সাথে কাজ করতে হবে।"
            },
            {
                "section": "৪. আইএসপি (ISP) অপারেটরদের অভ্যন্তরীণ বিরোধ",
                "summary": "গত এক মাসে ৯টি আইএসপি অপারেটরের পক্ষ থেকে একে অপরের ফাইবার কাটা এবং কর্মীদের মারধর করার অভিযোগ এসেছে। ইএন্ডআই বিভাগ তাদেরকে ডেকে শুনানি করে মুচলেকা বা অঙ্গীকারনামা নিয়েছে।",
                "decision": "মারামারি বা সিভিল/ক্রিমিনাল বিরোধের মতো বিষয়গুলো দেখা বিটিআরসির কাজ নয়, এগুলো পুলিশের এখতিয়ারভুক্ত। বিটিআরসি শুধুমাত্র রেগুলেটরি কমপ্লায়েন্স, লাইসেন্স নবায়ন, এবং গ্রাহক ভোগান্তি হচ্ছে কিনা তা দেখবে। এ ধরনের মাইক্রো-ম্যানেজমেন্টে ইএন্ডআই বিভাগের সময় নষ্ট না করার নির্দেশ দেওয়া হয়।"
            },
            {
                "section": "৫. সাসপেক্টেড ভিওআইপি সিম ব্লক ও আনব্লক",
                "summary": "সন্দেহভাজন ভিওআইপি ব্যবহারের কারণে ব্লক হওয়া সিম পুনরায় চালুর জন্য প্রায় ৮০০টি আবেদন জমা পড়েছে। যাচাই-বাছাই শেষে ইতোমধ্যে ৬৪টি সিম আনব্লক করার চিঠি দেওয়া হয়েছে, ৩৬টির প্রক্রিয়া চলমান এবং ৫১টি স্থায়ীভাবে বন্ধ করে দেওয়া হয়েছে।",
                "decision": "সিম ব্লক/আনব্লক করার ক্ষেত্রে নিয়মাবলী অনুসরণপূর্বক দ্রুত কেইস নিষ্পত্তি করা হবে।"
            },
            {
                "section": "৬. নতুন পলিসি ও গাইডলাইন আপডেট",
                "summary": "নতুন সরকারের অধীনে আইএলডিটিএস (ILDTS) পলিসি এবং গাইডলাইন হালনাগাদের কাজ প্রায় শেষ পর্যায়ে রয়েছে। পূর্বে যেখানে ২৯ ক্যাটাগরির লাইসেন্স ছিল, তা সিম্প্লিফাই করে ৩টি লেয়ারে নামিয়ে আনা হচ্ছে। এই নতুন পলিসি জারি হলে লাইসেন্স প্রদান ও নবায়ন প্রক্রিয়া আরও সহজ হবে। কল সেন্টারগুলোর লাইসেন্সিং ডিরেগুলেট (Deregulate) করা হচ্ছে বলেও জানানো হয়।",
                "decision": "পলিসি ও গাইডলাইন হালনাগাদকরণ কার্যক্রম ত্বরান্বিত করা।"
            },
            {
                "section": "৭. লজিস্টিকস এবং অপারেশনাল চ্যালেঞ্জসমূহ",
                "summary": "field পরিদর্শনে যাতায়াত ও নিরাপত্তার জন্য পর্যাপ্ত গাড়ির অভাব রয়েছে। পরিদর্শনের ছবি, অডিও ও আলামত সংরক্ষণের জন্য সেন্ট্রাল স্টোরেজ বা ক্যামেরার অভাব রয়েছে, বর্তমানে কর্মকর্তারা ব্যক্তিগত ডিভাইস ব্যবহার করছেন।",
                "decision": "মাঠ পর্যায়ে পরিদর্শনের সময় কর্মকর্তাদের ব্যক্তিগত নিরাপত্তা সর্বোচ্চ অগ্রাধিকার দেওয়ার নির্দেশ দেওয়া হয়। ডেটা সংরক্ষণের জন্য ডেডিকেটেড ডিভাইস/ল্যাপটপ ক্রয়ের ব্যবস্থা গ্রহণ করা এবং কাজের পরিধি সুনির্দিষ্ট করতে এসওপি (SOP) সঠিকভাবে মেনে চলার নির্দেশ।"
            }
        ]
    }
]

commissioner_instructions = [
    {
        "title": "কমিশনার (স্পেকট্রাম বিভাগ) মহোদয় কর্তৃক ইএন্ডআই (E&I) ডিরেক্টরেটের জন্য প্রদত্ত নির্দেশনাসমূহের সারসংক্ষেপ",
        "date": "২৩ মে, ২০২৬",
        "source": "cm_sm_instruction.docx",
        "instructions": [
            {
                "id": "১",
                "title": "মাস্টার ডেটাবেস তৈরি",
                "text": "বিগত দিনের ৮৩টি কেইসের (জরিমানা ও রেভিনিউ শেয়ারিং) প্রতিষ্ঠানের নাম, পরিদর্শনের তারিখ, আপিলের অবস্থা এবং বর্তমান দায়িত্বপ্রাপ্ত কর্মকর্তার নামসহ একটি পূর্ণাঙ্গ মাস্টার এক্সেল শিট তৈরি করতে হবে (যা আগামী মিটিংয়ে পর্যালোচনা করা হবে)।"
            },
            {
                "id": "২",
                "title": "অর্থ বিভাগের সাথে সমন্বয়",
                "text": "জরিমানা ও বকেয়া আদায়ের সামারি নিয়ে অর্থ বিভাগের সাথে মিটিং করা এবং কত টাকা আদায় (Recovered) হয়েছে বা বাকি আছে, তার একটি স্টেটমেন্ট বের করে কোঅর্ডিনেট করা।"
            },
            {
                "id": "৩",
                "title": "আইএসপি বিরোধে না জড়ানো",
                "text": "আইএসপিদের অভ্যন্তরীণ মারামারি বা তার কাটার মতো ফৌজদারি/দেওয়ানি বিষয়ে বিটিআরসি জড়াবে না, তাদের পুলিশের কাছে যাওয়ার পরামর্শ দিতে হবে। বিটিআরসি শুধু লাইসেন্স কমপ্লায়েন্স ও গ্রাহক ভোগান্তি দেখবে এবং এ পর্যন্ত নেওয়া মুচলেকাগুলো কমিশনকে অবহিত করবে।"
            },
            {
                "id": "৪",
                "title": "অবৈধ ভিওআইপি আপডেট",
                "text": "কুমিল্লা ও চট্টগ্রামে সাম্প্রতিক অভিযানের তথ্য, জব্দকৃত সিম ও মামলার আপডেট পরবর্তী কমিশন সভায় উপস্থাপন করতে হবে।"
            },
            {
                "id": "৫",
                "title": "অব্যবহৃত ব্যালেন্স উদ্ধার",
                "text": "মোবাইল অপারেটরদের 'ডেটা ডিলিট' দাবির সত্যতা কোম্পানি আইন (১২ বছরের ডেটা রাখার বাধ্যবাককতা) অনুযায়ী যাচাই করতে হবে। এ বিষয়ে পরবর্তীতে আলাদাভাবে বসা হবে।"
            },
            {
                "id": "৬",
                "title": "কর্মকর্তাদের ব্যক্তিগত নিরাপত্তা",
                "text": "মাঠ পর্যায়ে পরিদর্শনের সময় কর্মকর্তাদের ব্যক্তিগত সুরক্ষাকে সর্বোচ্চ অগ্রাধিকার দিতে হবে এবং কোনো অপ্রয়োজনীয় ঝুঁকি বা সংঘাতে জড়ানো যাবে না।"
            },
            {
                "id": "৭",
                "title": "সেন্ট্রাল ডেটা স্টোরেজ",
                "text": "পরিদর্শনের ছবি, অডিও ও আলামত সংরক্ষণের জন্য ব্যক্তিগত মোবাইল/ল্যাপটপ ব্যবহার না করে একটি সেন্ট্রাল স্টোরেজ এবং ভালো রেজুলেশনের ক্যামেরা/ডিভাইসের ব্যবস্থা করতে হবে।"
            }
        ]
    }
]

master_db = {
    "meetings": sorted_meetings,
    "coordination_mom": coordination_mom,
    "commissioner_instructions": commissioner_instructions
}

# Save to JSON
with open(output_json_path, "w", encoding="utf-8") as f:
    json.dump(master_db, f, ensure_ascii=False, indent=2)

print(f"Master Database successfully compiled and saved to: {output_json_path}")
print(f"Total meetings compiled: {len(sorted_meetings)}")
sys.stdout.flush()
