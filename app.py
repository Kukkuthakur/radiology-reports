"""
Radiology Report Generator — USG Whole Abdomen
Streamlit version (Phase 1)
"""

import io
import os
import sqlite3
from datetime import datetime

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# CONFIG
# ============================================================

DB_PATH = os.path.join(os.path.expanduser("~"), "RadiologyReports", "reports.db")

PAGE_TOP_MARGIN = 5.5
PAGE_BOTTOM_MARGIN = 4.0
PAGE_LEFT_MARGIN = 2.0
PAGE_RIGHT_MARGIN = 2.0

FONT_BODY = "Calibri"
FONT_SIZE_BODY = 11
FONT_SIZE_TITLE = 12
FONT_SIZE_DISCLAIMER = 10

DISCLAIMER_TEXT = (
    "THIS IS ONLY A PROFESSIONAL OPINION NOT THE FINAL DIAGNOSIS. IT SHOULD BE "
    "CORRELATED CLINICALLY AND WITH OTHER RELEVANT INVESTIGATIONS. NOT VALID "
    "FOR MEDICO LEGAL PURPOSES."
)

IMPRESSION_NORMAL_FEMALE = [
    "NO FREE FLUID, LYMPHADENOPATHY, OBVIOUS BOWEL WALL THICKENING OR ANY SIGN OF INFLAMMATION SEEN AT THE TIME OF SCAN.",
    "UNREMARKABLE ABDOMEN & PELVIC SCAN.",
]
IMPRESSION_NORMAL_MALE = [
    "NO FREE FLUID, LYMPHADENOPATHY, OBVIOUS BOWEL WALL THICKENING OR ANY SIGN OF INFLAMMATION SEEN AT THE TIME OF SCAN.",
    "UNREMARKABLE ABDOMEN SCAN.",
]
IMPRESSION_REST_UNREMARKABLE = "REST OF THE ABDOMEN SCAN IS UNREMARKABLE."


# ============================================================
# DATABASE
# ============================================================

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS referrers (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL)""")
    c.execute("""CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        patient_name TEXT, age TEXT, sex TEXT, report_date TEXT,
        referred_by TEXT, created_at TEXT, data_json TEXT)""")
    for name in ["SELF", "J.K HOSPITAL", "GOYAL CITY HOSPITAL", "SAROJ HOSPITAL",
                 "NAV JEEVAN CLINIC", "SAI CLINIC", "MAA PITAMBARA HOSPITAL",
                 "SANJEEVANI HOSPITAL", "PUNJABI CLINIC", "DR. P.K GUPTA(M.D)",
                 "DR. RAJU KHAN", "DR. K.L KUSHWAH", "DR. USHA", "DR. AJAY VEER",
                 "DR. BHUPENDRA SINGH", "DR. R.M SINGH", "RAM KRISHNA HOSPITAL"]:
        c.execute("INSERT OR IGNORE INTO referrers (name) VALUES (?)", (name,))
    conn.commit()
    conn.close()


def get_referrers():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT name FROM referrers ORDER BY name")
    rows = [r[0] for r in c.fetchall()]
    conn.close()
    return rows


def add_referrer(name):
    if not name.strip():
        return
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO referrers (name) VALUES (?)", (name.strip(),))
    conn.commit()
    conn.close()


def save_report(data):
    import json
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""INSERT INTO reports (patient_name, age, sex, report_date, referred_by, created_at, data_json)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
              (data["patient"]["name"], data["patient"]["age"], data["patient"]["sex"],
               data["patient"]["date"], data["patient"]["referred_by"],
               datetime.now().isoformat(), json.dumps(data)))
    conn.commit()
    conn.close()


# ============================================================
# REPORT DATA MODEL
# ============================================================

def new_report(sex="F"):
    return {
        "patient": {"name": "", "age": "", "sex": sex,
                    "date": datetime.now().strftime("%d-%b-%y"), "referred_by": ""},
        "liver": {"size_mm": "", "size_descriptor": "normal", "echotexture": "normal",
                  "steatosis_grade": None, "focal_lesion": "none", "focal_lesion_text": "",
                  "ihbr": "normal", "portal_vein": "normal", "portal_vein_mm": ""},
        "gall_bladder": {"status": "adequately_distended", "wall": "normal", "wall_mm": "",
                         "calculi": "none", "calculi_count": "multiple",
                         "calculi_largest_mm": "", "calculi_location": "body",
                         "pericholecystic_fluid": False, "sludge": "none",
                         "sludge_size_mm": "", "sludge_location": "", "comet_tail": False},
        "cbd": {"caliber_mm": "", "status": "normal", "calculi": False,
                "calculi_size_mm": "", "calculi_location": "distal", "ihbr_dilated": False},
        "pancreas": {"status": "normal", "fat_stranding_grade": "mild", "ln_size": "",
                     "mpd_dilated": False, "mpd_mm": ""},
        "spleen": {"size_mm": "", "size_descriptor": "normal", "portal_vein_mm": ""},
        "kidneys": {
            "right": {"status": "normal", "calculi": [], "cyst": "none", "cyst_size_mm": "",
                      "cyst_location": "", "bosniak": "", "hydronephrosis": "none",
                      "nephrocalcinosis": "none"},
            "left":  {"status": "normal", "calculi": [], "cyst": "none", "cyst_size_mm": "",
                      "cyst_location": "", "bosniak": "", "hydronephrosis": "none",
                      "nephrocalcinosis": "none"},
            "cortical_echogenicity": "normal"},
        "urinary_bladder": {"status": "adequately_distended", "mass_calculus": False,
                            "sedimentation": "none"},
        "uterus": {"status": "anteverted", "size": "", "myometrium": "homogenous",
                   "fibroid_text": "", "endometrial_thickness_mm": "",
                   "endometrial_collection": False},
        "ovaries": {"right_status": "normal", "right_size": "",
                    "right_cyst_type": "simple", "right_cyst_size": "",
                    "left_status": "normal", "left_size": "",
                    "left_cyst_type": "simple", "left_cyst_size": "",
                    "afc": "normal"},
        "prostate": {"status": "normal", "size_cc": ""},
        "bowel": {"wall_thickening": False, "free_fluid": "none",
                  "mesenteric_ln": "none", "ln_size_category": "sad_lt_7",
                  "ln_location": "bilateral", "ln_largest": "",
                  "ln_character": "discrete", "pleural_effusion": "none"},
        "appendix": {"status": "not_assessed", "diameter_mm": ""},
        "impression": {"lines": []},
    }


# ============================================================
# SENTENCE GENERATORS
# ============================================================

def seg(text, bold=False, underline=False):
    return (text, bold, underline)


def liver_sentence(d, sex, age):
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    echo = d["echotexture"]
    is_pediatric = False
    try:
        is_pediatric = int(age) <= 12 if age else False
    except (ValueError, TypeError):
        pass
    s = [seg("LIVER", True, True)]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}MM)"))
    elif desc == "mild":
        word = "MILDLY ENLARGED" + (" FOR AGE" if is_pediatric else "")
        s += [seg(" is "), seg(f"{word} IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("MODERATELY ENLARGED IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "severe":
        s += [seg(" is "), seg("SEVERELY ENLARGED IN SIZE", True), seg(f" ({size}MM)")]

    if echo == "normal":
        s.append(seg(" with normal outline and echotexture."))
    elif echo == "increased":
        s += [seg(" with normal outline and "), seg("INCREASED REFLECTIVITY", True), seg(".")]
    elif echo == "coarse":
        s += [seg(" with normal outline and "), seg("COARSE ECHOTEXTURE", True), seg(".")]

    if d["focal_lesion"] == "none":
        s.append(seg(" No focal lesion is seen."))
    elif d["focal_lesion"] == "calcified":
        s += [seg(" "), seg("A CALCIFIED FOCUS SEEN IN THE RIGHT HEPATIC LOBE.", True)]

    if d["ihbr"] == "normal":
        s.append(seg(" Intra hepatic biliary radicals are normal."))
    else:
        s += [seg(" Intra hepatic biliary radicals are "), seg("DILATED", True), seg(".")]

    if d["portal_vein"] == "normal":
        s.append(seg(" Portal vein is normal in course and caliber."))
    else:
        s += [seg(" Portal vein is "), seg("DILATED", True)]
        if d["portal_vein_mm"]:
            s.append(seg(f" ({d['portal_vein_mm']}MM)"))
        s.append(seg("."))
    return s


def gall_bladder_sentence(d):
    s = [seg("GALL BLADDER", True, True)]
    status = d["status"]
    if status == "adequately_distended":
        s.append(seg(" is adequately distended."))
    elif status == "over":
        s += [seg(" is "), seg("OVER-DISTENDED", True), seg(".")]
    elif status == "partially":
        s.append(seg(" is partially contracted."))
    elif status == "contracted":
        s.append(seg(" is contracted (suboptimal wall visualization)."))
    elif status == "empty":
        s.append(seg(" is empty."))
    elif status == "operated":
        s += [seg(" is operated. "), seg("GB FOSSA", True, True), seg(" is unremarkable.")]
        return s

    if d["wall"] == "normal":
        s.append(seg(" Wall thickness is normal."))
    else:
        s += [seg(" "), seg(f"WALL IS THICKENED UPTO {d['wall_mm']}MM", True), seg(".")]

    if d["calculi"] == "none":
        s.append(seg(" No obvious gall stones seen."))
    else:
        cnt = {"single": "A CALCULUS IS SEEN", "couple": "A COUPLE OF CALCULI ARE SEEN",
               "multiple": "MULTIPLE CALCULI ARE SEEN",
               "multiple_tiny": "MULTIPLE TINY CALCULI ARE SEEN",
               "few": "FEW CALCULI ARE SEEN"}.get(d["calculi_count"], "CALCULI ARE SEEN")
        s += [seg(" "), seg(f"{cnt} IN THE GB LUMEN", True)]
        if d["calculi_largest_mm"]:
            s.append(seg(f", LARGEST OF THESE MEASURING {d['calculi_largest_mm']}MM", True))
        s.append(seg(".", True))
        if not d["pericholecystic_fluid"]:
            s += [seg(" HOWEVER NO PERICHOLECYSTIC FLUID OR GB WALL THICKENING IS SEEN", True),
                  seg(".", True)]

    if d["sludge"] == "trace":
        s += [seg(" "), seg("TRACE SLUDGE SEEN IN THE GB LUMEN.", True)]
    elif d["sludge"] == "sludge":
        s += [seg(" "), seg("SLUDGE SEEN IN THE GB LUMEN.", True)]
    elif d["sludge"] == "sludge_ball":
        size_txt = f"({d['sludge_size_mm']}MM) " if d["sludge_size_mm"] else ""
        loc = d["sludge_location"] or "FUNDUS"
        s += [seg(" "), seg(f"A SLUDGE BALL {size_txt}SEEN AT THE GB {loc}.", True)]

    if d["comet_tail"]:
        s += [seg(" "), seg("MULTIPLE COMET TAIL ARTIFACTS SEEN ARISING FROM THE ANTERIOR GB WALL, "
                            "SUGGESTIVE OF ADENOMYOMATOSIS/CHOLESTEROLOSIS", True), seg(".")]
    return s


def cbd_sentence(d):
    s = [seg("COMMON BILE DUCT", True, True)]
    if d["status"] == "normal":
        s.append(seg(f" is normal in caliber({d['caliber_mm']}MM)." if d["caliber_mm"]
                     else " is normal in caliber."))
    else:
        s += [seg(" is "), seg("DILATED", True)]
        if d["caliber_mm"]:
            s.append(seg(f" UPTO {d['caliber_mm']}MM", True))
        s.append(seg(".", True))
        if d["calculi"]:
            s += [seg(" "), seg(f"SUGGESTION OF A CALCULUS MEASURING {d['calculi_size_mm']}MM "
                                f"IN THE {d['calculi_location'].upper()} SEGMENT", True), seg(".", True)]
        if d["ihbr_dilated"]:
            s += [seg(" "), seg("INTRA HEPATIC BILIARY RADICALS ARE ALSO PROXIMALLY DILATED.", True)]
    return s


def pancreas_sentence(d):
    s = [seg("PANCREAS", True, True)]
    if d["status"] == "normal":
        s.append(seg(" is normal in size, outline and echotexture. No focal lesion is seen. "
                     "No evidence of calcification is seen."))
    elif d["status"] == "fat_stranding":
        s.append(seg(" is normal in size, outline and echotexture. No focal lesion is seen. "
                     "No evidence of calcification is seen. "))
        s += [seg(f"{d['fat_stranding_grade'].upper()} PERI-PANCREATIC FAT STRANDING IS SEEN", True),
              seg(".")]
    elif d["status"] == "necrotic_ln":
        s.append(seg(" is normal in size, outline and echotexture. No focal lesion is seen. "
                     "No evidence of calcification is seen. "))
        s += [seg(f"A NECROTIC PERI-PANCREATIC LYMPH NODE MEASURING {d['ln_size']} IS SEEN", True),
              seg(".")]
    if d["mpd_dilated"]:
        s += [seg(" "), seg(f"MPD IS DILATED IN CALIBER ({d['mpd_mm']}MM)", True), seg(".")]
    return s


def spleen_sentence(d):
    s = [seg("SPLEEN", True, True)]
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}MM) with normal echotexture. Splenic vein is normal."))
    elif desc == "borderline":
        s += [seg(" is "), seg("BORDERLINE ENLARGED IN SIZE", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    elif desc == "mild":
        s += [seg(" is "), seg("MILDLY ENLARGED IN SIZE", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    elif desc == "moderate":
        s += [seg(" is "), seg("MODERATELY ENLARGED IN SIZE", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    if d["portal_vein_mm"]:
        s += [seg(" "), seg("PORTAL VEIN IS NORMAL IN COURSE AND CALIBER", True),
              seg(f" ({d['portal_vein_mm']}MM)"), seg(".", True)]
    return s


def kidneys_sentence(d):
    s = []
    r, l = d["right"], d["left"]
    both_normal = (r["status"] == "normal" and not r["calculi"] and r["cyst"] == "none"
                   and r["hydronephrosis"] == "none" and r["nephrocalcinosis"] == "none"
                   and l["status"] == "normal" and not l["calculi"] and l["cyst"] == "none"
                   and l["hydronephrosis"] == "none" and l["nephrocalcinosis"] == "none")
    if both_normal:
        s.append(seg("BOTH KIDNEYS", True, True))
        s.append(seg(" are normal in size, outline and "))
        if d["cortical_echogenicity"] == "mildly_raised_bilateral":
            s += [seg("MILDLY RAISED BILATERAL RENAL CORTICAL ECHOGENICITY", True),
                  seg(". Corticomedullary differentiation is maintained. "
                      "No evidence of hydronephrotic changes/calculus seen.")]
        else:
            s.append(seg("echogenicity. Corticomedullary differentiation is maintained. "
                         "No evidence of hydronephrotic changes/calculus seen."))
        return s

    def block(name, k, side_label):
        b = [seg(name, True, True)]
        if k["status"] == "not_visualized":
            b.append(seg(" is not visualized."))
            return b
        b.append(seg(" is normal in size, outline and echogenicity. "
                     "Corticomedullary differentiation is maintained."))
        for calc in k["calculi"]:
            b += [seg(" "), seg(f"A CALCULUS MEASURING {calc.get('size_mm','')}MM IS SEEN AT THE "
                                f"{calc.get('location','').upper()} OF {side_label} KIDNEY", True)]
            if calc.get("hydro") and calc["hydro"] != "none":
                b.append(seg(f" CAUSING {side_label} SIDED {calc['hydro'].upper()} "
                             f"HYDROURETERONEPHROSIS", True))
            b.append(seg(".", True))
        if k["cyst"] != "none":
            bosniak = f" (BOSNIAK CAT-{k['bosniak']})" if k["bosniak"] else ""
            b += [seg(" "), seg(f"A {k['cyst'].upper()} CYST MEASURING {k['cyst_size_mm']}MM IS SEEN "
                                f"AT THE {k['cyst_location'].upper()} OF {side_label} KIDNEY{bosniak}",
                                True), seg(".", True)]
        if k["hydronephrosis"] != "none":
            b += [seg(" "), seg(f"{k['hydronephrosis'].upper()} HYDROURETERONEPHROSIS IS PRESENT ON "
                                f"THE {side_label}", True), seg(".", True)]
        if k["nephrocalcinosis"] != "none":
            b += [seg(" "), seg(f"MULTIPLE FOCI OF CALCIFICATION SEEN IN THE {side_label} "
                                f"RENAL CORTEX", True), seg(".", True)]
        return b

    s.extend(block("RIGHT KIDNEY", r, "RIGHT"))
    s.append(seg("\n"))
    s.extend(block("LEFT KIDNEY", l, "LEFT"))
    return s


def urinary_bladder_sentence(d):
    s = [seg("URINARY BLADDER", True, True)]
    if d["status"] == "adequately_distended":
        s.append(seg(" is adequately distended."))
    elif d["status"] == "over":
        s += [seg(" is "), seg("OVER-DISTENDED", True), seg(".")]
    elif d["status"] == "partially":
        s.append(seg(" is partially empty."))
    elif d["status"] == "empty":
        s.append(seg(" is empty."))
    if not d["mass_calculus"]:
        s.append(seg(" No mass or calculus seen."))
    sed = d["sedimentation"]
    if sed == "trace":
        s += [seg(" "), seg("TRACE SEDIMENTATION SEEN IN THE UB LUMEN.", True)]
    elif sed == "free_floating":
        s += [seg(" "), seg("FREE FLOATING SEDIMENTATION SEEN IN THE UB LUMEN", True), seg(".")]
    elif sed == "significant":
        s += [seg(" "), seg("SIGNIFICANT SEDIMENTATION SEEN IN THE UB LUMEN", True), seg(".")]
    elif sed == "extensive":
        s += [seg(" "), seg("EXTENSIVE SEDIMENTATION SEEN IN THE UB LUMEN", True), seg(".")]
    return s


def uterus_sentence(d, pediatric=False):
    if pediatric:
        s = [seg("UTERUS & BOTH OVARIES", True, True)]
        if d["status"] == "not_visualized":
            s.append(seg(" are not visualized."))
        else:
            s.append(seg(" are normal for age."))
        return s
    s = [seg("UTERUS", True, True)]
    status = d["status"]
    if status == "operated":
        s.append(seg(" is operated."))
        return s
    if status == "not_visualized":
        s.append(seg(" is not visualized."))
        return s
    if status in ("anteverted", "retroverted", "retroflexed"):
        s.append(seg(f" is {status.upper() if status != 'anteverted' else 'anteverted'}"))
        if d["size"]:
            s.append(seg(f" and normal in size({d['size']}MM)"))
        s.append(seg(" with normal shape and echopattern."))
    elif status == "bulky":
        s += [seg(" is "), seg("BULKY IN SIZE", True)]
        if d["size"]:
            s.append(seg(f"({d['size']}MM)"))
        s.append(seg(" with normal shape and echopattern."))
    elif status == "partially":
        s.append(seg(" is partially visualized and appears anteverted"))
        if d["size"]:
            s.append(seg(f", normal in size({d['size']}MM)"))
        s.append(seg(" with normal shape and echopattern."))
    if d["myometrium"] == "homogenous":
        s.append(seg(" Myometrium appears homogenous and no focal lesion is seen."))
    elif d["myometrium"] == "fibroid" and d["fibroid_text"]:
        s += [seg(" "), seg(d["fibroid_text"].upper(), True), seg(".", True)]
    if d["endometrial_thickness_mm"]:
        s.append(seg(f" Endometrial echo is central and regular in "
                     f"thickness({d['endometrial_thickness_mm']}MM)."))
        if not d["endometrial_collection"]:
            s.append(seg(" No collection seen in the endometrial canal."))
    return s


def ovaries_sentence(d):
    s = []
    r, l = d["right_status"], d["left_status"]
    if r == "normal" and l == "normal":
        s += [seg("BOTH OVARIES", True, True),
              seg(" appears normal in size and echo pattern.")]
        if d["right_size"] or d["left_size"]:
            parts = []
            if d["right_size"]:
                parts.append(f"RO={d['right_size']}MM")
            if d["left_size"]:
                parts.append(f"LO={d['left_size']}MM")
            s += [seg(" "), seg(f"[{';'.join(parts)}].", True)]
        return s

    def block(label, status, size, cyst_type, cyst_size):
        b = [seg(label, True, True)]
        if status == "not_visualized":
            b.append(seg(" not visualized(likely atrophic)."))
        elif status == "normal":
            b.append(seg(" appears normal in size and echo pattern."))
            if size:
                b += [seg(" "), seg(f"[{size}MM].", True)]
        elif status == "cyst":
            if size:
                b += [seg(" appears normal in size and echo pattern. "), seg(f"[{size}MM].", True)]
            b += [seg(" "), seg(f"A {cyst_type.upper()} CYST MEASURING {cyst_size}MM IS SEEN IN THE "
                                f"{label.split()[0]} OVARY", True), seg(".", True)]
        return b

    s.extend(block("RIGHT OVARY", r, d["right_size"], d["right_cyst_type"], d["right_cyst_size"]))
    s.append(seg(" "))
    s.extend(block("LEFT OVARY", l, d["left_size"], d["left_cyst_type"], d["left_cyst_size"]))
    return s


def prostate_sentence(d, pediatric=False):
    s = [seg("PROSTATE", True, True)]
    if pediatric and d["status"] in ("normal", "age_appropriate"):
        s.append(seg(" is age-appropriate."))
        return s
    if d["status"] == "not_visualized":
        s.append(seg(" is not visualized."))
        return s
    if d["status"] == "normal":
        s.append(seg(" is normal in size and attenuation. No diffuse or focal lesion seen."))
    elif d["status"] == "borderline":
        s += [seg(" is "), seg("BORDERLINE ENLARGED IN SIZE", True)]
        if d["size_cc"]:
            s.append(seg(f"({d['size_cc']}CC)"))
        s.append(seg(" and attenuation. No diffuse or focal lesion seen."))
    elif d["status"] == "bulky":
        s += [seg(" is "), seg("BULKY IN SIZE", True)]
        if d["size_cc"]:
            s.append(seg(f"({d['size_cc']}CC)"))
        s.append(seg(" and attenuation. No diffuse or focal lesion seen."))
    elif d["status"] == "grade1":
        s += [seg(" shows "), seg(f"GRADE-I PROSTATOMEGALY ({d['size_cc']}CC)", True),
              seg(". No diffuse or focal lesion seen.")]
    return s


def bowel_sentence(d, sex):
    s = []
    if sex == "F":
        if not d["wall_thickening"]:
            s.append(seg("NO OBVIOUS BOWEL WALL THICKENING SEEN.", True))
        else:
            s.append(seg("BOWEL WALL THICKENING SEEN.", True))
        s.append(seg(" "))
        if d["free_fluid"] == "none":
            s.append(seg("No free fluid seen in the peritoneal cavity."))
        else:
            s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free fluid seen.", True))
    else:
        if d["free_fluid"] == "none":
            s.append(seg("No free fluid is seen in the peritoneal cavity."))
        else:
            s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free fluid seen.", True))
        s.append(seg(" "))
        if not d["wall_thickening"]:
            s.append(seg("NO OBVIOUS BOWEL WALL THICKENING SEEN.", True))
        else:
            s.append(seg("BOWEL WALL THICKENING SEEN.", True))
    if d["mesenteric_ln"] == "present":
        cat = {"sad_lt_7": "SMALL(SAD<7MM)", "sad_gt_7": "ENLARGED(SAD>7MM)",
               "sad_gt_10": "ENLARGED(SAD>10MM)"}.get(d["ln_size_category"], "MESENTERIC")
        loc = d["ln_location"].replace("_", " ").upper()
        largest = f" WITH LARGEST OF THESE MEASURING {d['ln_largest']}" if d["ln_largest"] else ""
        s += [seg(" "), seg(f"{cat} MESENTERIC LYMPH NODES ARE SEEN IN THE {loc} REGION{largest}",
                            True), seg(".", True)]
    pe = d["pleural_effusion"]
    if pe != "none":
        text = {"trace_right": "TRACE RIGHT PLEURAL EFFUSION IS SEEN",
                "trace_left": "TRACE LEFT PLEURAL EFFUSION IS SEEN",
                "mild_right": "MILD RIGHT PLEURAL EFFUSION IS SEEN",
                "mild_bilateral": "TRACE LEFT & MILD RIGHT PLEURAL EFFUSION SEEN"}.get(pe, "")
        s += [seg(" "), seg(text, True), seg(".", True)]
    return s


def appendix_sentence(d):
    if d["status"] == "not_assessed":
        return []
    s = []
    if d["status"] == "not_visualized":
        s.append(seg("APPENDIX IS NOT VISUALIZED."))
    elif d["status"] == "normal":
        s.append(seg("APPENDIX IS VISUALIZED AND APPEARS NORMAL", True))
        if d["diameter_mm"]:
            s.append(seg(f" ({d['diameter_mm']}MM IN DIAMETER)", True))
        s.append(seg(".", True))
    elif d["status"] == "dilated":
        s += [seg("APPENDIX IS DILATED UPTO", True), seg(f" {d['diameter_mm']}MM", True),
              seg(" - ?EVOLVING APPENDICITIS vs PHYSIOLOGICAL.", True)]
    return s


# ============================================================
# IMPRESSION GENERATOR
# ============================================================

def generate_impression(d, sex, age):
    lines = []
    is_pediatric = False
    try:
        is_pediatric = int(age) <= 12 if age else False
    except (ValueError, TypeError):
        pass

    gb, cbd = d["gall_bladder"], d["cbd"]
    if gb["calculi"] == "present" and cbd["status"] == "dilated":
        lines.append("CHOLELITHIASIS WITH CHOLEDOCHOLITHIASIS WITH DILATATION OF CBD"
                     + (f" UPTO {cbd['caliber_mm']}MM" if cbd["caliber_mm"] else "")
                     + (" WITH MILD DILATATION OF PROXIMAL IHBR" if cbd["ihbr_dilated"] else "")
                     + ". Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd["status"] == "dilated" and cbd["calculi"]:
        lines.append(f"CHOLEDOCHOLITHIASIS & DILATED CBD UPTO {cbd['caliber_mm']}MM"
                     + (", WITH MILD DILATATION OF PROXIMAL IHBR." if cbd["ihbr_dilated"] else "."))
    elif gb["calculi"] == "present":
        line = "CHOLELITHIASIS"
        if not gb["pericholecystic_fluid"]:
            line += " WITH NO APPRECIABLE PERICHOLECYSTIC FLUID OR GB WALL THICKENING"
        lines.append(line + ".")

    r, l = d["kidneys"]["right"], d["kidneys"]["left"]
    for side, k in [("RIGHT", r), ("LEFT", l)]:
        if k["calculi"]:
            sizes = ", ".join(f"{c['size_mm']}MM at {c['location']}" for c in k["calculi"])
            hydro = k["calculi"][0].get("hydro", "none")
            line = f"A {side} RENAL CALCULUS ({sizes})"
            if hydro and hydro != "none":
                line += f", CAUSING {side} SIDED {hydro.upper()} HYDROURETERONEPHROSIS"
            else:
                line += ", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN"
            lines.append(line + ".")
        if k["cyst"] != "none":
            bosniak = f" (BOSNIAK CAT-{k['bosniak']})" if k["bosniak"] else ""
            lines.append(f"A {side} RENAL {k['cyst'].upper()} CYST ({k['cyst_size_mm']}MM){bosniak}.")
        if k["nephrocalcinosis"] != "none":
            lines.append(f"MULTIPLE FOCI OF CALCIFICATION SEEN IN THE {side} RENAL CORTEX "
                         f"- LIKELY CORTICAL NEPHROCALCINOSIS.")

    liver = d["liver"]
    liver_line = None
    if liver["size_descriptor"] != "normal":
        liver_line = f"{liver['size_descriptor'].upper()} HEPATOMEGALY"
        if is_pediatric:
            liver_line += " FOR AGE"
    if liver["echotexture"] == "increased":
        steatosis = (f" WITH HEPATIC STEATOSIS({liver['steatosis_grade'].upper()})"
                     if liver["steatosis_grade"] else " WITH HEPATIC STEATOSIS")
        liver_line = (liver_line + steatosis) if liver_line else (
            "HEPATIC STEATOSIS" + (f"({liver['steatosis_grade'].upper()})"
                                    if liver["steatosis_grade"] else ""))
    if liver_line:
        lines.append(liver_line + ". Adv- LFT Correlation.")

    if d["spleen"]["size_descriptor"] != "normal":
        lines.append(f"{d['spleen']['size_descriptor'].upper()} SPLENOMEGALY "
                     f"({d['spleen']['size_mm']}MM).")

    if d["pancreas"]["status"] == "fat_stranding":
        lines.append(f"{d['pancreas']['fat_stranding_grade'].upper()} "
                     f"PERI-PANCREATIC FAT STRANDING.")

    if d["urinary_bladder"]["sedimentation"] in ("free_floating", "significant", "extensive"):
        lines.append("SEDIMENTATION SEEN IN THE UB LUMEN. Adv- Urine R/M Correlation.")

    if sex == "M" and not is_pediatric:
        p = d["prostate"]
        if p["status"] == "bulky":
            lines.append(f"GRADE-I PROSTATOMEGALY ({p['size_cc']}CC).")
        elif p["status"] == "borderline":
            lines.append(f"BORDERLINE PROSTATOMEGALY ({p['size_cc']}CC).")

    if sex == "F" and not is_pediatric:
        if d["uterus"]["status"] == "bulky":
            lines.append("BULKY UTERUS.")
        o = d["ovaries"]
        for side in ("right", "left"):
            if o[f"{side}_status"] == "cyst":
                lines.append(f"A {side.upper()} OVARIAN "
                             f"{o[f'{side}_cyst_type'].upper()} CYST "
                             f"({o[f'{side}_cyst_size']}MM).")

    b = d["bowel"]
    if b["mesenteric_ln"] == "present":
        loc = b["ln_location"].replace("_", " ").upper()
        lines.append(f"MESENTERIC LYMPH NODES IN THE {loc} REGION - ?SIGNIFICANCE. "
                     f"Adv- Lab & Clinical Correlation.")

    if b["free_fluid"] not in ("none",):
        lines.append(f"{b['free_fluid'].replace('_', ' ').upper()} FREE FLUID SEEN.")

    if b["pleural_effusion"] != "none":
        lines.append("PLEURAL EFFUSION SEEN - ?ETIOLOGY.")

    if not lines:
        if sex == "F" and not is_pediatric:
            return list(IMPRESSION_NORMAL_FEMALE)
        return list(IMPRESSION_NORMAL_MALE)

    lines.append(IMPRESSION_REST_UNREMARKABLE)
    return lines


# ============================================================
# DOCX BUILDER  -> returns bytes
# ============================================================

def _set_cell_border_none(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        tcBorders.append(e)
    tcPr.append(tcBorders)


def _add_run(paragraph, text, bold=False, underline=False, font=FONT_BODY,
             size=FONT_SIZE_BODY, italic=False):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline
    run.italic = italic
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rPr.append(rFonts)
    return run


def build_docx_bytes(data):
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(PAGE_TOP_MARGIN)
    section.bottom_margin = Cm(PAGE_BOTTOM_MARGIN)
    section.left_margin = Cm(PAGE_LEFT_MARGIN)
    section.right_margin = Cm(PAGE_RIGHT_MARGIN)

    style = doc.styles["Normal"]
    style.font.name = FONT_BODY
    style.font.size = Pt(FONT_SIZE_BODY)

    table = doc.add_table(rows=2, cols=2)
    table.autofit = True
    p = data["patient"]
    cells = [(0, 0, f"NAME :- {p['name']}", True),
             (0, 1, f"DATE :- {p['date']}", True),
             (1, 0, f"AGE/SEX :- {p['age']}Y/{p['sex']}", True),
             (1, 1, f"REF. BY: {p['referred_by']}", True)]
    for r, c, text, bold in cells:
        cell = table.cell(r, c)
        cell.text = ""
        para = cell.paragraphs[0]
        _add_run(para, text, bold=bold)
        _set_cell_border_none(cell)

    doc.add_paragraph()
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title_para, "ULTRASOUND WHOLE ABDOMEN", bold=True, underline=True,
             size=FONT_SIZE_TITLE)
    doc.add_paragraph()

    sex, age = p["sex"], p["age"]
    is_ped = False
    try:
        is_ped = int(age) <= 12 if age else False
    except (ValueError, TypeError):
        pass

    sections = [
        liver_sentence(data["liver"], sex, age),
        gall_bladder_sentence(data["gall_bladder"]),
        cbd_sentence(data["cbd"]),
        pancreas_sentence(data["pancreas"]),
        spleen_sentence(data["spleen"]),
        kidneys_sentence(data["kidneys"]),
        urinary_bladder_sentence(data["urinary_bladder"]),
    ]
    if sex == "F":
        sections.append(uterus_sentence(data["uterus"], pediatric=is_ped))
        if not is_ped:
            sections.append(ovaries_sentence(data["ovaries"]))
    else:
        sections.append(prostate_sentence(data["prostate"], pediatric=is_ped))
    sections.append(bowel_sentence(data["bowel"], sex))
    if data["appendix"]["status"] != "not_assessed":
        sections.append(appendix_sentence(data["appendix"]))

    for segs in sections:
        if not segs:
            continue
        para = doc.add_paragraph()
        for text, bold, underline in segs:
            if "\n" in text:
                parts = text.split("\n")
                for i, part in enumerate(parts):
                    if i > 0:
                        para.add_run().add_break()
                    _add_run(para, part, bold=bold, underline=underline)
            else:
                _add_run(para, text, bold=bold, underline=underline)
        para.paragraph_format.space_after = Pt(6)

    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True)
    for line in data["impression"]["lines"]:
        para = doc.add_paragraph(style="List Bullet")
        _add_run(para, line, bold=True)

    doc.add_paragraph()
    disc = doc.add_paragraph()
    _add_run(disc, DISCLAIMER_TEXT, bold=True, size=FONT_SIZE_DISCLAIMER)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="Radiology Report Generator", layout="wide")
init_db()

if "report" not in st.session_state:
    st.session_state.report = new_report("F")
if "impression_override" not in st.session_state:
    st.session_state.impression_override = ""


st.title("USG Whole Abdomen — Report Generator")

# ---- Patient header ----
with st.container():
    c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
    with c1:
        p_name = st.text_input("Name", value=st.session_state.report["patient"]["name"])
    with c2:
        p_age = st.text_input("Age", value=st.session_state.report["patient"]["age"])
    with c3:
        p_sex = st.radio("Sex", ["F", "M"], horizontal=True,
                         index=0 if st.session_state.report["patient"]["sex"] == "F" else 1)
    with c4:
        p_date = st.text_input("Date", value=st.session_state.report["patient"]["date"])
    with c5:
        p_ref = st.text_input("Referred by", value=st.session_state.report["patient"]["referred_by"])

st.markdown("---")

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("Findings")

    with st.expander("LIVER", expanded=True):
        c1, c2 = st.columns([1, 2])
        with c1:
            liver_size = st.text_input("Size (mm)", value=st.session_state.report["liver"]["size_mm"], key="liver_size")
        with c2:
            liver_status = st.radio("Status", ["normal", "mild", "moderate", "severe"],
                                    horizontal=True,
                                    index=["normal", "mild", "moderate", "severe"].index(
                                        st.session_state.report["liver"]["size_descriptor"]))
        liver_echo = st.radio("Echotexture", ["normal", "increased", "coarse"], horizontal=True,
                              format_func=lambda x: {"normal": "Normal", "increased": "Increased (steatosis)",
                                                     "coarse": "Coarse"}[x],
                              index=["normal", "increased", "coarse"].index(
                                  st.session_state.report["liver"]["echotexture"]))
        liver_steatosis = st.selectbox("Steatosis grade (if increased reflectivity)",
                                       ["", "Mild", "Mild+", "Mild-Moderate", "Moderate", "Severe"],
                                       index=0)
        liver_focal = st.radio("Focal lesion", ["none", "calcified"], horizontal=True,
                               format_func=lambda x: {"none": "None",
                                                      "calcified": "Calcified focus (right lobe)"}[x])
        liver_portal = st.radio("Portal vein", ["normal", "dilated"], horizontal=True)

    with st.expander("GALL BLADDER", expanded=False):
        gb_status = st.radio("Status", ["adequately_distended", "over", "partially", "contracted", "empty", "operated"],
                             horizontal=True,
                             format_func=lambda x: x.replace("_", " ").title())
        gb_calc = st.radio("Calculi", ["none", "present"], horizontal=True)
        if gb_calc == "present":
            c1, c2, c3 = st.columns(3)
            with c1:
                gb_count = st.selectbox("Count", ["single", "couple", "multiple", "multiple_tiny", "few"])
            with c2:
                gb_size = st.text_input("Largest (mm)")
            with c3:
                gb_wall = st.text_input("Wall (mm)")
        else:
            gb_count, gb_size, gb_wall = "multiple", "", ""
        gb_sludge = st.radio("Sludge", ["none", "trace", "sludge", "sludge_ball"], horizontal=True)

    with st.expander("COMMON BILE DUCT", expanded=False):
        c1, c2 = st.columns([1, 2])
        with c1:
            cbd_mm = st.text_input("Caliber (mm)")
        with c2:
            cbd_status = st.radio("Status", ["normal", "dilated"], horizontal=True)
        cbd_calc = st.checkbox("Calculus in CBD")
        cbd_calc_size = st.text_input("Largest calculus (mm)") if cbd_calc else ""
        cbd_ihbr = st.checkbox("IHBR dilated")

    with st.expander("PANCREAS", expanded=False):
        pn_status = st.radio("Status", ["normal", "fat_stranding", "necrotic_ln"], horizontal=True,
                             format_func=lambda x: x.replace("_", " ").title())

    with st.expander("SPLEEN", expanded=False):
        c1, c2 = st.columns([1, 2])
        with c1:
            sp_size = st.text_input("Size (mm)")
        with c2:
            sp_desc = st.radio("Status", ["normal", "borderline", "mild", "moderate"], horizontal=True)

    with st.expander("KIDNEYS", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Right Kidney**")
            kd_r_status = st.radio("Status", ["normal", "not_visualized"], horizontal=True, key="kd_r")
            kd_r_calc = st.text_area("Calculi (e.g. '4.6 upper-mid, 4.1 lower-mid')", height=60)
        with c2:
            st.markdown("**Left Kidney**")
            kd_l_status = st.radio("Status", ["normal", "not_visualized"], horizontal=True, key="kd_l")
            kd_l_calc = st.text_area("Calculi (e.g. '5.2 lower, 4.3 mid')", height=60)

    with st.expander("URINARY BLADDER", expanded=False):
        ub_status = st.radio("Status", ["adequately_distended", "over", "partially", "empty"],
                             horizontal=True, format_func=lambda x: x.replace("_", " ").title())
        ub_sed = st.radio("Sedimentation", ["none", "trace", "free_floating", "significant", "extensive"],
                          horizontal=True, format_func=lambda x: x.replace("_", " ").title())

    if p_sex == "F":
        with st.expander("UTERUS", expanded=False):
            ut_status = st.radio("Status", ["anteverted", "retroverted", "bulky", "operated", "not_visualized"],
                                 horizontal=True)
            c1, c2 = st.columns(2)
            with c1:
                ut_size = st.text_input("Size (mm, e.g. 72X25)")
            with c2:
                ut_et = st.text_input("ET (mm)")
        with st.expander("OVARIES", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Right Ovary**")
                ov_r = st.radio("Status", ["normal", "cyst", "not_visualized"], horizontal=True, key="ov_r")
                ov_r_size = st.text_input("Size (mm)", key="ov_r_size")
            with c2:
                st.markdown("**Left Ovary**")
                ov_l = st.radio("Status", ["normal", "cyst", "not_visualized"], horizontal=True, key="ov_l")
                ov_l_size = st.text_input("Size (mm)", key="ov_l_size")
    else:
        with st.expander("PROSTATE", expanded=False):
            c1, c2 = st.columns([1, 2])
            with c1:
                pr_cc = st.text_input("Size (cc)")
            with c2:
                pr_status = st.radio("Status", ["normal", "borderline", "bulky", "grade1"],
                                     horizontal=True,
                                     format_func=lambda x: {"normal": "Normal", "borderline": "Borderline",
                                                            "bulky": "Bulky", "grade1": "Grade-I BPH"}[x])

    with st.expander("BOWEL / FREE FLUID", expanded=False):
        bw_ff = st.radio("Free fluid", ["none", "minimal", "mild", "moderate"], horizontal=True)
        bw_ln = st.radio("Mesenteric LN", ["none", "present"], horizontal=True)

    with st.expander("APPENDIX (optional)", expanded=False):
        ap_status = st.radio("Status", ["not_assessed", "not_visualized", "normal", "dilated"],
                             horizontal=True, format_func=lambda x: x.replace("_", " ").title())
        ap_d = st.text_input("Diameter (mm)") if ap_status in ("normal", "dilated") else ""


# ---- Build data dict from widgets ----
def parse_calc(text):
    if not text.strip():
        return []
    out = []
    for part in text.split(","):
        part = part.strip()
        tokens = part.replace("mm", " ").split()
        size, loc = "", "mid"
        for t in tokens:
            if t.replace(".", "").isdigit():
                size = t
            elif any(p in t.lower() for p in ["upper", "mid", "lower", "pole"]):
                loc = t
        out.append({"size_mm": size, "location": loc, "hydro": "none"})
    return out


data = new_report(p_sex)
data["patient"] = {"name": p_name, "age": p_age, "sex": p_sex,
                   "date": p_date, "referred_by": p_ref}

data["liver"].update({"size_mm": liver_size, "size_descriptor": liver_status,
                      "echotexture": liver_echo,
                      "steatosis_grade": liver_steatosis if liver_steatosis else None,
                      "focal_lesion": liver_focal, "portal_vein": liver_portal})

data["gall_bladder"].update({"status": gb_status, "calculi": gb_calc,
                             "calculi_count": gb_count, "calculi_largest_mm": gb_size,
                             "wall_mm": gb_wall,
                             "wall": "thickened" if gb_wall else "normal",
                             "sludge": gb_sludge})

data["cbd"].update({"caliber_mm": cbd_mm, "status": cbd_status,
                    "calculi": cbd_calc, "calculi_size_mm": cbd_calc_size,
                    "ihbr_dilated": cbd_ihbr})

data["pancreas"]["status"] = pn_status
data["spleen"].update({"size_mm": sp_size, "size_descriptor": sp_desc})
data["kidneys"]["right"].update({"status": kd_r_status, "calculi": parse_calc(kd_r_calc)})
data["kidneys"]["left"].update({"status": kd_l_status, "calculi": parse_calc(kd_l_calc)})
data["urinary_bladder"].update({"status": ub_status, "sedimentation": ub_sed})

if p_sex == "F":
    data["uterus"].update({"status": ut_status, "size": ut_size,
                           "endometrial_thickness_mm": ut_et})
    data["ovaries"].update({"right_status": ov_r, "right_size": ov_r_size,
                            "left_status": ov_l, "left_size": ov_l_size})
else:
    data["prostate"].update({"status": pr_status, "size_cc": pr_cc})

data["bowel"].update({"free_fluid": bw_ff, "mesenteric_ln": bw_ln})
data["appendix"].update({"status": ap_status, "diameter_mm": ap_d})

# ---- Auto-generate impression ----
auto_impression = generate_impression(data, p_sex, p_age)
data["impression"]["lines"] = auto_impression

# ---- Preview ----
def render_preview(data):
    p = data["patient"]
    out = []
    out.append(f"NAME :- {p['name']}               DATE :- {p['date']}")
    out.append(f"AGE/SEX :- {p['age']}Y/{p['sex']}          REF. BY: {p['referred_by']}")
    out.append("")
    out.append("         ULTRASOUND WHOLE ABDOMEN")
    out.append("")
    sex, age = p["sex"], p["age"]
    is_ped = False
    try:
        is_ped = int(age) <= 12 if age else False
    except (ValueError, TypeError):
        pass
    secs = [liver_sentence(data["liver"], sex, age),
            gall_bladder_sentence(data["gall_bladder"]),
            cbd_sentence(data["cbd"]),
            pancreas_sentence(data["pancreas"]),
            spleen_sentence(data["spleen"]),
            kidneys_sentence(data["kidneys"]),
            urinary_bladder_sentence(data["urinary_bladder"])]
    if sex == "F":
        secs.append(uterus_sentence(data["uterus"], pediatric=is_ped))
        if not is_ped:
            secs.append(ovaries_sentence(data["ovaries"]))
    else:
        secs.append(prostate_sentence(data["prostate"], pediatric=is_ped))
    secs.append(bowel_sentence(data["bowel"], sex))
    if data["appendix"]["status"] != "not_assessed":
        secs.append(appendix_sentence(data["appendix"]))
    for s in secs:
        if not s:
            continue
        out.append("".join(x[0] for x in s))
        out.append("")
    out.append("IMPRESSION:")
    for line in data["impression"]["lines"]:
        out.append(f"  - {line}")
    out.append("")
    out.append(DISCLAIMER_TEXT)
    return "\n".join(out)


with col_right:
    st.subheader("Live Preview")
    st.text_area("Report preview", value=render_preview(data), height=700, disabled=True,
                 label_visibility="collapsed")

    st.subheader("Impression (editable)")
    edited_imp = st.text_area("Edit impression lines (one per line). Leave blank to use auto-generated.",
                              value="\n".join(auto_impression), height=180, key="imp_edit")
    if edited_imp.strip():
        data["impression"]["lines"] = [ln.strip() for ln in edited_imp.splitlines() if ln.strip()]
    else:
        data["impression"]["lines"] = auto_impression


# ---- Download / Save ----
st.markdown("---")
c1, c2, c3 = st.columns([1, 1, 4])
with c1:
    docx_bytes = build_docx_bytes(data)
    fname = f"{p_name or 'report'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    fname = "".join(ch for ch in fname if ch.isalnum() or ch in "._-")
    st.download_button("Download .docx", docx_bytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
with c2:
    if st.button("Save to Database"):
        if not p_name.strip():
            st.warning("Enter patient name first.")
        else:
            save_report(data)
            if p_ref.strip():
                add_referrer(p_ref)
            st.success("Saved to local database.")
