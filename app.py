"""
Radiology Report Generator — USG Whole Abdomen
Streamlit version (Phase 1)
Modules completed: Liver, Gall Bladder, CBD, Pancreas, Spleen, Kidneys, UB, Uterus/Ovaries/Prostate, Bowel, Appendix
"""

import io
import os
import re
import sqlite3
from datetime import datetime

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# CONFIG
# ============================================================

DB_PATH = os.path.join(os.path.expanduser("~"), "RadiologyReports", "reports.db")

PAGE_TOP_MARGIN = 4.5
PAGE_BOTTOM_MARGIN = 3.5
PAGE_LEFT_MARGIN = 2.0
PAGE_RIGHT_MARGIN = 2.0

FONT_BODY = "Calibri"
FONT_SIZE_BODY = 11
FONT_SIZE_TITLE = 13
FONT_SIZE_DISCLAIMER = 10

TITLE_COLOR = RGBColor(0x1F, 0x4E, 0x79)

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
# PEDIATRIC TABLES (mm)
# ============================================================

PEDIATRIC_SPLEEN_MAX_MM = {
    (0.0, 0.25, "F"): 55,  (0.0, 0.25, "M"): 68,
    (0.25, 0.5, "F"): 56,  (0.25, 0.5, "M"): 70,
    (0.5, 1.0, "F"): 75,   (0.5, 1.0, "M"): 74,
    (1.0, 2.0, "F"): 82,   (1.0, 2.0, "M"): 83,
    (2.0, 4.0, "F"): 89,   (2.0, 4.0, "M"): 99,
    (4.0, 6.0, "F"): 95,   (4.0, 6.0, "M"): 99,
    (6.0, 8.0, "F"): 100,  (6.0, 8.0, "M"): 105,
    (8.0, 10.0, "F"): 105, (8.0, 10.0, "M"): 112,
    (10.0, 12.0, "F"): 114,(10.0, 12.0, "M"): 113,
    (12.0, 14.0, "F"): 116,(12.0, 14.0, "M"): 117,
    (14.0, 18.0, "F"): 110,(14.0, 18.0, "M"): 125,
}

PEDIATRIC_LIVER_MAX_MM = {
    (0.0, 0.25): 90, (0.25, 0.5): 95, (0.5, 0.75): 100, (0.75, 1.0): 100,
    (1.0, 2.5): 105, (2.5, 3.0): 105, (3.0, 5.0): 115, (5.0, 7.0): 125,
    (7.0, 9.0): 130, (9.0, 11.0): 135, (11.0, 13.0): 140, (13.0, 15.0): 140,
    (15.0, 18.0): 145,
}


def parse_age(age_text):
    if age_text is None:
        return None
    txt = str(age_text).strip().upper().replace("Y", "").replace("M", "").strip()
    try:
        return float(txt)
    except (ValueError, TypeError):
        return None


def get_pediatric_liver_max_mm(age_years):
    for (lo, hi), max_mm in PEDIATRIC_LIVER_MAX_MM.items():
        if lo <= age_years < hi:
            return max_mm
    return None


def get_pediatric_spleen_max_mm(age_years, sex):
    for (lo, hi, s), max_mm in PEDIATRIC_SPLEEN_MAX_MM.items():
        if s == sex and lo <= age_years < hi:
            return max_mm
    return None


def classify_liver_adult(size_mm):
    if size_mm < 150: return "normal"
    elif size_mm < 151: return "borderline"
    elif size_mm <= 170: return "mild"
    elif size_mm <= 190: return "moderate"
    else: return "gross"


def classify_spleen_adult(size_mm):
    if size_mm < 120: return "normal"
    elif size_mm < 130: return "borderline"
    elif size_mm <= 145: return "mild"
    else: return "moderate"


def auto_classify_liver(size_text, age_text, sex):
    try:
        size = float(str(size_text).replace("MM", "").strip())
    except (ValueError, TypeError):
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return "normal"
    if age < 18:
        max_mm = get_pediatric_liver_max_mm(age)
        if max_mm is None:
            return "normal"
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_liver_adult(size)


def auto_classify_spleen(size_text, age_text, sex):
    try:
        size = float(str(size_text).replace("MM", "").strip())
    except (ValueError, TypeError):
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return "normal"
    if age < 18:
        max_mm = get_pediatric_spleen_max_mm(age, sex)
        if max_mm is None:
            return "normal"
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_spleen_adult(size)


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
    if not name.strip(): return
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO referrers (name) VALUES (?)", (name.strip(),))
    conn.commit()
    conn.close()


def save_report(data):
    import json
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""INSERT INTO reports
                 (patient_name, age, sex, report_date, referred_by, created_at, data_json)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
              (data["patient"]["name"], data["patient"]["age"], data["patient"]["sex"],
               data["patient"]["date"], data["patient"]["referred_by"],
               datetime.now().isoformat(), json.dumps(data)))
    conn.commit()
    conn.close()


# ============================================================
# DATA MODEL
# ============================================================

def new_report(sex="F"):
    return {
        "patient": {"name": "", "age": "", "sex": sex,
                    "date": datetime.now().strftime("%d-%b-%y"), "referred_by": ""},
        "liver": {"size_mm": "", "size_descriptor": "normal", "outline": "normal",
                  "echotexture": "normal", "steatosis_grade": None,
                  "focal_lesion": "none", "focal_lesion_text": "",
                  "cyst_count": "single", "cyst_single_lobe": "right",
                  "cyst_single_size_mm": "", "cyst_few_largest_mm": "",
                  "cyst_few_lobe": "right",
                  "hemangioma_count": "single", "hemangioma_single_lobe": "right",
                  "hemangioma_single_size_mm": "", "hemangioma_few_largest_mm": "",
                  "hemangioma_few_lobe": "right",
                  "abscess_count": "single", "abscess_lesions": [],
                  "ihbr": "normal", "portal_vein": "normal", "portal_vein_mm": ""},
        "gall_bladder": {
            "status": "adequately_distended",
            "wall_thickened": False, "wall_mm": "",
            "calculi": "none", "calculi_count": "single",
            "calculi_size_cat": "small", "calculi_size_mm": "",
            "calculi_neck": False, "calculi_neck_size_mm": "",
            "sludge": "none",
            "sludge_ball": False, "sludge_ball_count": "single",
            "sludge_ball_size_mm": "", "sludge_ball_wall": "anterior",
            "comet_tail": False, "comet_tail_count": "single",
            "comet_tail_wall": "anterior",
            "pericholecystic_fluid": False,
        },
        "cbd": {"caliber_mm": "", "status": "normal", "calculi": False,
                "calculi_size_mm": "", "calculi_location": "distal", "ihbr_dilated": False},
        "pancreas": {"status": "normal", "fat_stranding_grade": "mild", "ln_size": "",
                     "mpd_dilated": False, "mpd_mm": ""},
        "spleen": {"size_mm": "", "size_descriptor": "normal", "portal_vein_mm": ""},
        "kidneys": {
            "right": {"status": "normal", "calculi": [], "cyst": "none",
                      "cyst_size_mm": "", "cyst_location": "", "bosniak": "",
                      "hydronephrosis": "none", "nephrocalcinosis": "none"},
            "left": {"status": "normal", "calculi": [], "cyst": "none",
                     "cyst_size_mm": "", "cyst_location": "", "bosniak": "",
                     "hydronephrosis": "none", "nephrocalcinosis": "none"},
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


def liver_focal_sentence(d):
    fl = d["focal_lesion"]
    if fl == "none":
        return [seg(" No focal lesion is seen.")]
    if fl == "calcified":
        return [seg(" "), seg("A CALCIFIED FOCUS SEEN IN THE RIGHT HEPATIC LOBE.", True)]
    if fl == "other" and d.get("focal_lesion_text"):
        return [seg(" "), seg(d["focal_lesion_text"], True)]
    if fl == "cyst":
        count = d.get("cyst_count", "single")
        if count == "single":
            lobe = d.get("cyst_single_lobe", "right").lower()
            size = d.get("cyst_single_size_mm", "")
            text = f"A simple cyst({size} mm) seen in the {lobe} lobe of liver."
            return [seg(" "), seg(text, True)]
        else:
            size = d.get("cyst_few_largest_mm", "")
            lobe = d.get("cyst_few_lobe", "right")
            text = (f"Few simple cysts seen in the liver, largest of these "
                    f"measuring {size}mm in {lobe} lobe.")
            return [seg(" "), seg(text, True)]
    if fl == "hemangioma":
        count = d.get("hemangioma_count", "single")
        if count == "single":
            lobe = d.get("hemangioma_single_lobe", "right")
            size = d.get("hemangioma_single_size_mm", "")
            text = f"A hyperechoic small SOL({size}mm) seen in the {lobe} hepatic Lobe of liver."
            return [seg(" "), seg(text, True)]
        else:
            size = d.get("hemangioma_few_largest_mm", "")
            lobe = d.get("hemangioma_few_lobe", "right")
            text = (f"Few hyperechoic small SOLs seen in the liver, largest of these "
                    f"measuring {size}mm in {lobe} lobe.")
            return [seg(" "), seg(text, True)]
    if fl == "abscess":
        count = d.get("abscess_count", "single")
        lesions = d.get("abscess_lesions", [])
        if not lesions:
            return []
        if count == "single" and len(lesions) >= 1:
            l = lesions[0]
            text = (f"An irregular marginated ill-defined avascular SOL({l['dim']}; "
                    f"Vol= {l['vol']}cc) seen in the Segment {l['segment']}.")
            return [seg(" "), seg(text, True)]
        else:
            parts = [f"{l['dim']}; Vol= {l['vol']}cc in Segment {l['segment']}"
                     for l in lesions]
            joined = " & ".join(parts)
            keyword = "Few" if count == "few" else "Multiple"
            text = (f"{keyword} irregular marginated ill-defined avascular SOLs seen "
                    f"in the Liver, Largest of these measuring {joined}.")
            return [seg(" "), seg(text, True)]
    return []


def liver_sentence(d, sex, age):
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    outline = d.get("outline", "normal")
    echo = d["echotexture"]
    grade = d.get("steatosis_grade")
    s = [seg("LIVER", True, True)]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}MM)"))
    elif desc == "borderline":
        s += [seg(" is "), seg("BORDERLINE ENLARGED IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "mild":
        s += [seg(" is "), seg("MILDLY ENLARGED IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("MODERATELY ENLARGED IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "gross":
        s += [seg(" is "), seg("GROSSLY ENLARGED IN SIZE", True), seg(f" ({size}MM)")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("ENLARGED FOR AGE IN SIZE", True), seg(f" ({size}MM)")]

    if outline == "crenated":
        s.append(seg(", "))
        s.append(seg("CRENATED/NODULAR OUTLINE AND", True))
        if echo == "coarse":
            s.append(seg(" COARSE ECHOTEXTURE", True))
        elif echo == "increased":
            if grade == "Severe+++":
                s.append(seg(" SIGNIFICANT FATTY INFILTRATION", True))
            else:
                s.append(seg(" INCREASED REFLECTIVITY", True))
        elif echo == "low":
            s.append(seg(" LOW ECHOTEXTURE", True))
        else:
            s.append(seg(" NORMAL ECHOTEXTURE", True))
        s.append(seg("."))
    else:
        s.append(seg(" with normal outline and "))
        if echo == "normal":
            s.append(seg("echotexture."))
        elif echo == "increased":
            if grade == "Severe+++":
                s.append(seg("SIGNIFICANT FATTY INFILTRATION", True))
            else:
                s.append(seg("INCREASED REFLECTIVITY", True))
            s.append(seg("."))
        elif echo == "coarse":
            s.append(seg("COARSE ECHOTEXTURE", True)); s.append(seg("."))
        elif echo == "low":
            s.append(seg("LOW ECHOTEXTURE", True)); s.append(seg("."))

    s.extend(liver_focal_sentence(d))

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
        s.append(seg(" is partially contracted (suboptimal wall visualization)."))
    elif status == "contracted":
        s.append(seg(" is contracted (suboptimal wall visualization)."))
    elif status == "empty":
        s.append(seg(" is empty."))
    elif status == "operated":
        s += [seg(" is operated. "), seg("GB FOSSA", True, True), seg(" is unremarkable.")]
        return s

    # Wall (skip if contracted)
    if status != "contracted":
        if d.get("wall_thickened") and d.get("wall_mm"):
            try:
                w = float(d["wall_mm"])
            except ValueError:
                w = 0
            if w <= 8:
                s += [seg(" "), seg(f"WALL IS MILDLY THICKENED UPTO {d['wall_mm']}MM.", True)]
            else:
                s += [seg(" "), seg(f"WALL IS SIGNIFICANTLY THICKENED UPTO {d['wall_mm']}MM.", True)]
        else:
            s.append(seg(" Wall thickness is normal."))

    # Calculi
    if d.get("calculi") == "present":
        count = d.get("calculi_count", "single")
        neck = d.get("calculi_neck", False) and count != "innumerable"
        neck_size = d.get("calculi_neck_size_mm", "")
        size = d.get("calculi_size_mm", "")
        size_cat = d.get("calculi_size_cat", "small").lower()

        if count == "innumerable":
            s += [seg(" "), seg("Innumerable tiny calculi seen in the GB lumen.", True)]
        elif count == "single":
            if neck:
                s += [seg(" "), seg(f"A calculus measuring {size}mm, seen impacted at the GB neck.", True)]
            else:
                s += [seg(" "), seg(f"A calculus measuring {size}mm is seen in the GB lumen.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            if neck:
                s += [seg(" "), seg(f"{word} {size_cat} calculi seen in the GB lumen, "
                                     f"with a calculus measuring {neck_size}mm impacted at the GB neck.",
                                     True)]
            else:
                s += [seg(" "), seg(f"{word} {size_cat} calculi seen in the GB lumen, "
                                     f"largest of these measuring {size}mm.", True)]
    else:
        s.append(seg(" No obvious gall stones seen."))

    # Sludge (non-ball)
    sludge = d.get("sludge", "none")
    if sludge != "none":
        s += [seg(" "), seg(f"{sludge.upper()} SLUDGE SEEN IN THE GALLBLADDER LUMEN.", True)]

    # Sludge ball / polyp
    if d.get("sludge_ball"):
        count = d.get("sludge_ball_count", "single")
        size = d.get("sludge_ball_size_mm", "")
        wall = d.get("sludge_ball_wall", "anterior")
        if count == "single":
            s += [seg(" "), seg(f"A sludge ball/polyp({size}mm) seen impacted at the "
                                 f"{wall} GB wall.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            s += [seg(" "), seg(f"{word} sludge balls/polyps seen impacted at the "
                                 f"{wall} GB wall, largest of these measuring {size}mm.", True)]

    # Comet tail / adenomyomatosis
    if d.get("comet_tail"):
        count = d.get("comet_tail_count", "single")
        wall = d.get("comet_tail_wall", "anterior")
        if count == "single":
            s += [seg(" "), seg(f"A comet tail artifact seen arising from the {wall} GB wall.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            s += [seg(" "), seg(f"{word} comet tail artifacts seen arising from the {wall} GB wall.", True)]

    # Pericholecystic fluid
    if d.get("pericholecystic_fluid"):
        s += [seg(" "), seg("THIN RIM OF PERICHOLECYSTIC FLUID SEEN.", True)]

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
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("ENLARGED FOR AGE IN SIZE", True),
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
            b += [seg(" "), seg(f"A {k['cyst'].upper()} CYST MEASURING {k['cyst_size_mm']}MM "
                                f"IS SEEN AT THE {k['cyst_location'].upper()} OF {side_label} "
                                f"KIDNEY{bosniak}", True), seg(".", True)]
        if k["hydronephrosis"] != "none":
            b += [seg(" "), seg(f"{k['hydronephrosis'].upper()} HYDROURETERONEPHROSIS IS PRESENT "
                                f"ON THE {side_label}", True), seg(".", True)]
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
    s = [seg("UT
