"""
Radiology Report Generator — USG Whole Abdomen
v1.2.1

Progress:
- Liver, Gall Bladder, CBD: complete
- Pancreas: Normal + Early/evolving + Acute
- Fixes: combined fat+fluid, steato-hepatitis spelling, liver combination,
  acute pancreatitis bowel/ascites handling, text wrapping,
  justify alignment, non-breaking hyphen in STEATO-HEPATITIS

Walled-off necrosis, pancreatic pseudocyst, chronic pancreatitis: pending.
"""

import io
import os
import re
import html
import hashlib
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

PAGE_TOP_MARGIN = 4.0
PAGE_BOTTOM_MARGIN = 3.5
PAGE_LEFT_MARGIN = 2.0
PAGE_RIGHT_MARGIN = 2.0

FONT_BODY = "Calibri"
FONT_SIZE_BODY = 11
FONT_SIZE_TITLE = 13
FONT_SIZE_DISCLAIMER = 8

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
    txt = str(age_text).strip()
    if not txt:
        return None
    match = re.search(r"(\d+(\.\d+)?)", txt)
    if not match:
        return None
    try:
        return float(match.group(1))
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
    if size_mm < 150:
        return "normal"
    elif size_mm < 151:
        return "borderline"
    elif size_mm <= 170:
        return "mild"
    elif size_mm <= 190:
        return "moderate"
    else:
        return "gross"


def classify_spleen_adult(size_mm):
    if size_mm < 120:
        return "normal"
    elif size_mm < 130:
        return "borderline"
    elif size_mm <= 145:
        return "mild"
    else:
        return "moderate"


def auto_classify_liver(size_mm, age_text, sex):
    try:
        size = float(size_mm)
    except (ValueError, TypeError):
        return "normal"
    if size <= 0:
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return classify_liver_adult(size)
    if age < 18:
        max_mm = get_pediatric_liver_max_mm(age)
        if max_mm is None:
            return classify_liver_adult(size)
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_liver_adult(size)


def auto_classify_spleen(size_mm, age_text, sex):
    try:
        size = float(size_mm)
    except (ValueError, TypeError):
        return "normal"
    if size <= 0:
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return classify_spleen_adult(size)
    if age < 18:
        max_mm = get_pediatric_spleen_max_mm(age, sex)
        if max_mm is None:
            return classify_spleen_adult(size)
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_spleen_adult(size)


LIVER_STATUS_LABELS = {
    "normal": "Normal size", "borderline": "Borderline enlarged",
    "mild": "Mildly enlarged", "moderate": "Moderately enlarged",
    "gross": "Grossly enlarged", "enlarged_for_age": "Enlarged for age",
}
SPLEEN_STATUS_LABELS = {
    "normal": "Normal size", "borderline": "Borderline enlarged",
    "mild": "Mildly enlarged", "moderate": "Moderately enlarged",
    "enlarged_for_age": "Enlarged for age",
}


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
            "comet_tail_wall": "anterior", "pericholecystic_fluid": False,
        },
        "cbd": {"size_mm": "", "status": "normal", "calculi": False,
                "calculi_count": "single", "calculi_size_mm": "",
                "calculi_location": "distal", "ihbr": "normal"},
        "pancreas": {
            "status": "normal",
            "ee_size": "normal",
            "ee_fat_stranding": False,
            "ee_fat_location": "none",
            "ee_free_fluid": False,
            "ee_fluid_location": "none",
            "ac_size": "normal",
            "ac_echo": "normal",
            "ac_echo_location": "none",
            "ac_margins": "normal",
        },
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


# -------- LIVER --------

def liver_focal_sentence(d):
    fl = d["focal_lesion"]
    if fl == "none":
        return [seg(" No focal lesion is seen.")]
    if fl == "calcified":
        return [seg(" "), seg("A calcified focus seen in the right hepatic lobe.", True)]
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
            text = f"A hyperechoic small SOL({size}mm) seen in the {lobe} hepatic lobe of liver."
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
                    f"Vol= {l['vol']}cc) seen in the segment {l['segment']}.")
            return [seg(" "), seg(text, True)]
        else:
            parts = [f"{l['dim']}; Vol= {l['vol']}cc in segment {l['segment']}"
                     for l in lesions]
            joined = " & ".join(parts)
            keyword = "Few" if count == "few" else "Multiple"
            text = (f"{keyword} irregular marginated ill-defined avascular SOLs seen "
                    f"in the liver, largest of these measuring {joined}.")
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
        s += [seg(" is "), seg("borderline enlarged in size", True), seg(f" ({size}MM)")]
    elif desc == "mild":
        s += [seg(" is "), seg("mildly enlarged in size", True), seg(f" ({size}MM)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("moderately enlarged in size", True), seg(f" ({size}MM)")]
    elif desc == "gross":
        s += [seg(" is "), seg("grossly enlarged in size", True), seg(f" ({size}MM)")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("enlarged for age in size", True), seg(f" ({size}MM)")]

    if outline == "crenated":
        s.append(seg(", "))
        s.append(seg("crenated/nodular outline and", True))
        if echo == "coarse":
            s.append(seg(" coarse echotexture", True))
        elif echo == "increased":
            if grade == "Severe+++":
                s.append(seg(" significant fatty infiltration", True))
            else:
                s.append(seg(" increased reflectivity", True))
        elif echo == "low":
            s.append(seg(" low echotexture", True))
        else:
            s.append(seg(" normal echotexture", True))
        s.append(seg("."))
    else:
        s.append(seg(" with normal outline and "))
        if echo == "normal":
            s.append(seg("echotexture."))
        elif echo == "increased":
            if grade == "Severe+++":
                s.append(seg("significant fatty infiltration", True))
            else:
                s.append(seg("increased reflectivity", True))
            s.append(seg("."))
        elif echo == "coarse":
            s.append(seg("coarse echotexture", True))
            s.append(seg("."))
        elif echo == "low":
            s.append(seg("low echotexture", True))
            s.append(seg("."))

    s.extend(liver_focal_sentence(d))

    if d["ihbr"] == "normal":
        s.append(seg(" Intra hepatic biliary radicals are normal."))
    else:
        s += [seg(" Intra hepatic biliary radicals are "), seg("dilated", True), seg(".")]
    if d["portal_vein"] == "normal":
        s.append(seg(" Portal vein is normal in course and caliber."))
    else:
        s += [seg(" Portal vein is "), seg("dilated", True)]
        if d["portal_vein_mm"]:
            s.append(seg(f" ({d['portal_vein_mm']}MM)"))
        s.append(seg("."))
    return s


# -------- GALL BLADDER --------

def gall_bladder_sentence(d):
    s = [seg("GALL BLADDER", True, True)]
    status = d["status"]
    if status == "adequately_distended":
        s.append(seg(" is adequately distended."))
    elif status == "over":
        s += [seg(" is "), seg("over-distended", True), seg(".")]
    elif status == "partially":
        s.append(seg(" is partially contracted (suboptimal wall visualization)."))
    elif status == "contracted":
        s.append(seg(" is contracted (suboptimal wall visualization)."))
    elif status == "empty":
        s.append(seg(" is empty."))
    elif status == "operated":
        s += [seg(" is operated. "), seg("GB FOSSA", True, True), seg(" is unremarkable.")]
        return s

    if status != "contracted":
        if d.get("wall_thickened") and d.get("wall_mm"):
            try:
                w = float(d["wall_mm"])
            except ValueError:
                w = 0
            if w <= 8:
                s += [seg(" "), seg(f"wall is mildly thickened upto {d['wall_mm']}MM.", True)]
            else:
                s += [seg(" "), seg(f"wall is significantly thickened upto {d['wall_mm']}MM.", True)]
        else:
            s.append(seg(" Wall thickness is normal."))

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
                                     f"with a calculus measuring {neck_size}mm impacted "
                                     f"at the GB neck.", True)]
            else:
                s += [seg(" "), seg(f"{word} {size_cat} calculi seen in the GB lumen, "
                                     f"largest of these measuring {size}mm.", True)]
    else:
        s.append(seg(" No obvious gall stones seen."))

    sludge = d.get("sludge", "none")
    if sludge != "none":
        s += [seg(" "), seg(f"{sludge.title()} sludge seen in the gallbladder lumen.", True)]

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

    if d.get("comet_tail"):
        count = d.get("comet_tail_count", "single")
        wall = d.get("comet_tail_wall", "anterior")
        if count == "single":
            s += [seg(" "), seg(f"A comet tail artifact seen arising from the {wall} GB wall.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            s += [seg(" "), seg(f"{word} comet tail artifacts seen arising from the "
                                 f"{wall} GB wall.", True)]

    if d.get("pericholecystic_fluid"):
        s += [seg(" "), seg("Thin rim of pericholecystic fluid seen.", True)]
    return s


# -------- CBD --------

def cbd_location_word(loc):
    return {
        "proximal": "proximal segment", "mid": "mid segment",
        "distal": "distal segment", "mid_distal": "mid/distal segment",
    }.get(loc, "distal segment")


def cbd_sentence(d):
    s = [seg("COMMON BILE DUCT", True, True)]
    status = d.get("status", "normal")
    size = d.get("size_mm", "")
    calculi = d.get("calculi", False)
    calculi_count = d.get("calculi_count", "single")
    calculi_size = d.get("calculi_size_mm", "")
    loc_word = cbd_location_word(d.get("calculi_location", "distal"))
    ihbr = d.get("ihbr", "normal")

    if status == "normal":
        if size:
            s.append(seg(f" is normal in caliber({size}MM)"))
        else:
            s.append(seg(" is normal in caliber"))
    elif status == "proximal":
        s.append(seg(" is "))
        s.append(seg("proximally dilated in caliber", True))
        if size:
            s.append(seg(f" upto {size}MM", True))
    elif status == "dilated":
        s.append(seg(" is "))
        s.append(seg("dilated in caliber", True))
        if size:
            s.append(seg(f" upto {size}MM", True))

    if calculi and calculi_size:
        if calculi_count == "single":
            s.append(seg(f", with suggestion of a calculus measuring {calculi_size}MM "
                         f"in the {loc_word}", True))
        else:
            s.append(seg(f", with suggestion of few calculi in the {loc_word}, "
                         f"largest of these measuring {calculi_size}MM", True))
    s.append(seg("."))

    show_ihbr = (status in ("proximal", "dilated")) or calculi
    if show_ihbr:
        if ihbr == "normal":
            if status == "normal" and calculi:
                s.append(seg(" However IHBR are normal in caliber."))
            else:
                s.append(seg(" However no significant dilatation of IHBR appreciated "
                             "at the time of scan."))
        elif ihbr == "proximal":
            s.append(seg(" "))
            s.append(seg("Intra hepatic biliary radicals are proximally dilated.", True))
        elif ihbr == "dilated":
            s.append(seg(" "))
            s.append(seg("Intra hepatic biliary radicals are dilated.", True))
    return s


# -------- PANCREAS --------

def pancreas_sentence(d):
    s = [seg("PANCREAS", True, True)]
    status = d.get("status", "normal")

    if status == "normal":
        s.append(seg(" is normal in size, outline and echotexture. No focal lesion is "
                     "seen. No evidence of calcification is seen."))
        return s

    if status == "early_evolving":
        ee_size = d.get("ee_size", "normal")
        fat = d.get("ee_fat_stranding", False)
        fluid = d.get("ee_free_fluid", False)
        fat_loc = d.get("ee_fat_location", "none")
        fluid_loc = d.get("ee_fluid_location", "none")

        if ee_size == "mildly_bulky":
            s.append(seg(" is "))
            s.append(seg("mildly bulky in size", True))
            s.append(seg(" with normal echotexture."))
        else:
            s.append(seg(" is normal in size, outline and echotexture. No focal lesion "
                         "is seen. No evidence of calcification is seen."))

        loc_phrase_map = {
            "head_neck": "around the head and neck region",
            "body": "around the body region",
            "neck_body": "around the neck and body region",
            "perisplenic": "in the peri-splenic region",
            "none": "",
        }

        if fat and fluid:
            loc = loc_phrase_map.get(fat_loc, "")
            if fat_loc == "perisplenic":
                s.append(seg(" "))
                s.append(seg("Mild peri-splenic fat stranding and free fluid seen.", True))
            elif loc:
                s.append(seg(" "))
                s.append(seg(f"Mild peri-pancreatic fat stranding and free fluid "
                             f"seen {loc}.", True))
            else:
                s.append(seg(" "))
                s.append(seg("Mild peri-pancreatic fat stranding and free fluid "
                             "seen.", True))
        elif fat:
            loc = loc_phrase_map.get(fat_loc, "")
            if fat_loc == "perisplenic":
                s.append(seg(" "))
                s.append(seg("Mild peri-splenic fat stranding is seen.", True))
            elif loc:
                s.append(seg(" "))
                s.append(seg(f"Mild peri-pancreatic fat stranding is seen {loc}.", True))
            else:
                s.append(seg(" "))
                s.append(seg("Mild peri-pancreatic fat stranding is seen.", True))
        elif fluid:
            loc = loc_phrase_map.get(fluid_loc, "")
            if fluid_loc == "perisplenic":
                s.append(seg(" "))
                s.append(seg("Mild peri-splenic free fluid is seen.", True))
            elif loc:
                s.append(seg(" "))
                s.append(seg(f"Mild peri-pancreatic free fluid is seen {loc}.", True))
            else:
                s.append(seg(" "))
                s.append(seg("Mild peri-pancreatic free fluid is seen.", True))

        return s

    if status == "acute":
        ac_size = d.get("ac_size", "normal")
        ac_echo = d.get("ac_echo", "normal")
        ac_echo_loc = d.get("ac_echo_location", "none")
        ac_margins = d.get("ac_margins", "normal")

        if ac_size == "bulky":
            s.append(seg(" is "))
            s.append(seg("bulky in size", True))
            if ac_echo == "normal" and ac_margins == "normal":
                s.append(seg(" with normal echotexture."))
            elif ac_echo == "hypoechoic":
                if ac_echo_loc == "head_neck":
                    s.append(seg(" with hypoechoic heterogeneous echotexture in the "
                                 "head and neck region", True))
                elif ac_echo_loc == "body":
                    s.append(seg(" with hypoechoic heterogeneous echotexture in the "
                                 "body region", True))
                else:
                    s.append(seg(" with hypoechoic heterogeneous echotexture", True))
                if ac_margins == "irregular":
                    s.append(seg(" and irregular/fuzzy margins", True))
                s.append(seg("."))
            elif ac_margins == "irregular":
                s.append(seg(" with normal echotexture and irregular/fuzzy margins", True))
                s.append(seg("."))
        else:
            if ac_echo == "normal" and ac_margins == "normal":
                s.append(seg(" is normal in size, outline and echotexture. No focal "
                             "lesion is seen."))
            elif ac_echo == "hypoechoic":
                if ac_echo_loc == "head_neck":
                    s.append(seg(" appears "))
                    s.append(seg("hypoechoic and heterogeneous in the head and neck region",
                                 True))
                elif ac_echo_loc == "body":
                    s.append(seg(" appears "))
                    s.append(seg("hypoechoic and heterogeneous in the body region", True))
                else:
                    s.append(seg(" echotexture appears "))
                    s.append(seg("hypoechoic and heterogeneous", True))
                if ac_margins == "irregular":
                    s.append(seg(" with irregular/fuzzy margins", True))
                s.append(seg("."))
            elif ac_margins == "irregular":
                s.append(seg(" shows "))
                s.append(seg("irregular/fuzzy margins", True))
                s.append(seg("."))

        s.append(seg(" "))
        s.append(seg("Mild to moderate peri-pancreatic fat stranding and mild free "
                     "fluid is seen.", True))
        return s

    return s


# -------- SPLEEN --------

def spleen_sentence(d):
    s = [seg("SPLEEN", True, True)]
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}MM) with normal echotexture. "
                     f"Splenic vein is normal."))
    elif desc == "borderline":
        s += [seg(" is "), seg("borderline enlarged in size", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    elif desc == "mild":
        s += [seg(" is "), seg("mildly enlarged in size", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    elif desc == "moderate":
        s += [seg(" is "), seg("moderately enlarged in size", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("enlarged for age in size", True),
              seg(f" ({size}MM) with normal echotexture. Splenic vein is normal.")]
    if d["portal_vein_mm"]:
        s += [seg(" "), seg("Portal vein is normal in course and caliber", True),
              seg(f" ({d['portal_vein_mm']}MM)"), seg(".", True)]
    return s


# -------- KIDNEYS --------

def kidneys_sentence(d):
    s = []
    r, l = d["right"], d["left"]
    both_normal = (r["status"] == "normal" and not r["calculi"] and r["cyst"] == "none"
                   and r["hydronephrosis"] == "none" and r["nephrocalcinosis"] == "none"
                   and l["status"] == "normal" and not l["calculi"] and l["cyst"] == "none"
                   and l["hydronephrosis"] == "none"
                   and l["nephrocalcinosis"] == "none")
    if both_normal:
        s.append(seg("BOTH KIDNEYS", True, True))
        s.append(seg(" are normal in size, outline and "))
        if d["cortical_echogenicity"] == "mildly_raised_bilateral":
            s += [seg("mildly raised bilateral renal cortical echogenicity", True),
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
            b += [seg(" "), seg(f"A calculus measuring {calc.get('size_mm','')}MM is seen "
                                 f"at the {calc.get('location','')} of "
                                 f"{side_label.lower()} kidney", True)]
            if calc.get("hydro") and calc["hydro"] != "none":
                b.append(seg(f" causing {side_label.lower()} sided "
                             f"{calc['hydro']} hydroureteronephrosis", True))
            b.append(seg(".", True))
        if k["cyst"] != "none":
            bosniak = f" (Bosniak cat-{k['bosniak']})" if k["bosniak"] else ""
            b += [seg(" "), seg(f"A {k['cyst']} cyst measuring "
                                 f"{k['cyst_size_mm']}MM is seen at the "
                                 f"{k['cyst_location']} of {side_label.lower()} "
                                 f"kidney{bosniak}", True), seg(".", True)]
        if k["hydronephrosis"] != "none":
            b += [seg(" "), seg(f"{k['hydronephrosis'].title()} hydroureteronephrosis "
                                 f"is present on the {side_label.lower()}", True),
                  seg(".", True)]
        if k["nephrocalcinosis"] != "none":
            b += [seg(" "), seg(f"Multiple foci of calcification seen in the "
                                 f"{side_label.lower()} renal cortex", True),
                  seg(".", True)]
        return b

    s.extend(block("RIGHT KIDNEY", r, "RIGHT"))
    s.append(seg("\n"))
    s.extend(block("LEFT KIDNEY", l, "LEFT"))
    return s


# -------- URINARY BLADDER --------

def urinary_bladder_sentence(d):
    s = [seg("URINARY BLADDER", True, True)]
    if d["status"] == "adequately_distended":
        s.append(seg(" is adequately distended."))
    elif d["status"] == "over":
        s += [seg(" is "), seg("over-distended", True), seg(".")]
    elif d["status"] == "partially":
        s.append(seg(" is partially empty."))
    elif d["status"] == "empty":
        s.append(seg(" is empty."))
    if not d["mass_calculus"]:
        s.append(seg(" No mass or calculus seen."))
    sed = d["sedimentation"]
    if sed == "trace":
        s += [seg(" "), seg("Trace sedimentation seen in the UB lumen.", True)]
    elif sed == "free_floating":
        s += [seg(" "), seg("Free floating sedimentation seen in the UB lumen", True),
              seg(".")]
    elif sed == "significant":
        s += [seg(" "), seg("Significant sedimentation seen in the UB lumen", True),
              seg(".")]
    elif sed == "extensive":
        s += [seg(" "), seg("Extensive sedimentation seen in the UB lumen", True),
              seg(".")]
    return s


# -------- UTERUS --------

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
    if status == "anteverted":
        s.append(seg(" is anteverted"))
        if d["size"]:
            s.append(seg(f" and normal in size({d['size']}MM)"))
        s.append(seg(" with normal shape and echopattern."))
    elif status in ("retroverted", "retroflexed"):
        s.append(seg(f" is {status}", True))
        if d["size"]:
            s.append(seg(f" and normal in size({d['size']}MM)"))
        s.append(seg(" with normal shape and echopattern."))
    elif status == "bulky":
        s += [seg(" is "), seg("bulky in size", True)]
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
        s += [seg(" "), seg(d["fibroid_text"].capitalize(), True), seg(".", True)]
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
                b += [seg(" appears normal in size and echo pattern. "),
                      seg(f"[{size}MM].", True)]
            b += [seg(" "), seg(f"A {cyst_type} cyst measuring {cyst_size}MM is "
                                 f"seen in the {label.split()[0].lower()} ovary", True),
                  seg(".", True)]
        return b

    s.extend(block("RIGHT OVARY", r, d["right_size"], d["right_cyst_type"],
                   d["right_cyst_size"]))
    s.append(seg(" "))
    s.extend(block("LEFT OVARY", l, d["left_size"], d["left_cyst_type"],
                   d["left_cyst_size"]))
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
        s.append(seg(" is normal in size and attenuation. No diffuse or focal lesion "
                     "seen."))
    elif d["status"] == "borderline":
        s += [seg(" is "), seg("borderline enlarged in size", True)]
        if d["size_cc"]:
            s.append(seg(f"({d['size_cc']}CC)"))
        s.append(seg(" and attenuation. No diffuse or focal lesion seen."))
    elif d["status"] == "bulky":
        s += [seg(" is "), seg("bulky in size", True)]
        if d["size_cc"]:
            s.append(seg(f"({d['size_cc']}CC)"))
        s.append(seg(" and attenuation. No diffuse or focal lesion seen."))
    elif d["status"] == "grade1":
        s += [seg(" shows "), seg(f"grade-I prostatomegaly ({d['size_cc']}CC)", True),
              seg(". No diffuse or focal lesion seen.")]
    return s


def bowel_sentence(d, sex, pancreas_status="normal"):
    """If pancreas_status == 'acute', suppress free fluid phrase and append Mild ascites."""
    acute = (pancreas_status == "acute")
    s = []
    if sex == "F":
        if not d["wall_thickening"]:
            s.append(seg("No obvious bowel wall thickening seen.", True))
        else:
            s.append(seg("Bowel wall thickening seen.", True))
        if not acute:
            s.append(seg(" "))
            if d["free_fluid"] == "none":
                s.append(seg("No free fluid seen in the peritoneal cavity."))
            else:
                s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free fluid seen.",
                             True))
    else:
        if not acute:
            if d["free_fluid"] == "none":
                s.append(seg("No free fluid is seen in the peritoneal cavity."))
            else:
                s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free fluid seen.",
                             True))
            s.append(seg(" "))
        if not d["wall_thickening"]:
            s.append(seg("No obvious bowel wall thickening seen.", True))
        else:
            s.append(seg("Bowel wall thickening seen.", True))
    if d["mesenteric_ln"] == "present":
        cat = {"sad_lt_7": "Small(SAD<7MM)", "sad_gt_7": "Enlarged(SAD>7MM)",
               "sad_gt_10": "Enlarged(SAD>10MM)"}.get(d["ln_size_category"], "Mesenteric")
        loc = d["ln_location"].replace("_", " ").title()
        largest = (f" with largest of these measuring {d['ln_largest']}"
                   if d["ln_largest"] else "")
        s += [seg(" "), seg(f"{cat} mesenteric lymph nodes are seen in the {loc} region"
                             f"{largest}", True), seg(".", True)]
    pe = d["pleural_effusion"]
    if pe != "none":
        text = {"trace_right": "Trace right pleural effusion is seen",
                "trace_left": "Trace left pleural effusion is seen",
                "mild_right": "Mild right pleural effusion is seen",
                "mild_bilateral": "Trace left & mild right pleural effusion seen"}.get(pe, "")
        s += [seg(" "), seg(text, True), seg(".", True)]
    if acute:
        s.append(seg("\nMild ascites is seen.", True))
    return s


def appendix_sentence(d):
    if d["status"] == "not_assessed":
        return []
    s = []
    if d["status"] == "not_visualized":
        s.append(seg("Appendix is not visualized."))
    elif d["status"] == "normal":
        s.append(seg("Appendix is visualized and appears normal", True))
        if d["diameter_mm"]:
            s.append(seg(f" ({d['diameter_mm']}MM in diameter)", True))
        s.append(seg(".", True))
    elif d["status"] == "dilated":
        s += [seg("Appendix is dilated upto", True), seg(f" {d['diameter_mm']}MM", True),
              seg(" - ?Evolving appendicitis vs physiological.", True)]
    return s


# ============================================================
# IMPRESSION GENERATOR
# ============================================================

def generate_impression(d, sex, age):
    lines = []
    age_years = parse_age(age)
    is_pediatric = age_years is not None and age_years < 18

    # ---- PANCREAS ----
    p = d["pancreas"]
    p_status = p.get("status", "normal")

    if p_status == "early_evolving":
        ee_size = p.get("ee_size", "normal")
        fat = p.get("ee_fat_stranding", False)
        fluid = p.get("ee_free_fluid", False)
        fat_loc = p.get("ee_fat_location", "none")
        fluid_loc = p.get("ee_fluid_location", "none")

        loc_full = {
            "head_neck": "HEAD AND NECK REGION",
            "body": "BODY REGION",
            "neck_body": "NECK AND BODY REGION",
            "perisplenic": "PERI-SPLENIC REGION",
            "none": "",
        }

        def tail():
            return (" ?EARLY / EVOLVING PANCREATITIS. "
                    "Adv- S.Amylase/Lipase Correlation.")

        if fat and fluid:
            if fat_loc == "perisplenic":
                lines.append("MILD PERI-SPLENIC FAT STRANDING & FREE FLUID SEEN."
                             + tail())
            elif fat_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FAT STRANDING & FREE FLUID SEEN "
                             f"AROUND THE {loc_full[fat_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FAT STRANDING & FREE FLUID SEEN."
                             + tail())
        elif fat:
            if fat_loc == "perisplenic":
                lines.append("MILD PERI-SPLENIC FAT STRANDING SEEN." + tail())
            elif fat_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FAT STRANDING SEEN AROUND THE "
                             f"{loc_full[fat_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FAT STRANDING SEEN." + tail())
        elif fluid:
            if fluid_loc == "perisplenic":
                lines.append("MILD FREE FLUID SEEN IN THE PERI-SPLENIC REGION."
                             + tail())
            elif fluid_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FREE FLUID SEEN AROUND THE "
                             f"{loc_full[fluid_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FREE FLUID SEEN." + tail())

        if ee_size == "mildly_bulky":
            if lines and (fat or fluid):
                last = lines.pop()
                new_last = "MILDLY BULKY PANCREAS WITH " + last
                lines.append(new_last)
            else:
                lines.append("MILDLY BULKY PANCREAS, HOWEVER NO PERI-PANCREATIC "
                             "FAT STRANDING OR FREE FLUID SEEN." + tail())

    elif p_status == "acute":
        ac_size = p.get("ac_size", "normal")
        ac_echo = p.get("ac_echo", "normal")
        ac_margins = p.get("ac_margins", "normal")

        if ac_size == "bulky":
            lines.append(
                "MILD ASCITES WITH FEATURES SUGGESTIVE OF ACUTE EDEMATOUS "
                "PANCREATITIS. Adv- S.Amylase/Lipase Correlation."
            )
        elif ac_echo == "hypoechoic" or ac_margins == "irregular":
            lines.append(
                "MILD ASCITES WITH FEATURES SUGGESTIVE OF ACUTE NECROTIZING "
                "PANCREATITIS. Adv- S.Amylase/Lipase Correlation."
            )
        else:
            lines.append(
                "MILD ASCITES WITH MILD TO MODERATE PERI-PANCREATIC FAT STRANDING "
                "AND MILD PERI-PANCREATIC FREE FLUID - ?ACUTE PANCREATITIS. "
                "Adv- S.Amylase/Lipase Correlation."
            )

    # ---- CBD + GB ----
    cbd = d["cbd"]
    cbd_status = cbd.get("status", "normal")
    cbd_size = cbd.get("size_mm", "")
    cbd_has_calc = cbd.get("calculi", False) and cbd.get("calculi_size_mm")
    cbd_ihbr = cbd.get("ihbr", "normal")
    cbd_dilated = cbd_status in ("proximal", "dilated")

    def ihbr_clause():
        if cbd_ihbr == "normal":
            return ". HOWEVER IHBR ARE NORMAL IN CALIBER."
        elif cbd_ihbr == "proximal":
            return " WITH PROXIMAL DILATATION OF IHBR."
        elif cbd_ihbr == "dilated":
            return " WITH DILATATION OF IHBR."
        return ""

    gb = d["gall_bladder"]
    gb_has_calculi = gb.get("calculi") == "present"
    gb_wall_thickened = gb.get("wall_thickened") and gb.get("wall_mm")
    gb_peri_fluid = gb.get("pericholecystic_fluid", False)
    gb_over = gb.get("status") == "over"
    gb_contracted = gb.get("status") == "contracted"
    gb_sludge = gb.get("sludge", "none") != "none"
    gb_sludge_ball = gb.get("sludge_ball", False)
    gb_comet_tail = gb.get("comet_tail", False)

    wall_descriptor = ""
    if gb_wall_thickened:
        try:
            w = float(gb["wall_mm"])
            wall_descriptor = "MILD" if w <= 8 else "SIGNIFICANT"
        except (ValueError, TypeError):
            wall_descriptor = ""

    if cbd_has_calc and gb_has_calculi and cbd_dilated:
        lines.append(f"CHOLELITHIASIS WITH CHOLEDOCHOLITHIASIS AND DILATED CBD UPTO "
                     f"{cbd_size}MM" + ihbr_clause() +
                     " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and gb_has_calculi and not cbd_dilated:
        lines.append("CHOLELITHIASIS WITH CHOLEDOCHOLITHIASIS AND NORMAL CALIBER CBD"
                     + ihbr_clause() + " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and cbd_dilated:
        lines.append(f"CHOLEDOCHOLITHIASIS WITH DILATED CBD UPTO {cbd_size}MM"
                     + ihbr_clause() + " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and not cbd_dilated:
        lines.append("CHOLEDOCHOLITHIASIS WITH NORMAL CALIBER CBD"
                     + ihbr_clause() + " Adv- MRCP/CECT Abdomen Correlation.")

    gb_calc_already_reported = cbd_has_calc and gb_has_calculi and cbd_dilated

    if not gb_calc_already_reported:
        if gb_over and gb_has_calculi and (gb_wall_thickened or gb_peri_fluid):
            lines.append("OVERDISTENDED GALL BLADDER WITH CHOLELITHIASIS & FEATURES "
                         "SUGGESTIVE OF ACUTE CHOLECYSTITIS.")
        elif gb_over and gb_has_calculi:
            lines.append("OVERDISTENDED GALL BLADDER WITH CHOLELITHIASIS. HOWEVER NO "
                         "PERICHOLECYSTIC FLUID OR GB WALL THICKENING APPRECIATED.")
        elif gb_contracted and gb_has_calculi:
            lines.append("CHOLELITHIASIS WITH ?CHRONIC CHOLECYSTITIS.")
        elif gb_has_calculi and (gb_wall_thickened or gb_peri_fluid):
            lines.append("CHOLELITHIASIS WITH ?ACUTE CHOLECYSTITIS.")
        elif gb_has_calculi:
            lines.append("CHOLELITHIASIS WITH NO APPRECIABLE PERICHOLECYSTIC FLUID OR "
                         "GB WALL THICKENING.")
        elif gb_wall_thickened and gb_peri_fluid:
            lines.append(f"GB WALL THICKENING {wall_descriptor}, SEEN UP TO "
                         f"{gb['wall_mm']}MM, WITH THIN RIM OF PERICHOLECYSTIC FLUID - "
                         f"?ACALCULUS CHOLECYSTITIS. Adv- LFT, Lab & Clinical Correlation.")
        elif gb_wall_thickened and not gb_has_calculi:
            lines.append("ISOLATED GB WALL THICKENING WITHOUT ANY CALCULUS - "
                         "?ACALCULUS CHOLECYSTITIS. Adv- LFT, Lab and Clinical Correlation.")
        elif gb_peri_fluid and not gb_has_calculi:
            lines.append("THIN RIM OF PERICHOLECYSTIC FLUID SEEN - ?SIGNIFICANCE. "
                         "Adv- LFT, Lab and Clinical Correlation.")
        elif gb_sludge and not gb_has_calculi:
            lines.append(f"{gb['sludge'].upper()} SLUDGE SEEN IN THE GALLBLADDER LUMEN. "
                         f"Adv- Review scan after a month.")
        elif gb_sludge_ball and not gb_has_calculi:
            count = gb.get("sludge_ball_count", "single")
            wall = gb.get("sludge_ball_wall", "anterior").upper()
            if count == "single":
                lines.append(f"A SLUDGE BALL/POLYP SEEN IMPACTED AT THE {wall} GB WALL. "
                             f"Adv- Review scan after a month.")
            else:
                word = "FEW" if count == "few" else "MULTIPLE"
                lines.append(f"{word} SLUDGE BALLS/POLYPS SEEN IMPACTED AT THE {wall} GB "
                             f"WALL. Adv- Review scan after a month.")

    if gb_comet_tail:
        if gb_has_calculi:
            lines.append("CHOLELITHIASIS WITH GALL BLADDER "
                         "ADENOMYOMATOSIS/CHOLESTEROLOSIS.")
        elif gb_sludge:
            lines.append("GALL BLADDER SLUDGE WITH GALL BLADDER "
                         "ADENOMYOMATOSIS/CHOLESTEROLOSIS.")
        else:
            lines.append("GALL BLADDER ADENOMYOMATOSIS/CHOLESTEROLOSIS.")

    # ---- LIVER (combined) ----
    liver = d["liver"]
    desc = liver["size_descriptor"]
    hepatomegaly = None
    if desc == "enlarged_for_age":
        hepatomegaly = "HEPATOMEGALY FOR AGE"
    elif desc != "normal":
        dm = {"borderline": "BORDERLINE HEPATOMEGALY",
              "mild": "MILD HEPATOMEGALY",
              "moderate": "MODERATE HEPATOMEGALY",
              "gross": "GROSS HEPATOMEGALY"}
        hepatomegaly = dm.get(desc, "")

    lf = []

    if liver["echotexture"] == "increased":
        grade = liver.get("steatosis_grade") or ""
        if grade == "Severe+++":
            lf.append("SIGNIFICANT FATTY INFILTRATION - "
                      "?NON-ALCOHOLIC STEATO\u2011HEPATITIS")
        elif grade:
            lf.append(f"HEPATIC STEATOSIS({grade.upper()})")
        else:
            lf.append("HEPATIC STEATOSIS")

    fl = liver["focal_lesion"]
    if fl == "calcified":
        lf.append("A CALCIFIED FOCUS IN THE RIGHT HEPATIC LOBE")
    elif fl == "cyst":
        count = liver.get("cyst_count", "single")
        if count == "single":
            lobe = liver.get("cyst_single_lobe", "right").upper()
            lf.append(f"A SIMPLE HEPATIC CYST IN {lobe} LOBE")
        else:
            lobe = liver.get("cyst_few_lobe", "right").upper()
            lf.append(f"FEW SIMPLE HEPATIC CYSTS IN {lobe} LOBE")
    elif fl == "hemangioma":
        count = liver.get("hemangioma_count", "single")
        if count == "single":
            lobe = liver.get("hemangioma_single_lobe", "right").upper()
            lf.append(f"A HEPATIC HEMANGIOMA IN {lobe} LOBE")
        else:
            lobe = liver.get("hemangioma_few_lobe", "right").upper()
            lf.append(f"FEW HEPATIC HEMANGIOMAS IN {lobe} LOBE")
    elif fl == "abscess":
        count = liver.get("abscess_count", "single")
        lesions = liver.get("abscess_lesions", [])
        if lesions:
            if count == "single":
                l0 = lesions[0]
                lf.append(f"AN IRREGULAR MARGINATED ILL-DEFINED AVASCULAR SOL IN THE "
                          f"LIVER MEASURING VOL= {l0['vol']}CC IN SEGMENT "
                          f"{l0['segment']} - LIKELY LIVER ABSCESS")
            else:
                parts = [f"VOL= {l['vol']}CC IN SEGMENT {l['segment']}"
                         for l in lesions]
                joined = " & ".join(parts)
                keyword = "FEW" if count == "few" else "MULTIPLE"
                lf.append(f"{keyword} IRREGULAR MARGINATED ILL-DEFINED AVASCULAR SOLS "
                          f"IN THE LIVER, LARGEST OF THESE MEASURING {joined} - "
                          f"LIKELY LIVER ABSCESSES")

    if hepatomegaly and lf:
        lines.append(hepatomegaly + " WITH " + " AND ".join(lf)
                     + ". Adv- LFT Correlation.")
    elif hepatomegaly:
        lines.append(hepatomegaly + ". Adv- LFT Correlation.")
    elif lf:
        lines.append(" AND ".join(lf) + ". Adv- LFT Correlation.")

    # ---- SPLEEN ----
    spleen_desc = d["spleen"]["size_descriptor"]
    if spleen_desc == "enlarged_for_age":
        lines.append("SPLENOMEGALY FOR AGE.")
    elif spleen_desc != "normal":
        lines.append(f"{spleen_desc.upper()} SPLENOMEGALY "
                     f"({d['spleen']['size_mm']}MM).")

    if d["urinary_bladder"]["sedimentation"] in ("free_floating", "significant",
                                                  "extensive"):
        lines.append("SEDIMENTATION SEEN IN THE UB LUMEN. Adv- Urine R/M Correlation.")

    if sex == "M" and not is_pediatric:
        pr = d["prostate"]
        if pr["status"] == "bulky":
            lines.append(f"GRADE-I PROSTATOMEGALY ({pr['size_cc']}CC).")
        elif pr["status"] == "borderline":
            lines.append(f"BORDERLINE PROSTATOMEGALY ({pr['size_cc']}CC).")

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
    if p_status != "acute" and b["free_fluid"] not in ("none",):
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
# DOCX BUILDER
# ============================================================

def _set_table_borders(table, size=6):
    tbl = table._tbl
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(size))
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)


def _set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("left", left),
                      ("bottom", bottom), ("right", right)):
        m = OxmlElement(f"w:{edge}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tcPr.append(mar)


def _add_run(paragraph, text, bold=False, underline=False, font=FONT_BODY,
             size=FONT_SIZE_BODY, italic=False, color=None):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.underline = underline
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rPr.append(rFonts)
    return run


def _split_adv(text):
    idx = text.find("Adv-")
    if idx == -1:
        return text, ""
    return text[:idx].rstrip(), text[idx:]


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
    _set_table_borders(table)
    p = data["patient"]
    cells = [
        (0, 0, f"NAME :- {p['name']}"),
        (0, 1, f"DATE :- {p['date']}"),
        (1, 0, f"AGE/SEX :- {p['age']}Y/{p['sex']}"),
        (1, 1, f"REF. BY: {p['referred_by']}"),
    ]
    for r, c, text in cells:
        cell = table.cell(r, c)
        cell.text = ""
        para = cell.paragraphs[0]
        _add_run(para, text, bold=True)

    doc.add_paragraph()
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(title_para, "ULTRASOUND WHOLE ABDOMEN", bold=True, underline=True,
             size=FONT_SIZE_TITLE, color=TITLE_COLOR)
    doc.add_paragraph()

    sex, age = p["sex"], p["age"]
    age_years = parse_age(age)
    is_ped = age_years is not None and age_years < 18
    p_status = data["pancreas"].get("status", "normal")

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
    sections.append(bowel_sentence(data["bowel"], sex, pancreas_status=p_status))
    if data["appendix"]["status"] != "not_assessed":
        sections.append(appendix_sentence(data["appendix"]))

    for segs in sections:
        if not segs:
            continue
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for text, bold, underline in segs:
            italic = bold and not underline
            if "\n" in text:
                parts = text.split("\n")
                for i, part in enumerate(parts):
                    if i > 0:
                        para.add_run().add_break()
                    _add_run(para, part, bold=bold, underline=underline,
                             italic=italic)
            else:
                _add_run(para, text, bold=bold, underline=underline,
                         italic=italic)
        para.paragraph_format.space_after = Pt(6)

    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True,
             size=FONT_SIZE_BODY, color=TITLE_COLOR)

    for line in data["impression"]["lines"]:
        para = doc.add_paragraph(style="List Bullet")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        main, adv = _split_adv(line)
        _add_run(para, main, bold=True)
        if adv:
            _add_run(para, " " + adv, bold=True, italic=True)

    doc.add_paragraph()
    disc_table = doc.add_table(rows=1, cols=1)
    _set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    _set_cell_margins(cell, top=20, bottom=20, left=80, right=80)
    cell.text = ""
    para = cell.paragraphs[0]
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    _add_run(para, DISCLAIMER_TEXT, bold=True, size=FONT_SIZE_DISCLAIMER)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="PG Imaging & Diagnostics", layout="wide")
init_db()

st.markdown(
    """
    <style>
    :root { color-scheme: light !important; }

    #MainMenu {visibility: hidden;}
    header[data-testid="stHeader"] {visibility: hidden; height: 0;}
    footer {visibility: hidden;}

    .stApp, [data-testid="stAppViewContainer"], section.main {
        background-color: #f4f6f9 !important;
    }

    .stApp, .stApp p, .stApp label, .stApp span, .stApp div,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp li {
        color: #111111;
    }

    .pg-title {
        text-align: center; margin: 0 0 0.15rem 0; padding: 0;
        color: #1F4E79 !important; font-weight: 700; font-size: 2rem;
        letter-spacing: 0.5px;
    }
    .pg-tagline {
        text-align: center; margin: 0 0 1.0rem 0; padding: 0;
        color: #555 !important; font-style: italic; font-size: 1.05rem;
    }

    @media (max-width: 900px) {
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important; gap: 0.4rem !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            min-width: 0 !important;
        }
    }

    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        display: flex !important; flex-direction: column !important;
        justify-content: flex-end !important;
    }

    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #cfd6dd !important;
        border-radius: 6px !important;
        margin-bottom: 4px !important;
    }
    div[data-testid="stExpander"] details > summary {
        padding: 8px 12px !important; min-height: 42px !important;
        box-sizing: border-box !important; display: flex !important;
        align-items: center !important;
        background-color: #ffffff !important;
    }
    div[data-testid="stExpander"] details[open] > summary {
        background-color: #eef2f7 !important;
        border-bottom: 1px solid #cfd6dd !important;
    }
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] summary *,
    div[data-testid="stExpander"] details[open] > summary,
    div[data-testid="stExpander"] details[open] > summary * {
        color: #111111 !important; font-weight: 600 !important;
        font-size: 14px !important;
        white-space: nowrap !important; overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    div[data-testid="stExpander"] details[open] > summary svg,
    div[data-testid="stExpander"] details[open] > summary svg path {
        fill: #111111 !important;
        stroke: #111111 !important;
    }
    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"],
    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"] * {
        color: #111111 !important;
    }

    div[data-testid="stNumberInput"], div[data-testid="stNumberInput"] > div {
        margin: 0 !important;
    }
    div[data-testid="stNumberInput"] input {
        background-color: #ffffff !important; color: #111111 !important;
        border: 1px solid #cfd6dd !important; border-radius: 6px !important;
        min-height: 42px !important; box-sizing: border-box !important;
        font-weight: 600 !important; text-align: center !important;
        padding: 0 6px !important;
    }
    div[data-testid="stNumberInput"] button { display: none !important; }

    .stTextInput input, .stTextArea textarea,
    [data-baseweb="input"] input, [data-baseweb="base-input"] input {
        background-color: #ffffff !important; color: #111111 !important;
        border: 1px solid #cfd6dd !important;
        min-height: 42px !important; box-sizing: border-box !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: #888 !important;
    }

    .stRadio label, .stRadio span, .stCheckbox label, .stCheckbox span {
        color: #111111 !important;
    }
    .stButton button, .stDownloadButton button {
        background-color: #ffffff !important; color: #111111 !important;
        border: 1px solid #cfd6dd !important;
    }

    .preview-box,
    .stApp .preview-box,
    .stApp .preview-box * {
        color: #ffffff !important;
    }
    .preview-box {
        background-color: #000000 !important;
        font-family: Consolas, Menlo, 'Courier New', monospace !important;
        font-size: 12.5px !important;
        line-height: 1.5 !important;
        padding: 14px 16px !important;
        border-radius: 8px !important;
        border: 1px solid #333 !important;
        white-space: pre-wrap !important;
        word-wrap: break-word !important;
        overflow-wrap: anywhere !important;
        word-break: break-word !important;
        margin-bottom: 8px !important;
        display: block !important;
    }

    .st-key-impression_box textarea,
    div[class*="st-key-impression_box"] textarea {
        background-color: #000000 !important;
        color: #ffffff !important;
        font-family: Consolas, Menlo, 'Courier New', monospace !important;
        font-size: 12.5px !important;
        line-height: 1.5 !important;
        border: 1px solid #333 !important;
        white-space: pre-wrap !important;
        word-wrap: break-word !important;
        overflow-wrap: anywhere !important;
        word-break: break-word !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="pg-title">PG Imaging &amp; Diagnostics</div>
    <div class="pg-tagline">Precision Imaging. Trusted Diagnostics. Agra.</div>
    """,
    unsafe_allow_html=True,
)

with st.container():
    c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3],
                                     vertical_alignment="bottom")
    with c1:
        p_name = st.text_input("Name", key="p_name_input")
    with c2:
        p_age = st.text_input("Age (years)", key="p_age_input",
                              help="Enter age in years, e.g. 25 or 0.5")
    with c3:
        p_sex = st.radio("Sex", ["F", "M"], horizontal=True, key="p_sex_input")
    with c4:
        p_date = st.text_input("Date",
                                value=datetime.now().strftime("%d-%b-%y"),
                                key="p_date_input")
    with c5:
        p_ref = st.text_input("Referred by", key="p_ref_input")

st.markdown("---")

col_find, col_prev = st.columns([1, 1])

with col_find:
    st.subheader("Findings")

    # ------- LIVER -------
    col_liver_main, col_liver_sz = st.columns([5, 1],
                                              vertical_alignment="bottom")
    with col_liver_main:
        liver_expander = st.expander("LIVER findings (click to open)",
                                     expanded=False)
    with col_liver_sz:
        liver_size_num = st.number_input(
            "Liver size mm", min_value=0, max_value=500, value=None, step=1,
            key="liver_size_num", label_visibility="collapsed",
            placeholder="LIVER mm")
    liver_size = str(int(liver_size_num)) if (liver_size_num and liver_size_num > 0) else ""
    liver_status = auto_classify_liver(liver_size_num or 0, p_age, p_sex)
    if liver_size_num and liver_size_num > 0:
        if liver_status == "normal":
            st.success(f"✓ Liver: **{LIVER_STATUS_LABELS[liver_status]}**")
        elif liver_status == "enlarged_for_age":
            st.info(f"→ Liver: **{LIVER_STATUS_LABELS[liver_status]}**")
        else:
            st.warning(f"→ Liver: **{LIVER_STATUS_LABELS[liver_status]}**")

    with liver_expander:
        liver_outline = "normal"
        liver_echo = "normal"
        liver_steatosis = None
        liver_focal = "none"
        liver_focal_text = ""
        liver_ihbr = "normal"
        liver_portal = "normal"
        liver_portal_mm = ""
        cyst_count = "single"
        cyst_single_lobe = "right"
        cyst_single_size_mm = ""
        cyst_few_largest_mm = ""
        cyst_few_lobe = "right"
        hemangioma_count = "single"
        hemangioma_single_lobe = "right"
        hemangioma_single_size_mm = ""
        hemangioma_few_largest_mm = ""
        hemangioma_few_lobe = "right"
        abscess_count = "single"
        abscess_lesions = []

        liver_outline = st.radio(
            "Outline", ["normal", "crenated"], horizontal=True,
            format_func=lambda x: {"normal": "Normal",
                                   "crenated": "Crenated / nodular"}[x],
            key="liver_outline")
        liver_echo = st.radio(
            "Echotexture", ["normal", "increased", "coarse", "low"],
            horizontal=True,
            format_func=lambda x: {"normal": "Normal",
                                   "increased": "Increased (steatosis)",
                                   "coarse": "Coarse", "low": "Low"}[x],
            key="liver_echo")
        if liver_echo == "increased":
            grade_opts = ["Mild+", "Mild to Moderate++", "Moderate++",
                          "Moderate to Severe+++", "Severe+++"]
            liver_steatosis = st.radio("Steatosis grade (impression only)",
                                       grade_opts, horizontal=True,
                                       key="liver_steatosis")
        st.markdown("**Focal lesion**")
        liver_focal = st.radio(
            "Focal lesion type",
            ["none", "calcified", "cyst", "hemangioma", "abscess", "other"],
            horizontal=True, label_visibility="collapsed",
            format_func=lambda x: {"none": "None", "calcified": "Calcified",
                                   "cyst": "Simple cyst(s)",
                                   "hemangioma": "Hemangioma(s)",
                                   "abscess": "Abscess(es)", "other": "Other"}[x],
            key="liver_focal")
        if liver_focal == "cyst":
            cyst_count = st.radio("Number", ["single", "few"], horizontal=True,
                                  key="cyst_count")
            if cyst_count == "single":
                c_a, c_b = st.columns(2)
                with c_a:
                    cyst_single_lobe = st.radio("Lobe", ["right", "left"],
                                                horizontal=True,
                                                key="cyst_single_lobe")
                with c_b:
                    cyst_single_size_mm = st.text_input("Size (mm)",
                                                        key="cyst_single_size")
            else:
                c_a, c_b = st.columns(2)
                with c_a:
                    cyst_few_lobe = st.radio("Lobe", ["right", "left"],
                                             horizontal=True, key="cyst_few_lobe")
                with c_b:
                    cyst_few_largest_mm = st.text_input("Largest (mm)",
                                                        key="cyst_few_largest")
        elif liver_focal == "hemangioma":
            hemangioma_count = st.radio("Number", ["single", "few"],
                                        horizontal=True, key="hemangioma_count")
            if hemangioma_count == "single":
                c_a, c_b = st.columns(2)
                with c_a:
                    hemangioma_single_lobe = st.radio(
                        "Lobe", ["right", "left"], horizontal=True,
                        key="hemangioma_single_lobe")
                with c_b:
                    hemangioma_single_size_mm = st.text_input(
                        "Size (mm)", key="hemangioma_single_size")
            else:
                c_a, c_b = st.columns(2)
                with c_a:
                    hemangioma_few_lobe = st.radio(
                        "Lobe", ["right", "left"], horizontal=True,
                        key="hemangioma_few_lobe")
                with c_b:
                    hemangioma_few_largest_mm = st.text_input(
                        "Largest (mm)", key="hemangioma_few_largest")
        elif liver_focal == "abscess":
            abscess_count = st.radio("Number", ["single", "few", "multiple"],
                                     horizontal=True, key="abscess_count")
            n_abs = 1 if abscess_count == "single" else int(st.number_input(
                "How many lesions?", min_value=1, max_value=5, value=2,
                key="abscess_n"))
            for i in range(n_abs):
                st.markdown(f"*Lesion {i+1}*")
                c_a, c_b, c_c = st.columns([1, 2, 1])
                with c_a:
                    seg_letter = st.selectbox("Segment",
                                              ["I", "II", "III", "IV",
                                               "V", "VI", "VII", "VIII"],
                                              key=f"abs_seg_{i}")
                with c_b:
                    dim = st.text_input("Dimensions (XxYxZ mm)",
                                        key=f"abs_dim_{i}")
                with c_c:
                    vol = st.text_input("Volume (cc)", key=f"abs_vol_{i}")
                abscess_lesions.append({"segment": seg_letter,
                                        "dim": dim, "vol": vol})
        elif liver_focal == "other":
            liver_focal_text = st.text_input("Description",
                                             key="liver_focal_text")

        liver_ihbr = st.radio("IHBR", ["normal", "dilated"], horizontal=True,
                              key="liver_ihbr")
        c_a, c_b = st.columns([2, 1])
        with c_a:
            liver_portal = st.radio("Portal vein", ["normal", "dilated"],
                                    horizontal=True, key="liver_portal")
        with c_b:
            if liver_portal == "dilated":
                liver_portal_mm = st.text_input("Portal vein size (mm)",
                                                key="liver_portal_mm")

    # ------- GALL BLADDER -------
    with st.expander("GALL BLADDER (click to open findings)", expanded=False):
        gb_status = st.radio(
            "Distension",
            ["adequately_distended", "over", "partially", "contracted",
             "empty", "operated"],
            horizontal=True,
            format_func=lambda x: {"adequately_distended": "Adequate",
                                   "over": "Over-distended",
                                   "partially": "Partially contracted",
                                   "contracted": "Contracted",
                                   "empty": "Empty",
                                   "operated": "Operated"}[x],
            key="gb_status")

        gb_wall_thickened = False
        gb_wall_mm = ""
        gb_calculi = "none"
        gb_calculi_count = "single"
        gb_calculi_size_cat = "small"
        gb_calculi_size_mm = ""
        gb_calculi_neck = False
        gb_calculi_neck_size_mm = ""
        gb_sludge = "none"
        gb_sludge_ball = False
        gb_sludge_ball_count = "single"
        gb_sludge_ball_size_mm = ""
        gb_sludge_ball_wall = "anterior"
        gb_comet_tail = False
        gb_comet_tail_count = "single"
        gb_comet_tail_wall = "anterior"
        gb_peri_fluid = False

        if gb_status not in ("contracted", "operated"):
            gb_wall_thickened = st.checkbox("Wall thickening present",
                                            key="gb_wall_check")
            if gb_wall_thickened:
                gb_wall_mm = st.text_input("Wall thickness (mm)",
                                           key="gb_wall_mm")

        gb_calculi = st.radio("Calculi", ["none", "present"], horizontal=True,
                              key="gb_calculi")
        if gb_calculi == "present":
            c_a, c_b = st.columns(2)
            with c_a:
                gb_calculi_count = st.radio(
                    "Count", ["single", "few", "multiple", "innumerable"],
                    horizontal=True, format_func=lambda x: x.title(),
                    key="gb_calc_count")
            with c_b:
                if gb_calculi_count != "innumerable":
                    gb_calculi_size_cat = st.radio(
                        "Size", ["small", "large"], horizontal=True,
                        format_func=lambda x: x.title(),
                        key="gb_calc_size_cat")
            if gb_calculi_count != "innumerable":
                gb_calculi_size_mm = st.text_input("Largest size (mm)",
                                                   key="gb_calc_size")
                gb_calculi_neck = st.checkbox("Calculus at GB neck",
                                              key="gb_calc_neck")
                if gb_calculi_neck:
                    gb_calculi_neck_size_mm = st.text_input(
                        "Neck calculus size (mm)", key="gb_neck_size")

        gb_sludge = st.radio(
            "Sludge", ["none", "trace", "significant", "echogenic", "organized"],
            horizontal=True, format_func=lambda x: x.title(),
            key="gb_sludge")

        gb_sludge_ball = st.checkbox("Sludge ball / polyp present",
                                     key="gb_sludge_ball")
        if gb_sludge_ball:
            c_a, c_b = st.columns(2)
            with c_a:
                gb_sludge_ball_count = st.radio(
                    "Count", ["single", "few", "multiple"],
                    horizontal=True, key="gb_sb_count")
            with c_b:
                gb_sludge_ball_wall = st.radio(
                    "Wall", ["anterior", "posterior"], horizontal=True,
                    format_func=lambda x: x.title(), key="gb_sb_wall")
            gb_sludge_ball_size_mm = st.text_input("Size (mm)",
                                                   key="gb_sb_size")

        gb_comet_tail = st.checkbox(
            "Comet tail artifacts (adenomyomatosis / cholesterolosis)",
            key="gb_comet_tail")
        if gb_comet_tail:
            c_a, c_b = st.columns(2)
            with c_a:
                gb_comet_tail_count = st.radio(
                    "Count", ["single", "few", "multiple"], horizontal=True,
                    key="gb_ct_count")
            with c_b:
                gb_comet_tail_wall = st.radio(
                    "Wall", ["anterior", "posterior"], horizontal=True,
                    format_func=lambda x: x.title(), key="gb_ct_wall")

        gb_peri_fluid = st.checkbox(
            "Thin rim of pericholecystic fluid present",
            key="gb_peri_fluid")

    # ------- CBD -------
    col_cbd_main, col_cbd_sz = st.columns([5, 1],
                                          vertical_alignment="bottom")
    with col_cbd_sz:
        cbd_mm_num = st.number_input(
            "CBD caliber mm", min_value=0.0, max_value=50.0, value=None,
            step=0.1, format="%.2f",
            key="cbd_mm_num", label_visibility="collapsed",
            placeholder="CBD mm")
    cbd_mm = (f"{cbd_mm_num:g}" if (cbd_mm_num and cbd_mm_num > 0) else "")

    with col_cbd_main:
        with st.expander("COMMON BILE DUCT findings (click to open)",
                         expanded=False):
            cbd_status = st.radio(
                "Status", ["normal", "proximal", "dilated"], horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "proximal": "Proximally dilated",
                                       "dilated": "Dilated throughout"}[x],
                key="cbd_status")

            cbd_calc = False
            cbd_calc_count = "single"
            cbd_calc_size = ""
            cbd_calc_location = "distal"
            cbd_ihbr = "normal"

            cbd_calc = st.checkbox("Calculus in CBD", key="cbd_calc")
            if cbd_calc:
                c_a, c_b = st.columns(2)
                with c_a:
                    cbd_calc_count = st.radio(
                        "Count", ["single", "few"], horizontal=True,
                        format_func=lambda x: x.title(), key="cbd_calc_count")
                with c_b:
                    cbd_calc_size = st.text_input("Size (mm, largest if few)",
                                                  key="cbd_calc_size")
                cbd_calc_location = st.radio(
                    "Location", ["proximal", "mid", "distal", "mid_distal"],
                    horizontal=True,
                    format_func=lambda x: {"proximal": "Proximal", "mid": "Mid",
                                           "distal": "Distal",
                                           "mid_distal": "Mid/Distal"}[x],
                    key="cbd_calc_location")

            if cbd_status in ("proximal", "dilated") or cbd_calc:
                cbd_ihbr = st.radio(
                    "IHBR", ["normal", "proximal", "dilated"], horizontal=True,
                    format_func=lambda x: {"normal": "Normal",
                                           "proximal": "Proximally dilated",
                                           "dilated": "Dilated"}[x],
                    key="cbd_ihbr")

    # ------- PANCREAS -------
    with st.expander("PANCREAS (click to open findings)", expanded=False):
        pn_status = st.radio(
            "Status",
            ["normal", "early_evolving", "acute"],
            horizontal=True,
            format_func=lambda x: {"normal": "Normal",
                                   "early_evolving": "Early/evolving pancreatitis",
                                   "acute": "Acute pancreatitis"}[x],
            key="pn_status")

        pn_ee_size = "normal"
        pn_ee_fat = False
        pn_ee_fat_loc = "none"
        pn_ee_fluid = False
        pn_ee_fluid_loc = "none"
        pn_ac_size = "normal"
        pn_ac_echo = "normal"
        pn_ac_echo_loc = "none"
        pn_ac_margins = "normal"

        if pn_status == "early_evolving":
            pn_ee_size = st.radio(
                "Pancreas size",
                ["normal", "mildly_bulky"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal size",
                                       "mildly_bulky": "Mildly bulky"}[x],
                key="pn_ee_size")

            pn_ee_fat = st.checkbox("Mild peri-pancreatic fat stranding",
                                    key="pn_ee_fat")
            if pn_ee_fat:
                pn_ee_fat_loc = st.radio(
                    "Fat stranding location",
                    ["none", "head_neck", "body", "neck_body", "perisplenic"],
                    horizontal=True,
                    format_func=lambda x: {"none": "None",
                                           "head_neck": "Head & neck",
                                           "body": "Body",
                                           "neck_body": "Neck & body",
                                           "perisplenic": "Peri-splenic"}[x],
                    key="pn_ee_fat_loc")

            pn_ee_fluid = st.checkbox("Mild peri-pancreatic free fluid",
                                      key="pn_ee_fluid")
            if pn_ee_fluid:
                pn_ee_fluid_loc = st.radio(
                    "Free fluid location",
                    ["none", "head_neck", "body", "neck_body", "perisplenic"],
                    horizontal=True,
                    format_func=lambda x: {"none": "None",
                                           "head_neck": "Head & neck",
                                           "body": "Body",
                                           "neck_body": "Neck & body",
                                           "perisplenic": "Peri-splenic"}[x],
                    key="pn_ee_fluid_loc")

        elif pn_status == "acute":
            pn_ac_size = st.radio(
                "Pancreas size",
                ["normal", "bulky"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal size",
                                       "bulky": "Bulky"}[x],
                key="pn_ac_size")

            pn_ac_echo = st.radio(
                "Echotexture",
                ["normal", "hypoechoic"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "hypoechoic": "Hypoechoic heterogeneous"}[x],
                key="pn_ac_echo")
            if pn_ac_echo == "hypoechoic":
                pn_ac_echo_loc = st.radio(
                    "Echotexture location",
                    ["none", "head_neck", "body"],
                    horizontal=True,
                    format_func=lambda x: {"none": "None (all)",
                                           "head_neck": "Head & neck",
                                           "body": "Body"}[x],
                    key="pn_ac_echo_loc")

            pn_ac_margins = st.radio(
                "Margins",
                ["normal", "irregular"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "irregular": "Irregular / fuzzy"}[x],
                key="pn_ac_margins")

            st.caption("Mild to moderate peri-pancreatic fat stranding and mild "
                       "free fluid are automatically included. Bowel free-fluid "
                       "line suppressed; Mild ascites will appear above impression.")

    # ------- SPLEEN -------
    col_sp_main, col_sp_sz = st.columns([5, 1], vertical_alignment="bottom")
    with col_sp_main:
        spleen_expander = st.expander("SPLEEN findings (click to open)",
                                      expanded=False)
    with col_sp_sz:
        sp_size_num = st.number_input(
            "Spleen size mm", min_value=0, max_value=500, value=None, step=1,
            key="spleen_size_num", label_visibility="collapsed",
            placeholder="SPLEEN mm")
    sp_size = str(int(sp_size_num)) if (sp_size_num and sp_size_num > 0) else ""
    sp_desc = auto_classify_spleen(sp_size_num or 0, p_age, p_sex)
    if sp_size_num and sp_size_num > 0:
        if sp_desc == "normal":
            st.success(f"✓ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")
        elif sp_desc == "enlarged_for_age":
            st.info(f"→ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")
        else:
            st.warning(f"→ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")
    with spleen_expander:
        st.caption("Spleen status is derived automatically from size.")

    # ------- KIDNEYS -------
    with st.expander("KIDNEYS (click to open findings)", expanded=False):
        kd_r_status = st.radio("Right kidney", ["normal", "not_visualized"],
                               horizontal=True, key="kd_r_status")
        kd_l_status = st.radio("Left kidney", ["normal", "not_visualized"],
                               horizontal=True, key="kd_l_status")

        kd_r_calc = ""
        kd_l_calc = ""
        c_a, c_b = st.columns(2)
        with c_a:
            kd_r_calc = st.text_area(
                "Right calculi (e.g. '4.6 upper-mid, 4.1 lower-mid')",
                height=60, key="kd_r_calc")
        with c_b:
            kd_l_calc = st.text_area(
                "Left calculi (e.g. '5.2 lower, 4.3 mid')",
                height=60, key="kd_l_calc")

    # ------- URINARY BLADDER -------
    with st.expander("URINARY BLADDER (click to open findings)", expanded=False):
        ub_status = st.radio("Status",
                             ["adequately_distended", "over", "partially", "empty"],
                             horizontal=True,
                             format_func=lambda x: x.replace("_", " ").title(),
                             key="ub_status")
        ub_sed = st.radio(
            "Sedimentation",
            ["none", "trace", "free_floating", "significant", "extensive"],
            horizontal=True,
            format_func=lambda x: x.replace("_", " ").title(), key="ub_sed")

    # ------- UTERUS / OVARIES or PROSTATE -------
    ut_status = "anteverted"
    ut_size = ""
    ut_et = ""
    ov_r = "normal"
    ov_r_size = ""
    ov_l = "normal"
    ov_l_size = ""
    pr_cc = ""
    pr_status = "normal"

    if p_sex == "F":
        u_exp_col, u_sz_col, u_et_col = st.columns(
            [4, 1, 1], vertical_alignment="bottom")
        with u_exp_col:
            uterus_expander = st.expander("UTERUS findings (click to open)",
                                          expanded=False)
        with u_sz_col:
            ut_size = st.text_input("UTERUS size (mm)", key="ut_size",
                                    placeholder="UTERUS mm",
                                    label_visibility="collapsed")
        with u_et_col:
            ut_et = st.text_input("Endometrium (mm)", key="ut_et",
                                  placeholder="Endometrium",
                                  label_visibility="collapsed")
        with uterus_expander:
            ut_status = st.radio("Status",
                                 ["anteverted", "retroverted", "bulky",
                                  "operated", "not_visualized"],
                                 horizontal=True, key="ut_status")

        o_exp_col, o_r_col, o_l_col = st.columns(
            [4, 1, 1], vertical_alignment="bottom")
        with o_exp_col:
            ovaries_expander = st.expander("OVARIES findings (click to open)",
                                           expanded=False)
        with o_r_col:
            ov_r_size = st.text_input("Rt Ovary (mm)", key="ov_r_size",
                                      placeholder="Rt Ovary",
                                      label_visibility="collapsed")
        with o_l_col:
            ov_l_size = st.text_input("LT Ovary (mm)", key="ov_l_size",
                                      placeholder="LT Ovary",
                                      label_visibility="collapsed")
        with ovaries_expander:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Right Ovary**")
                ov_r = st.radio("Status", ["normal", "cyst", "not_visualized"],
                                horizontal=True, key="ov_r_status")
            with c2:
                st.markdown("**Left Ovary**")
                ov_l = st.radio("Status", ["normal", "cyst", "not_visualized"],
                                horizontal=True, key="ov_l_status")
    else:
        p_exp_col, p_sz_col = st.columns([5, 1], vertical_alignment="bottom")
        with p_exp_col:
            prostate_expander = st.expander("PROSTATE findings (click to open)",
                                            expanded=False)
        with p_sz_col:
            pr_cc = st.text_input("PROSTATE size (cc)", key="pr_cc",
                                  placeholder="PROSTATE cc",
                                  label_visibility="collapsed")
        with prostate_expander:
            pr_status = st.radio("Status",
                                 ["normal", "borderline", "bulky", "grade1"],
                                 horizontal=True,
                                 format_func=lambda x: {"normal": "Normal",
                                                        "borderline": "Borderline",
                                                        "bulky": "Bulky",
                                                        "grade1": "Grade-I BPH"}[x],
                                 key="pr_status")

    # ------- BOWEL -------
    with st.expander("BOWEL / FREE FLUID (click to open findings)", expanded=False):
        bw_ff = st.radio("Free fluid",
                         ["none", "minimal", "mild", "moderate"],
                         horizontal=True, key="bw_ff")
        bw_ln = st.radio("Mesenteric LN", ["none", "present"],
                         horizontal=True, key="bw_ln")

    # ------- APPENDIX -------
    with st.expander("APPENDIX (optional) (click to open findings)", expanded=False):
        ap_status = st.radio("Status",
                             ["not_assessed", "not_visualized", "normal", "dilated"],
                             horizontal=True,
                             format_func=lambda x: x.replace("_", " ").title(),
                             key="ap_status")
        ap_d = ""
        if ap_status in ("normal", "dilated"):
            ap_d = st.text_input("Diameter (mm)", key="ap_d")


# ============================================================
# Assemble data
# ============================================================

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

data["liver"].update({
    "size_mm": liver_size, "size_descriptor": liver_status,
    "outline": liver_outline, "echotexture": liver_echo,
    "steatosis_grade": liver_steatosis,
    "focal_lesion": liver_focal, "focal_lesion_text": liver_focal_text,
    "cyst_count": cyst_count, "cyst_single_lobe": cyst_single_lobe,
    "cyst_single_size_mm": cyst_single_size_mm,
    "cyst_few_largest_mm": cyst_few_largest_mm, "cyst_few_lobe": cyst_few_lobe,
    "hemangioma_count": hemangioma_count,
    "hemangioma_single_lobe": hemangioma_single_lobe,
    "hemangioma_single_size_mm": hemangioma_single_size_mm,
    "hemangioma_few_largest_mm": hemangioma_few_largest_mm,
    "hemangioma_few_lobe": hemangioma_few_lobe,
    "abscess_count": abscess_count, "abscess_lesions": abscess_lesions,
    "ihbr": liver_ihbr, "portal_vein": liver_portal,
    "portal_vein_mm": liver_portal_mm if liver_portal == "dilated" else "",
})

data["gall_bladder"].update({
    "status": gb_status, "wall_thickened": gb_wall_thickened,
    "wall_mm": gb_wall_mm, "calculi": gb_calculi,
    "calculi_count": gb_calculi_count, "calculi_size_cat": gb_calculi_size_cat,
    "calculi_size_mm": gb_calculi_size_mm, "calculi_neck": gb_calculi_neck,
    "calculi_neck_size_mm": gb_calculi_neck_size_mm, "sludge": gb_sludge,
    "sludge_ball": gb_sludge_ball, "sludge_ball_count": gb_sludge_ball_count,
    "sludge_ball_size_mm": gb_sludge_ball_size_mm,
    "sludge_ball_wall": gb_sludge_ball_wall, "comet_tail": gb_comet_tail,
    "comet_tail_count": gb_comet_tail_count,
    "comet_tail_wall": gb_comet_tail_wall,
    "pericholecystic_fluid": gb_peri_fluid,
})

data["cbd"].update({
    "size_mm": cbd_mm, "status": cbd_status, "calculi": cbd_calc,
    "calculi_count": cbd_calc_count, "calculi_size_mm": cbd_calc_size,
    "calculi_location": cbd_calc_location, "ihbr": cbd_ihbr,
})

data["pancreas"].update({
    "status": pn_status,
    "ee_size": pn_ee_size, "ee_fat_stranding": pn_ee_fat,
    "ee_fat_location": pn_ee_fat_loc, "ee_free_fluid": pn_ee_fluid,
    "ee_fluid_location": pn_ee_fluid_loc,
    "ac_size": pn_ac_size, "ac_echo": pn_ac_echo,
    "ac_echo_location": pn_ac_echo_loc, "ac_margins": pn_ac_margins,
})

data["spleen"].update({"size_mm": sp_size, "size_descriptor": sp_desc})
data["kidneys"]["right"].update({"status": kd_r_status,
                                  "calculi": parse_calc(kd_r_calc)})
data["kidneys"]["left"].update({"status": kd_l_status,
                                 "calculi": parse_calc(kd_l_calc)})
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

auto_impression = generate_impression(data, p_sex, p_age)
data["impression"]["lines"] = auto_impression


def render_preview_findings(data):
    p = data["patient"]
    out = []
    out.append("         ULTRASOUND WHOLE ABDOMEN")
    out.append("")
    sex, age = p["sex"], p["age"]
    age_years = parse_age(age)
    is_ped = age_years is not None and age_years < 18
    p_status = data["pancreas"].get("status", "normal")
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
    secs.append(bowel_sentence(data["bowel"], sex, pancreas_status=p_status))
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
    return "\n".join(out)


preview_text = render_preview_findings(data)


with col_prev:
    st.subheader("📄 Live Preview")
    _safe = html.escape(preview_text).replace("\n", "<br>")
    st.markdown(
        f'<div class="preview-box">{_safe}</div>',
        unsafe_allow_html=True,
    )

    st.subheader("✏️ Impression (editable)")
    st.caption("Edit any line. Auto-updates with findings unless you type here.")

    auto_imp_str = "\n".join(auto_impression)

    def _h(s):
        return hashlib.md5(s.encode("utf-8")).hexdigest()

    if "_auto_imp_hash" not in st.session_state:
        st.session_state["impression_box"] = auto_imp_str
        st.session_state["_auto_imp_hash"] = _h(auto_imp_str)
        st.session_state["_last_set_content"] = auto_imp_str

    new_auto_hash = _h(auto_imp_str)
    if new_auto_hash != st.session_state["_auto_imp_hash"]:
        cur_widget = st.session_state.get("impression_box", "")
        if cur_widget == st.session_state["_last_set_content"]:
            st.session_state["impression_box"] = auto_imp_str
            st.session_state["_last_set_content"] = auto_imp_str
        st.session_state["_auto_imp_hash"] = new_auto_hash

    edited_imp = st.text_area("Impression lines (one per line)",
                              height=200, label_visibility="collapsed",
                              key="impression_box")

    if st.button("↺ Reset to auto-generated impression", key="reset_imp_btn"):
        st.session_state["impression_box"] = auto_imp_str
        st.session_state["_last_set_content"] = auto_imp_str
        st.session_state["_auto_imp_hash"] = _h(auto_imp_str)
        st.rerun()

    st.markdown("---")
    c_a, c_b = st.columns(2)
    with c_a:
        final_data = dict(data)
        if edited_imp.strip():
            final_data["impression"] = {
                "lines": [ln.strip() for ln in edited_imp.splitlines()
                          if ln.strip()]
            }
        docx_bytes = build_docx_bytes(final_data)
        fname = f"{p_name or 'report'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        fname = "".join(ch for ch in fname if ch.isalnum() or ch in "._-")
        st.download_button(
            "⬇️ Download .docx", docx_bytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.document")
    with c_b:
        if st.button("💾 Save to Database", key="save_db_btn"):
            if not p_name.strip():
                st.warning("Enter patient name first.")
            else:
                final_data = dict(data)
                if edited_imp.strip():
                    final_data["impression"] = {
                        "lines": [ln.strip() for ln in edited_imp.splitlines()
                                  if ln.strip()]
                    }
                save_report(final_data)
                if p_ref.strip():
                    add_referrer(p_ref)
                st.success("Saved to local database.")
