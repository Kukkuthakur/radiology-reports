"""
Radiology Report Generator — USG Whole Abdomen
v1.7.2-stable

v1.7.2: mirror-pattern fix. Widget keys are mirrored into plain session
  keys (_mirror_*) on every run, so on_change callbacks (e.g. from the
  addendum box) can never wipe collapsed organs' findings.

Frozen rules:
- Impression text ALL CAPS except "Adv- ... Correlation." and "(UB is empty)".
- Body findings bold, italic, mixed case.
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
# SESSION-STATE MIRROR
# ============================================================
# Streamlit prunes widget-owned session keys when the owning widget is not
# rendered on a rerun that is triggered by an on_change callback. To guard
# against that, we mirror every widget value we care about into a plain
# (non-widget) session key on every read. Plain keys are never pruned.

def _mirror(widget_key):
    mirror_key = f"_mirror_{widget_key}"
    if widget_key in st.session_state:
        st.session_state[mirror_key] = st.session_state[widget_key]
        return st.session_state[widget_key]
    return st.session_state.get(mirror_key)


def ss(widget_key, default):
    v = _mirror(widget_key)
    return v if v is not None else default


# ============================================================
# HELPERS
# ============================================================

def capitalize_sentences(text):
    if not text:
        return text
    out = []
    capitalize_next = True
    for ch in text:
        if capitalize_next and ch.isalpha():
            out.append(ch.upper())
            capitalize_next = False
        else:
            out.append(ch)
            if ch in ".!?\n":
                capitalize_next = True
    return "".join(out)


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


def _parse_mm_value(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", "."))
    except ValueError:
        return None


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
    elif size_mm < 121:
        return "borderline"
    elif size_mm < 140:
        return "mild"
    elif size_mm < 141:
        return "mild_to_moderate"
    elif size_mm <= 160:
        return "moderate"
    elif size_mm <= 170:
        return "moderate_to_gross"
    else:
        return "gross"


def classify_portal_vein(pv_mm):
    v = _parse_mm_value(pv_mm)
    if v is None:
        return "unknown"
    if v <= 13.0:
        return "normal"
    elif v < 14.0:
        return "prominent"
    else:
        return "dilated"


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
    "normal": "Normal size",
    "borderline": "Borderline enlarged",
    "mild": "Mildly enlarged",
    "mild_to_moderate": "Mild to moderately enlarged",
    "moderate": "Moderately enlarged",
    "moderate_to_gross": "Moderately to grossly enlarged",
    "gross": "Grossly enlarged",
    "enlarged_for_age": "Enlarged for age",
}
SPLEEN_IMPRESSION_LABELS = {
    "borderline": "BORDERLINE SPLENOMEGALY",
    "mild": "MILD SPLENOMEGALY",
    "mild_to_moderate": "MILD TO MODERATE SPLENOMEGALY",
    "moderate": "MODERATE SPLENOMEGALY",
    "moderate_to_gross": "MODERATE TO GROSS SPLENOMEGALY",
    "gross": "GROSS SPLENOMEGALY",
}

HEPATOSPLENOMEGALY_COMBINE = {
    "borderline": "BORDERLINE",
    "mild": "MILD",
    "moderate": "MODERATE",
    "gross": "GROSS",
    "mild_to_moderate": "MILD TO MODERATE",
    "moderate_to_gross": "MODERATE TO GROSS",
}


# ============================================================
# KIDNEY CONSTANTS
# ============================================================

URETER_LEVELS = {
    "renal_pelvis": ("in", "HYDRONEPHROSIS", "RENAL PELVIS",
                     "renal pelvis", "renal pelvis cal"),
    "puj": ("at", "HYDRONEPHROSIS", "PELVI-URETERIC JUNCTION",
            "pelvi-ureteric junction", "pelvi-ureteric junction cal"),
    "proximal_ureter": ("in", "HYDROURETERONEPHROSIS", "PROXIMAL URETERIC",
                        "proximal ureter", "proximal ureteric cal"),
    "mid_ureter": ("in", "HYDROURETERONEPHROSIS", "MID-URETERIC",
                   "mid-ureter", "mid-ureteric cal"),
    "distal_ureter": ("in", "HYDROURETERONEPHROSIS", "DISTAL URETERIC",
                      "distal ureter", "distal ureteric cal"),
    "vuj": ("at", "HYDROURETERONEPHROSIS", "VESICO-URETERIC JUNCTION",
            "vesico-ureteric junction", "vesico-ureteric junction cal"),
}

URETER_GRADES = {
    "none": None,
    "no_significant": "NO SIGNIFICANT",
    "minimal": "MINIMAL",
    "mild": "MILD",
    "moderate": "MODERATE",
}

URETER_GRADES_SENTENCE = {
    "no_significant": "no significant",
    "minimal": "minimal",
    "mild": "mild",
    "moderate": "moderate",
}

POLE_LABELS = {
    "upper": "upper-pole",
    "upper_mid": "upper-mid pole",
    "mid": "mid pole",
    "lower_mid": "lower-mid pole",
    "lower": "lower-pole",
}

BOSNIAK_OPTIONS = ["I", "II", "IIF"]


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

def _new_ureter_calculus():
    return {
        "count": "none",
        "sizes": [],
        "level": "distal_ureter",
        "grade": "none",
    }


def _new_kidney_side():
    return {
        "status": "normal",
        "size_text": "",
        "calculi": [],
        "ureter_calculus": _new_ureter_calculus(),
        "ureter_not_traced": False,
        "cyst": "none",
        "cyst_size_mm": "",
        "cyst_location": "",
        "bosniak": "",
        "hydronephrosis": "none",
        "hydronephrosis_no_obstructive_calculus": False,
        "recently_passed_calculus_suspected": False,
        "nephrocalcinosis": "none",
        "contralateral_normal_clause": False,
    }


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
            "wp_type": "won",
            "wp_dims": "",
            "wp_vol": "",
            "wp_location": "lesser_sac",
            "ch_foci": False,
            "ch_mpd": False,
            "ch_mpd_size": "",
            "ch_fat": False,
        },
        "spleen": {
            "size_mm": "", "size_descriptor": "normal",
            "focal_lesion": "none",
            "focal_count": "few",
            "portal_vein_mm": "",
            "accessory_spleen": False, "accessory_size_mm": "",
            "accessory_location": "hilum",
        },
        "kidneys": {
            "right": _new_kidney_side(),
            "left": _new_kidney_side(),
            "cortical_echogenicity": "normal",
            "cortical_echogenicity_laterality": "bilateral",
            "age_related_echogenicity": False,
            "negative_renal_line": False,
            "bilateral_ureter_calculi": False,
        },
        "urinary_bladder": {"status": "adequately_distended", "mass_calculus": False,
                            "sedimentation": "none",
                            "force_empty_suboptimal": False},
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
        "additional_body_findings": "",
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


def liver_sentence(d, sex, age, spleen_enlarged=False):
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

    if not spleen_enlarged:
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

    if status == "won_pseudocyst":
        wp_type = d.get("wp_type", "won")
        dims = d.get("wp_dims", "")
        vol = d.get("wp_vol", "")
        loc = d.get("wp_location", "lesser_sac")
        loc_phrase = "in the lesser sac" if loc == "lesser_sac" else "overlying the body of the pancreas"

        dims_vol = f"({dims}mm; Vol= {vol}cc)"

        if wp_type == "won":
            s.append(seg(" is "))
            s.append(seg("irregularly defined and obscured by a thick-walled collection",
                         True))
            s.append(seg(f"{dims_vol} {loc_phrase} with debris within and significant "
                         f"peripancreatic fat stranding and mild free fluid."))
        elif wp_type == "pseudocyst":
            s.append(seg(" appears "))
            s.append(seg("hypotrophied with irregular margins with a loculated collection",
                         True))
            s.append(seg(f"{dims_vol} with clear contents {loc_phrase} and mild "
                         f"peripancreatic fat stranding with mild free fluid."))
        else:
            s.append(seg(" appears "))
            s.append(seg("hypotrophied with irregular margins & a loculated mildly thick "
                         "walled collection", True))
            s.append(seg(f"{dims_vol} with internal echoes seen {loc_phrase} and mild "
                         f"peripancreatic fat stranding with mild free fluid."))
        return s

    if status == "chronic":
        foci = d.get("ch_foci", False)
        mpd = d.get("ch_mpd", False)
        mpd_size = d.get("ch_mpd_size", "")
        fat = d.get("ch_fat", False)

        s.append(seg(" is "))
        s.append(seg("poorly defined and appears hypotrophied", True))

        if foci and mpd and fat:
            s.append(seg(", studded with foci of calcification and dilated MPD"
                         f"(upto {mpd_size} mm). "))
            s.append(seg("Associated mild fat stranding is also present.", True))
        elif foci and mpd:
            s.append(seg(", studded with foci of calcification & dilated MPD"
                         f"(upto {mpd_size} mm). ", True))
            s.append(seg("However, no associated fat stranding appreciated."))
        elif foci and fat:
            s.append(seg(", studded with foci of calcification & reveals associated "
                         "mild fat stranding. ", True))
            s.append(seg("MPD is, however, not dilated."))
        elif mpd and fat:
            s.append(seg(f", with dilated MPD(upto {mpd_size} mm) & reveals associated "
                         f"mild fat stranding.", True))
        elif foci:
            s.append(seg(", studded with foci of calcification. ", True))
            s.append(seg("MPD is, however, not dilated."))
        elif mpd:
            s.append(seg(f", with dilated MPD(upto {mpd_size} mm). ", True))
            s.append(seg("However, no foci of calcification or associated fat stranding "
                         "appreciated."))
        elif fat:
            s.append(seg(", associated with mild fat stranding. ", True))
            s.append(seg("However, MPD is not dilated & no foci of calcification "
                         "appreciated."))
        else:
            s.append(seg("."))

        return s

    return s


# -------- SPLEEN --------

def spleen_sentence(d):
    s = [seg("SPLEEN", True, True)]
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]

    if desc == "normal":
        s.append(seg(f" is normal in size ({size}MM)"))
    elif desc == "borderline":
        s += [seg(" is "), seg("borderline enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "mild":
        s += [seg(" is "), seg("mildly enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "mild_to_moderate":
        s += [seg(" is "), seg("mild to moderately enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("moderately enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "moderate_to_gross":
        s += [seg(" is "), seg("moderately to grossly enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "gross":
        s += [seg(" is "), seg("grossly enlarged in size", True),
              seg(f" ({size}MM)")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("enlarged for age in size", True),
              seg(f" ({size}MM)")]

    s.append(seg(" with normal echotexture."))

    fl = d.get("focal_lesion", "none")
    cnt = d.get("focal_count", "few")

    if fl == "none":
        s.append(seg(" No focal lesion is seen."))
    elif fl == "hyperechoic_foci":
        if cnt == "multiple":
            s += [seg(" "), seg("Multiple hyperechoic foci seen scattered "
                                 "across the splenic parenchyma.", True)]
        else:
            s += [seg(" "), seg("Few hyperechoic foci scattered across splenic "
                                 "parenchyma.", True)]
    elif fl == "hypoechoic_foci":
        if cnt == "multiple":
            s += [seg(" "), seg("Multiple hypoechoic foci seen scattered "
                                 "across the splenic parenchyma.", True)]
        else:
            s += [seg(" "), seg("Few hypoechoic foci scattered across splenic "
                                 "parenchyma.", True)]

    enlarged = desc != "normal"

    if enlarged:
        pv_mm = d.get("portal_vein_mm", "")
        pv_class = classify_portal_vein(pv_mm) if pv_mm else "unknown"
        if pv_class == "prominent":
            pv_text = f" Portal vein is prominent in caliber ({pv_mm}MM)."
        elif pv_class == "dilated":
            pv_text = f" Portal vein is dilated in caliber ({pv_mm}MM)."
        else:
            pv_text = " Portal vein is normal in course and caliber"
            if pv_mm:
                pv_text += f" ({pv_mm}MM)"
            pv_text += "."
        s.append(seg(pv_text, True))
    else:
        s.append(seg(" Splenic vein is normal in course and caliber."))

    if d.get("accessory_spleen"):
        asize = d.get("accessory_size_mm", "")
        aloc = d.get("accessory_location", "hilum")
        loc_phrase = {"hilum": "at the splenic hilum",
                      "upper_pole": "at the upper pole",
                      "lower_pole": "at the lower pole"}.get(aloc, "at the splenic hilum")
        size_phrase = f"({asize}MM) " if asize else ""
        s += [seg(" "), seg(f"An accessory spleen {size_phrase}is seen "
                             f"{loc_phrase}.", True)]

    return s


# -------- KIDNEYS --------

def _kidney_side_is_present(k):
    return k["status"] == "normal"


def _both_kidneys_present(k):
    return (_kidney_side_is_present(k["right"])
            and _kidney_side_is_present(k["left"]))


def _size_bracket(k, presence):
    r_present, l_present = presence
    r = k["right"].get("size_text", "").strip()
    l = k["left"].get("size_text", "").strip()
    parts = []
    if r_present and r:
        parts.append(f"RK={r}MM")
    if l_present and l:
        parts.append(f"LK={l}MM")
    if not parts:
        return ""
    return "[" + ";".join(parts) + "]"


def _format_pole(pole_key):
    return POLE_LABELS.get(pole_key, pole_key.replace("_", "-"))


def _kidney_side_has_finding(k):
    if k["calculi"]:
        return True
    if k["ureter_calculus"]["count"] != "none":
        return True
    if k["cyst"] != "none":
        return True
    if k["hydronephrosis"] != "none":
        return True
    if k["nephrocalcinosis"] != "none":
        return True
    return False


def _renal_calculi_body(k, side_low):
    calcs = k["calculi"]
    if not calcs:
        return []
    n = len(calcs)
    out = [seg(" ")]
    if n == 1:
        c = calcs[0]
        out.append(seg(
            f"A calculus measuring {c['size_mm']}MM is seen at the "
            f"{_format_pole(c['pole'])} of {side_low} kidney", True))
        out.append(seg(".", True))
        return out

    groups = []
    for c in calcs:
        found = False
        for g in groups:
            if g["pole"] == c["pole"]:
                g["sizes"].append(c["size_mm"])
                found = True
                break
        if not found:
            groups.append({"pole": c["pole"], "sizes": [c["size_mm"]]})

    def group_text(g):
        sizes = g["sizes"]
        pole_txt = _format_pole(g["pole"])
        if len(sizes) == 1:
            return f"{sizes[0]}MM at the {pole_txt}"
        elif len(sizes) == 2:
            return f"{sizes[0]}MM & {sizes[1]}MM, both at the {pole_txt}"
        else:
            joined = ", ".join(f"{s}MM" for s in sizes[:-1])
            return f"{joined} & {sizes[-1]}MM, all at the {pole_txt}"

    parts = [group_text(g) for g in groups]
    if n == 2:
        lead = "A couple of calculi"
    else:
        lead = "Few calculi"
    if len(parts) == 1:
        joined = parts[0]
    elif len(parts) == 2:
        joined = f"{parts[0]} & {parts[1]}"
    else:
        joined = ", ".join(parts[:-1]) + " & " + parts[-1]
    out.append(seg(f"{lead} seen in the {side_low} kidney, largest of these "
                   f"measuring {joined}", True))
    out.append(seg(".", True))
    return out


def _ureter_side_label(side_key):
    return "RIGHT" if side_key == "right" else "LEFT"


def _ureter_calculus_body(d_side, side_key):
    uc = d_side["ureter_calculus"]
    if uc["count"] == "none":
        return []
    prep, term, imp_adj, body_level, _bi = URETER_LEVELS[uc["level"]]
    side_up = _ureter_side_label(side_key)
    side_low = side_key
    grade = uc["grade"]
    sizes = uc["sizes"]

    s = [seg(" ")]
    if uc["count"] == "single":
        size_txt = sizes[0] if sizes else ""
        s.append(seg(f"A calculus measuring {size_txt}MM seen {prep} the "
                     f"{side_low} {body_level}", True))
    elif uc["count"] == "couple":
        a = sizes[0] if len(sizes) > 0 else ""
        b = sizes[1] if len(sizes) > 1 else ""
        s.append(seg(f"A couple of calculi seen {prep} the {side_low} "
                     f"{body_level}, measuring {a}MM & {b}MM", True))
    elif uc["count"] in ("few", "multiple"):
        size_txt = sizes[0] if sizes else ""
        word = "Few" if uc["count"] == "few" else "Multiple"
        s.append(seg(f"{word} calculi seen {prep} the {side_low} "
                     f"{body_level}, largest of these measuring {size_txt}MM", True))

    if grade == "none":
        s.append(seg(".", True))
    elif grade == "no_significant":
        term_sentence = term.lower()
        s.append(seg(f", however causing no significant {term_sentence}", True))
        s.append(seg(".", True))
    else:
        g_word = URETER_GRADES_SENTENCE[grade]
        term_sentence = term.lower()
        s.append(seg(f", causing {side_low} sided {g_word} {term_sentence}", True))
        s.append(seg(".", True))
    return s


def _ureter_calculus_bilateral_body(k):
    r = k["right"]["ureter_calculus"]
    l = k["left"]["ureter_calculus"]
    if r["count"] == "none" or l["count"] == "none":
        return []

    def max_size(uc):
        nums = []
        for s in uc["sizes"]:
            v = _parse_mm_value(s)
            if v is not None:
                nums.append(v)
        return max(nums) if nums else 0

    if max_size(l) > max_size(r):
        first_side, second_side = "left", "right"
    else:
        first_side, second_side = "right", "left"

    def side_phrase(side_key):
        uc = k[side_key]["ureter_calculus"]
        _, _, _, _, bilateral_phrase = URETER_LEVELS[uc["level"]]
        rt_lt = "Rt" if side_key == "right" else "Lt"
        sizes_join = " & ".join(f"{s}mm" for s in uc["sizes"] if s)
        return f"{rt_lt} {bilateral_phrase} = {sizes_join}"

    p1 = side_phrase(first_side)
    p2 = side_phrase(second_side)

    def consequence(side_key):
        uc = k[side_key]["ureter_calculus"]
        _, term, _, _, _ = URETER_LEVELS[uc["level"]]
        grade = uc["grade"]
        side_low = side_key
        if grade == "none":
            return None
        if grade == "no_significant":
            return f"no significant {term.lower()} seen on the {side_low}"
        g_word = URETER_GRADES_SENTENCE[grade]
        return f"{side_low} {g_word} {term.lower()}"

    c1 = consequence(first_side)
    c2 = consequence(second_side)

    s = [seg(" ")]
    s.append(seg(f"Bilateral ureteric calculi are present ({p1} & {p2})", True))
    if c1 and c2:
        s.append(seg(f" causing {c1} & {c2}", True))
    elif c1 and not c2:
        s.append(seg(f" causing {c1}, while {c2}", True))
    elif c2 and not c1:
        s.append(seg(f" causing {c2}, while {c1}", True))
    s.append(seg(".", True))
    return s


def _cyst_body(k, side_low):
    if k["cyst"] == "none":
        return []
    loc = k.get("cyst_location", "")
    loc_txt = f" at the {_format_pole(loc)}" if loc else ""
    return [seg(" "), seg(f"A {k['cyst']} cyst measuring {k['cyst_size_mm']}MM "
                          f"is seen{loc_txt} of {side_low} kidney", True),
            seg(".", True)]


def _standalone_hydro_body(k, side_low, ub_status):
    if k["hydronephrosis"] == "none":
        return []
    grade = k["hydronephrosis"]
    term = "hydroureteronephrosis"
    s = [seg(" ")]
    s.append(seg(f"{side_low.capitalize()} {grade} {term} is present", True))
    if k.get("ureter_not_traced"):
        s.append(seg(f", {side_low} distal ureter could however not be traced", True))
        s.append(seg(" ", True))
        s.append(seg("(UB is empty)", True))
    elif k.get("hydronephrosis_no_obstructive_calculus"):
        s.append(seg(", however no obstructive calculus is seen upto the "
                     "visualized distal ureter", True))
    s.append(seg(".", True))
    return s


def _nephrocalcinosis_body(k, side_low):
    if k["nephrocalcinosis"] == "none":
        return []
    return [seg(" "), seg(f"Multiple foci of calcification seen in the "
                          f"{side_low} renal cortex", True), seg(".", True)]


def kidneys_sentence(d, ub_status="adequately_distended"):
    s = []
    r, l = d["right"], d["left"]
    r_present = _kidney_side_is_present(r)
    l_present = _kidney_side_is_present(l)

    if not r_present or not l_present:
        for side_key, side in (("right", r), ("left", l)):
            if side["status"] == "absent_agenesis":
                s.append(seg(f"{side_key.upper()} RENAL FOSSA", True, True))
                s.append(seg(" is empty"))
                s.append(seg(" "))
                s.append(seg("(agenesis/hypoplasia)", True))
                s.append(seg("."))
                s.append(seg("\n"))
            elif side["status"] == "absent_ectopic":
                s.append(seg(f"{side_key.upper()} RENAL FOSSA", True, True))
                s.append(seg(" is empty. "))
                s.append(seg(f"Ectopic {side_key} kidney", True))
                s.append(seg(f" seen lying in the {side_key} pelvic region."))
                s.append(seg("\n"))

        for side_key, side in (("right", r), ("left", l)):
            if not _kidney_side_is_present(side):
                continue
            heading = f"{side_key.upper()} KIDNEY"
            s.append(seg(heading, True, True))
            bracket = _size_bracket(d, (side_key == "right", side_key == "left"))
            if bracket:
                s.append(seg(f" is normal in size{bracket}, outline and echogenicity. "
                             f"Corticomedullary differentiation is maintained."))
            else:
                s.append(seg(" is normal in size, outline and echogenicity. "
                             "Corticomedullary differentiation is maintained."))
            s.extend(_renal_calculi_body(side, side_key))
            s.extend(_cyst_body(side, side_key))
            s.extend(_standalone_hydro_body(side, side_key, ub_status))
            s.extend(_nephrocalcinosis_body(side, side_key))

            if side["ureter_calculus"]["count"] != "none":
                s.extend(_ureter_calculus_body(side, side_key))
        return s

    s.append(seg("BOTH KIDNEYS", True, True))
    bracket = _size_bracket(d, (True, True))
    if bracket:
        s.append(seg(f" are normal in size{bracket}, outline and "))
    else:
        s.append(seg(" are normal in size, outline and "))
    if d["cortical_echogenicity"] == "mildly_raised":
        laterality = d.get("cortical_echogenicity_laterality", "bilateral")
        if laterality == "bilateral":
            s += [seg("mildly raised bilateral renal cortical echogenicity", True),
                  seg(". Corticomedullary differentiation is maintained.")]
        else:
            s += [seg(f"mildly raised {laterality} renal cortical echogenicity", True),
                  seg(". Corticomedullary differentiation is maintained.")]
    else:
        s.append(seg("echogenicity. Corticomedullary differentiation is maintained."))

    bilateral_uc = (d.get("bilateral_ureter_calculi", False)
                    and r["ureter_calculus"]["count"] != "none"
                    and l["ureter_calculus"]["count"] != "none")
    if bilateral_uc:
        s.extend(_ureter_calculus_bilateral_body(d))
        for side_key, side in (("right", r), ("left", l)):
            s.extend(_cyst_body(side, side_key))
            s.extend(_nephrocalcinosis_body(side, side_key))
        return s

    for side_key, side in (("right", r), ("left", l)):
        s.extend(_renal_calculi_body(side, side_key))
        s.extend(_cyst_body(side, side_key))
        s.extend(_standalone_hydro_body(side, side_key, ub_status))
        s.extend(_nephrocalcinosis_body(side, side_key))
        if side["ureter_calculus"]["count"] != "none":
            s.extend(_ureter_calculus_body(side, side_key))

    return s


# -------- URINARY BLADDER --------

def urinary_bladder_sentence(d):
    forced = d.get("force_empty_suboptimal", False)
    s = [seg("URINARY BLADDER", True, True)]
    if forced:
        s.append(seg(" is empty "))
        s.append(seg("(suboptimal pelvic assessment)", True))
        s.append(seg("."))
        if not d["mass_calculus"]:
            s.append(seg(" No mass or calculus seen."))
        return s

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


# -------- UTERUS / OVARIES / PROSTATE --------

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
    acute = (pancreas_status == "acute")
    s = []
    if sex == "F":
        if not d["wall_thickening"]:
            s.append(seg("No obvious bowel wall thickening or lymphadenitis appreciated.",
                         True))
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
            s.append(seg("No obvious bowel wall thickening or lymphadenitis appreciated.",
                         True))
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

def _ureter_impression_term(level):
    return URETER_LEVELS[level][1]


def _ureter_impression_adjective(level):
    return URETER_LEVELS[level][2]


def _ureter_impression_consequence(uc, side_up):
    grade = uc["grade"]
    term = _ureter_impression_term(uc["level"])
    if grade == "none":
        return ""
    if grade == "no_significant":
        return f" HOWEVER CAUSING NO SIGNIFICANT {term}"
    g = URETER_GRADES[grade]
    return f" CAUSING {side_up} SIDED {g} {term}"


def _ureter_calculus_impression_single(side_up, uc):
    adj = URETER_LEVELS[uc["level"]][2]
    sz = uc["sizes"][0] if uc["sizes"] else ""
    suffix = _ureter_impression_consequence(uc, side_up)
    return f"A {side_up} {adj} CALCULUS({sz}MM){suffix}"


def _ureter_calculus_impression_couple(side_up, uc):
    a = uc["sizes"][0] if len(uc["sizes"]) > 0 else ""
    b = uc["sizes"][1] if len(uc["sizes"]) > 1 else ""
    place = {
        "renal_pelvis": f"IN THE {side_up} RENAL PELVIS",
        "puj": f"AT THE {side_up} PELVI-URETERIC JUNCTION",
        "proximal_ureter": f"IN THE {side_up} PROXIMAL URETER",
        "mid_ureter": f"IN THE {side_up} MID-URETER",
        "distal_ureter": f"IN THE {side_up} DISTAL URETER",
        "vuj": f"AT THE {side_up} VESICO-URETERIC JUNCTION",
    }[uc["level"]]
    suffix = _ureter_impression_consequence(uc, side_up)
    return f"A COUPLE OF CALCULI({a}MM & {b}MM) {place}{suffix}"


def _ureter_calculus_impression_few(side_up, uc):
    word = "FEW" if uc["count"] == "few" else "MULTIPLE"
    sz = uc["sizes"][0] if uc["sizes"] else ""
    place = {
        "renal_pelvis": f"IN THE {side_up} RENAL PELVIS",
        "puj": f"AT THE {side_up} PELVI-URETERIC JUNCTION",
        "proximal_ureter": f"IN THE {side_up} PROXIMAL URETER",
        "mid_ureter": f"IN THE {side_up} MID-URETER",
        "distal_ureter": f"IN THE {side_up} DISTAL URETER",
        "vuj": f"AT THE {side_up} VESICO-URETERIC JUNCTION",
    }[uc["level"]]
    suffix = _ureter_impression_consequence(uc, side_up)
    return f"{word} CALCULI, LARGEST MEASURING {sz}MM, {place}{suffix}"


def _ureter_calculus_impression_line(side_key, uc):
    if uc["count"] == "none":
        return None
    side_up = _ureter_side_label(side_key)
    if uc["count"] == "single":
        return _ureter_calculus_impression_single(side_up, uc) + "."
    if uc["count"] == "couple":
        return _ureter_calculus_impression_couple(side_up, uc) + "."
    if uc["count"] in ("few", "multiple"):
        return _ureter_calculus_impression_few(side_up, uc) + "."
    return None


def _renal_calculi_clause(k, side_key):
    calcs = k["calculi"]
    if not calcs:
        return None
    side_up = _ureter_side_label(side_key)
    n = len(calcs)
    if n == 1:
        return f"A {side_up} RENAL CALCULUS"
    if n == 2:
        return f"A COUPLE OF {side_up} RENAL CALCULI"
    return f"FEW {side_up} RENAL CALCULI"


def _renal_calculi_standalone(k, side_key):
    calcs = k["calculi"]
    if not calcs:
        return None
    side_up = _ureter_side_label(side_key)
    n = len(calcs)
    if n == 1:
        return f"A {side_up} RENAL CALCULUS"
    if n == 2:
        return f"A COUPLE OF {side_up} RENAL CALCULI"
    return f"FEW {side_up} RENAL CALCULI"


def _cyst_impression_clause(k, side_key):
    if k["cyst"] == "none":
        return None
    side_up = _ureter_side_label(side_key)
    bosniak = k.get("bosniak", "")
    if not bosniak and k["cyst"] in ("cortical", "simple"):
        bosniak = "I"
    bosniak_txt = f" (BOSNIAK CAT-{bosniak})" if bosniak else ""
    return f"A {side_up} RENAL {k['cyst'].upper()} CYST{bosniak_txt}"


def _spleen_impression_line(spleen):
    desc = spleen["size_descriptor"]
    spleen_focal = spleen.get("focal_lesion", "none")
    spleen_count = spleen.get("focal_count", "few")

    if desc == "enlarged_for_age":
        spleno_term = "SPLENOMEGALY FOR AGE"
        is_enlarged = True
    elif desc in SPLEEN_IMPRESSION_LABELS:
        spleno_term = (f"{SPLEEN_IMPRESSION_LABELS[desc]} "
                       f"({spleen['size_mm']}MM)")
        is_enlarged = True
    else:
        spleno_term = ""
        is_enlarged = False

    out = []
    if is_enlarged:
        pv_mm = spleen.get("portal_vein_mm", "")
        pv_class = classify_portal_vein(pv_mm) if pv_mm else "unknown"

        focal_line = None
        if spleen_focal == "hyperechoic_foci":
            if spleen_count == "multiple":
                focal_line = ("STARRY SKY SPLEEN APPEARANCE - "
                              "?OLD GRANULOMATOUS ETIOLOGY.")
            else:
                focal_line = ("FEW HYPERECHOIC FOCI SCATTERED ACROSS SPLENIC "
                              "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")
        elif spleen_focal == "hypoechoic_foci":
            if spleen_count == "multiple":
                focal_line = ("MULTIPLE HYPOECHOIC FOCI SCATTERED ACROSS "
                              "SPLENIC PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")
            else:
                focal_line = ("FEW HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
                              "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")

        if pv_class == "prominent":
            pv_clause = f"PORTAL VEIN PROMINENT IN CALIBER ({pv_mm}MM)"
        elif pv_class == "dilated":
            if desc in ("moderate", "moderate_to_gross", "gross"):
                tail = "SUGGESTIVE OF PORTAL HYPERTENSION."
            else:
                tail = "? PORTAL HYPERTENSION."
            pv_clause = f"PORTAL VEIN DILATED IN CALIBER ({pv_mm}MM) - {tail}"
        elif pv_class == "normal":
            pv_clause = f"NORMAL CALIBER PORTAL VEIN ({pv_mm}MM)"
        else:
            pv_clause = None

        if focal_line and pv_clause and pv_class in ("normal", "prominent"):
            merged = f"{spleno_term} WITH {pv_clause} & {focal_line}"
            out.append(merged)
            return out

        if pv_class == "prominent":
            out.append(f"{spleno_term} WITH PORTAL VEIN PROMINENT IN "
                       f"CALIBER ({pv_mm}MM).")
        elif pv_class == "dilated":
            out.append(f"{spleno_term} WITH PORTAL VEIN DILATED IN CALIBER "
                       f"({pv_mm}MM) - {tail}")
        elif pv_class == "normal":
            out.append(f"{spleno_term} WITH NORMAL CALIBER PORTAL VEIN "
                       f"({pv_mm}MM).")
        else:
            out.append(f"{spleno_term}.")

        if focal_line:
            out.append(focal_line)
    else:
        if spleen_focal == "hyperechoic_foci":
            if spleen_count == "multiple":
                out.append("STARRY SKY SPLEEN APPEARANCE - ?OLD GRANULOMATOUS ETIOLOGY.")
            else:
                out.append("FEW HYPERECHOIC FOCI SCATTERED ACROSS SPLENIC "
                           "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")
        elif spleen_focal == "hypoechoic_foci":
            if spleen_count == "multiple":
                out.append("MULTIPLE HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
                           "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")
            else:
                out.append("FEW HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
                           "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")

    return out


def _liver_impression_line(liver):
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

    out = []
    if hepatomegaly and lf:
        out.append(hepatomegaly + " WITH " + " AND ".join(lf)
                   + ". Adv- LFT Correlation.")
    elif hepatomegaly:
        out.append(hepatomegaly + ". Adv- LFT Correlation.")
    elif lf:
        out.append(" AND ".join(lf) + ". Adv- LFT Correlation.")
    return out


def _try_hepatosplenomegaly(liver, spleen):
    ldesc = liver["size_descriptor"]
    sdesc = spleen["size_descriptor"]
    if ldesc == "normal" or sdesc == "normal":
        return None
    if ldesc not in HEPATOSPLENOMEGALY_COMBINE:
        return None
    if sdesc not in HEPATOSPLENOMEGALY_COMBINE:
        return None
    if HEPATOSPLENOMEGALY_COMBINE[ldesc] != HEPATOSPLENOMEGALY_COMBINE[sdesc]:
        return None
    if liver["focal_lesion"] != "none":
        return None
    if spleen.get("focal_lesion", "none") != "none":
        return None
    desc_word = HEPATOSPLENOMEGALY_COMBINE[ldesc]
    base = f"{desc_word} HEPATOSPLENOMEGALY"
    liver_extras = []
    if liver["echotexture"] == "increased":
        grade = liver.get("steatosis_grade") or ""
        if grade:
            liver_extras.append(f"HEPATIC STEATOSIS ({grade.upper()})")
        else:
            liver_extras.append("HEPATIC STEATOSIS")
    pv_mm = spleen.get("portal_vein_mm", "")
    pv_class = classify_portal_vein(pv_mm) if pv_mm else "unknown"
    pv_clause = None
    if pv_class == "normal":
        pv_clause = f"NORMAL CALIBER PORTAL VEIN ({pv_mm}MM)"
    elif pv_class == "prominent":
        pv_clause = f"PORTAL VEIN PROMINENT IN CALIBER ({pv_mm}MM)"
    elif pv_class == "dilated":
        return None

    body_parts = []
    if liver_extras:
        body_parts.append(" WITH " + " AND ".join(liver_extras))
    if pv_clause:
        body_parts.append(" & " + pv_clause)
    combined = base + "".join(body_parts) + ". Adv- LFT Correlation."
    return combined


def generate_impression(d, sex, age):
    lines = []
    age_years = parse_age(age)
    is_pediatric = age_years is not None and age_years < 18

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

    elif p_status == "won_pseudocyst":
        wp_type = p.get("wp_type", "won")
        vol = p.get("wp_vol", "")
        loc = p.get("wp_location", "lesser_sac")
        loc_caps = "IN THE LESSER SAC" if loc == "lesser_sac" else "OVERLYING THE BODY OF THE PANCREAS"

        if wp_type == "won":
            lines.append(
                f"IRREGULARLY DEFINED PANCREAS OBSCURED BY A THICK-WALLED COLLECTION "
                f"{loc_caps} (VOL= {vol}CC) WITH DEBRIS WITHIN AND SIGNIFICANT "
                f"PERIPANCREATIC FAT STRANDING AND MILD FREE FLUID - LIKELY "
                f"WALLED-OFF NECROSIS (WON) AS A SEQUELAE TO ACUTE PANCREATITIS. "
                f"Adv- S.Amylase/Lipase Correlation."
            )
        elif wp_type == "pseudocyst":
            lines.append(
                f"HYPOTROPHIED PANCREAS WITH IRREGULAR MARGINS WITH A LOCULATED "
                f"COLLECTION (VOL= {vol}CC) WITH CLEAR CONTENTS {loc_caps} AND MILD "
                f"PERIPANCREATIC FAT STRANDING WITH MILD FREE FLUID - LIKELY "
                f"PANCREATIC PSEUDOCYST AS A SEQUELAE TO ACUTE PANCREATITIS. "
                f"Adv- S.Amylase/Lipase Correlation."
            )
        else:
            lines.append(
                f"HYPOTROPHIED PANCREAS WITH IRREGULAR MARGINS & A LOCULATED MILDLY "
                f"THICK WALLED COLLECTION (VOL= {vol}CC) WITH INTERNAL ECHOES "
                f"{loc_caps} AND MILD PERIPANCREATIC FAT STRANDING WITH MILD FREE "
                f"FLUID - LIKELY WON/PANCREATIC PSEUDOCYST AS A SEQUELAE TO ACUTE "
                f"PANCREATITIS. Adv- S.Amylase/Lipase Correlation."
            )

    elif p_status == "chronic":
        foci = p.get("ch_foci", False)
        mpd = p.get("ch_mpd", False)
        fat = p.get("ch_fat", False)

        if foci and mpd and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif foci and mpd:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif foci and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif mpd and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS / "
                         "SEQUELAE TO ACUTE PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif foci:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif mpd:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS / "
                         "SEQUELAE TO ACUTE PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        else:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")

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

    liver = d["liver"]
    spleen = d["spleen"]

    combined = _try_hepatosplenomegaly(liver, spleen)
    if combined:
        lines.append(combined)
    else:
        for ln in _liver_impression_line(liver):
            lines.append(ln)
        for ln in _spleen_impression_line(spleen):
            lines.append(ln)

    k = d["kidneys"]
    r = k["right"]
    l = k["left"]
    r_present = _kidney_side_is_present(r)
    l_present = _kidney_side_is_present(l)

    for side_key, side in (("right", r), ("left", l)):
        if side["status"] == "absent_agenesis":
            lines.append(f"EMPTY {side_key.upper()} RENAL FOSSA - "
                         f"?AGENESIS/HYPOPLASIA.")
        elif side["status"] == "absent_ectopic":
            lines.append(f"EMPTY {side_key.upper()} RENAL FOSSA - ECTOPIC KIDNEY "
                         f"LYING IN {side_key.upper()} PELVIS.")

    r_uc = r["ureter_calculus"]
    l_uc = l["ureter_calculus"]
    bilateral_flag = k.get("bilateral_ureter_calculi", False)

    renal_lines = []

    cyst_clauses = []
    for side_key, side in (("right", r), ("left", l)):
        if _kidney_side_is_present(side):
            clause = _cyst_impression_clause(side, side_key)
            if clause:
                cyst_clauses.append(clause)

    if (bilateral_flag
            and r_uc["count"] != "none"
            and l_uc["count"] != "none"):
        def side_desc(side_key, uc):
            rt_lt = "Rt" if side_key == "right" else "Lt"
            _, _, _, _, bi = URETER_LEVELS[uc["level"]]
            sizes_join = " & ".join(f"{s}mm" for s in uc["sizes"] if s)
            return f"{rt_lt} {bi} = {sizes_join}"

        def max_size(uc):
            nums = []
            for s in uc["sizes"]:
                v = _parse_mm_value(s)
                if v is not None:
                    nums.append(v)
            return max(nums) if nums else 0

        if max_size(l_uc) > max_size(r_uc):
            first, second = "left", "right"
        else:
            first, second = "right", "left"
        p1 = side_desc(first, k[first]["ureter_calculus"])
        p2 = side_desc(second, k[second]["ureter_calculus"])

        def cons(side_key, uc):
            _, term, _, _, _ = URETER_LEVELS[uc["level"]]
            grade = uc["grade"]
            if grade == "none":
                return None
            if grade == "no_significant":
                return f"no significant {term.lower()} seen on the {side_key}"
            g = URETER_GRADES[grade]
            return f"{side_key} {g.lower()} {term.lower()}"

        c1 = cons(first, k[first]["ureter_calculus"])
        c2 = cons(second, k[second]["ureter_calculus"])
        if c1 and c2:
            imp = (f"BILATERAL URETERIC CALCULI ({p1} & {p2}) CAUSING "
                   f"{c1.upper()} & {c2.upper()}.")
        elif c1:
            imp = (f"BILATERAL URETERIC CALCULI ({p1} & {p2}) CAUSING "
                   f"{c1.upper()}, WHILE {c2.upper()}.")
        elif c2:
            imp = (f"BILATERAL URETERIC CALCULI ({p1} & {p2}) CAUSING "
                   f"{c2.upper()}, WHILE {c1.upper()}.")
        else:
            imp = f"BILATERAL URETERIC CALCULI ({p1} & {p2})."
        renal_lines.append(imp)
    else:
        for side_key in ("right", "left"):
            uc = k[side_key]["ureter_calculus"]
            if uc["count"] == "none":
                continue
            line = _ureter_calculus_impression_line(side_key, uc)
            if line:
                r_renal = _renal_calculi_clause(r, "right") if _kidney_side_is_present(r) else None
                l_renal = _renal_calculi_clause(l, "left") if _kidney_side_is_present(l) else None
                renal_clause = None
                if r_renal and l_renal:
                    renal_clause = "BILATERAL RENAL CALCULI"
                elif r_renal:
                    renal_clause = r_renal
                elif l_renal:
                    renal_clause = l_renal
                if renal_clause:
                    base = line.rstrip(".")
                    line = base + " & " + renal_clause + "."
                renal_lines.append(line)

    has_ureter_line = any(
        ("URETERIC CALCULUS" in ln or "URETERIC CALCULI" in ln
         or "URETER CALCULUS" in ln or "RENAL PELVIS CALCULUS" in ln
         or "PELVI-URETERIC JUNCTION CALCULUS" in ln
         or "VESICO-URETERIC JUNCTION CALCULUS" in ln
         or "BILATERAL URETERIC CALCULI" in ln)
        for ln in renal_lines)
    if not has_ureter_line:
        r_calcs = r["calculi"] if _kidney_side_is_present(r) else []
        l_calcs = l["calculi"] if _kidney_side_is_present(l) else []
        no_hydro = (r["hydronephrosis"] == "none" and l["hydronephrosis"] == "none")
        if r_calcs and l_calcs:
            suffix = (", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN."
                      if no_hydro else ".")
            renal_lines.append(f"BILATERAL RENAL CALCULI{suffix}")
        elif r_calcs:
            phrase = _renal_calculi_standalone(r, "right")
            suffix = (", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN."
                      if no_hydro else ".")
            renal_lines.append(f"{phrase}{suffix}")
        elif l_calcs:
            phrase = _renal_calculi_standalone(l, "left")
            suffix = (", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN."
                      if no_hydro else ".")
            renal_lines.append(f"{phrase}{suffix}")

    for side_key, side in (("right", r), ("left", l)):
        if not _kidney_side_is_present(side):
            continue
        if side["hydronephrosis"] == "none":
            continue
        if side["ureter_calculus"]["count"] != "none":
            continue
        side_up = _ureter_side_label(side_key)
        g = side["hydronephrosis"].upper()
        extra = ""
        if side.get("ureter_not_traced"):
            extra = (f", {side_up} DISTAL URETER COULD HOWEVER NOT BE TRACED "
                     f"(UB is empty)")
        elif side.get("hydronephrosis_no_obstructive_calculus"):
            extra = (", HOWEVER NO OBSTRUCTIVE CALCULUS IS SEEN UPTO THE "
                     "VISUALIZED DISTAL URETER")
        rpc = ""
        if side.get("recently_passed_calculus_suspected"):
            rpc = " - ?RECENTLY PASSED CALCULUS"
        renal_lines.append(f"{g} {side_up} HYDROURETERONEPHROSIS IS PRESENT"
                           f"{extra}{rpc}.")

    if k["cortical_echogenicity"] == "mildly_raised":
        laterality = k.get("cortical_echogenicity_laterality", "bilateral")
        lat_txt = {"bilateral": "BILATERAL", "right": "RIGHT",
                   "left": "LEFT"}.get(laterality, "BILATERAL")
        tail = ("- ?AGE RELATED. Adv- KFT Correlation"
                if k.get("age_related_echogenicity")
                else ". Adv- KFT Correlation")
        renal_lines.append(f"MILDLY RAISED {lat_txt} RENAL CORTICAL ECHOGENICITY{tail}")

    for side_key, side in (("right", r), ("left", l)):
        if not _kidney_side_is_present(side):
            continue
        if not side.get("contralateral_normal_clause"):
            continue
        other = "left" if side_key == "right" else "right"
        other_up = other.upper()
        if side["calculi"]:
            for i in range(len(renal_lines) - 1, -1, -1):
                if ("RENAL CALCULUS" in renal_lines[i]
                        or "RENAL CALCULI" in renal_lines[i]
                        or "NEPHROLITHIASIS" in renal_lines[i]):
                    base = renal_lines[i].rstrip(".")
                    base = base.replace(
                        ", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN", "")
                    base = base.rstrip("., ")
                    renal_lines[i] = (base
                                      + f", HOWEVER NO CALCULUS/HYDRONEPHROSIS SEEN "
                                        f"ON THE {other_up} KIDNEY AT THE TIME OF SCAN.")
                    break

    if k.get("negative_renal_line"):
        renal_lines.append("NO EVIDENCE OF HYDRONEPHROTIC CHANGES/CALCULUS SEEN AT "
                           "THE TIME OF SCAN.")

    if renal_lines and cyst_clauses:
        joined_main = " ".join(renal_lines)
        if not joined_main.endswith("."):
            joined_main += "."
        cyst_phrase = " ".join(c + " IS ALSO SEEN." for c in cyst_clauses)
        lines.append(joined_main + " " + cyst_phrase)
    elif renal_lines:
        for rl in renal_lines:
            lines.append(rl)
    elif cyst_clauses:
        for c in cyst_clauses:
            lines.append(c + ".")

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
    tblPr = tbl.tblPr
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
    spleen_enlarged = data["spleen"]["size_descriptor"] != "normal"
    ub_status = data["urinary_bladder"].get("status", "adequately_distended")

    sections = [
        liver_sentence(data["liver"], sex, age, spleen_enlarged=spleen_enlarged),
        gall_bladder_sentence(data["gall_bladder"]),
        cbd_sentence(data["cbd"]),
        pancreas_sentence(data["pancreas"]),
        spleen_sentence(data["spleen"]),
        kidneys_sentence(data["kidneys"], ub_status=ub_status),
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

    addendum = (data.get("additional_body_findings") or "").strip()
    if addendum:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        lines = addendum.split("\n")
        for i, ln in enumerate(lines):
            if i > 0:
                para.add_run().add_break()
            _add_run(para, ln, bold=True, italic=True)
        para.paragraph_format.space_after = Pt(6)

    doc.add_paragraph()
    imp_head = doc.add_paragraph()
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True,
             size=FONT_SIZE_BODY, color=TITLE_COLOR)

    for line in data["impression"]["lines"]:
        para = doc.add_paragraph(style="List Bullet")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if "(UB is empty)" in line:
            parts = line.split("(UB is empty)")
            _add_run(para, parts[0], bold=True)
            _add_run(para, "(UB is empty)", bold=True, italic=True)
            if len(parts) > 1:
                _add_run(para, parts[1], bold=True)
        else:
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
    .stApp li { color: #111111; }
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
    div[class*="st-key-organ_btn_"] button {
        justify-content: flex-start !important;
        text-align: left !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 10px 14px !important;
        min-height: 44px !important;
        border-radius: 6px !important;
        border: 1px solid #cfd6dd !important;
        background-color: #ffffff !important;
        color: #111111 !important;
        letter-spacing: 0.3px !important;
    }
    div[class*="st-key-organ_btn_"] button[kind="primary"] {
        background-color: #eef2f7 !important;
        border-color: #305496 !important;
        color: #1F4E79 !important;
    }
    div[class*="st-key-organ_btn_"] button p {
        text-align: left !important; width: 100% !important;
        font-weight: 600 !important;
    }
    .preview-box, .stApp .preview-box, .stApp .preview-box * {
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

if "open_organ" not in st.session_state:
    st.session_state.open_organ = None


def organ_button(name, label):
    is_open = (st.session_state.open_organ == name)
    arrow = "▼" if is_open else "▶"
    if st.button(f"{arrow}  {label}",
                 key=f"organ_btn_{name}",
                 use_container_width=True,
                 type="primary" if is_open else "secondary"):
        st.session_state.open_organ = None if is_open else name
        st.rerun()
    return is_open


with col_find:
    st.subheader("Findings")

    # ------- LIVER -------
    col_liver_main, col_liver_sz = st.columns([5, 1],
                                              vertical_alignment="bottom")
    with col_liver_main:
        liver_open = organ_button("LIVER", "LIVER")
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

    if liver_open:
        with st.container(border=True):
            st.radio(
                "Outline", ["normal", "crenated"], horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "crenated": "Crenated / nodular"}[x],
                key="liver_outline")
            st.radio(
                "Echotexture", ["normal", "increased", "coarse", "low"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "increased": "Increased (steatosis)",
                                       "coarse": "Coarse", "low": "Low"}[x],
                key="liver_echo")
            if ss("liver_echo", "normal") == "increased":
                grade_opts = ["Mild+", "Mild to Moderate++", "Moderate++",
                              "Moderate to Severe+++", "Severe+++"]
                st.radio("Steatosis grade (impression only)",
                         grade_opts, horizontal=True,
                         key="liver_steatosis")
            st.markdown("**Focal lesion**")
            st.radio(
                "Focal lesion type",
                ["none", "calcified", "cyst", "hemangioma", "abscess", "other"],
                horizontal=True, label_visibility="collapsed",
                format_func=lambda x: {"none": "None", "calcified": "Calcified",
                                       "cyst": "Simple cyst(s)",
                                       "hemangioma": "Hemangioma(s)",
                                       "abscess": "Abscess(es)", "other": "Other"}[x],
                key="liver_focal")
            _lf = ss("liver_focal", "none")
            if _lf == "cyst":
                st.radio("Number", ["single", "few"], horizontal=True,
                         key="cyst_count")
                if ss("cyst_count", "single") == "single":
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.radio("Lobe", ["right", "left"],
                                 horizontal=True, key="cyst_single_lobe")
                    with c_b:
                        st.text_input("Size (mm)", key="cyst_single_size")
                else:
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.radio("Lobe", ["right", "left"],
                                 horizontal=True, key="cyst_few_lobe")
                    with c_b:
                        st.text_input("Largest (mm)", key="cyst_few_largest")
            elif _lf == "hemangioma":
                st.radio("Number", ["single", "few"],
                         horizontal=True, key="hemangioma_count")
                if ss("hemangioma_count", "single") == "single":
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.radio("Lobe", ["right", "left"], horizontal=True,
                                 key="hemangioma_single_lobe")
                    with c_b:
                        st.text_input("Size (mm)", key="hemangioma_single_size")
                else:
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.radio("Lobe", ["right", "left"], horizontal=True,
                                 key="hemangioma_few_lobe")
                    with c_b:
                        st.text_input("Largest (mm)", key="hemangioma_few_largest")
            elif _lf == "abscess":
                st.radio("Number", ["single", "few", "multiple"],
                         horizontal=True, key="abscess_count")
                _ac = ss("abscess_count", "single")
                n_abs = 1 if _ac == "single" else int(st.number_input(
                    "How many lesions?", min_value=1, max_value=5, value=2,
                    key="abscess_n"))
                for i in range(n_abs):
                    st.markdown(f"*Lesion {i+1}*")
                    c_a, c_b, c_c = st.columns([1, 2, 1])
                    with c_a:
                        st.selectbox("Segment",
                                     ["I", "II", "III", "IV",
                                      "V", "VI", "VII", "VIII"],
                                     key=f"abs_seg_{i}")
                    with c_b:
                        st.text_input("Dimensions (XxYxZ mm)", key=f"abs_dim_{i}")
                    with c_c:
                        st.text_input("Volume (cc)", key=f"abs_vol_{i}")
            elif _lf == "other":
                st.text_input("Description", key="liver_focal_text")

            st.radio("IHBR", ["normal", "dilated"], horizontal=True,
                     key="liver_ihbr")
            c_a, c_b = st.columns([2, 1])
            with c_a:
                st.radio("Portal vein", ["normal", "dilated"],
                         horizontal=True, key="liver_portal")
            with c_b:
                if ss("liver_portal", "normal") == "dilated":
                    st.text_input("Portal vein size (mm)", key="liver_portal_mm")

    # ------- GALL BLADDER -------
    gb_open = organ_button("GALL BLADDER", "GALL BLADDER")
    if gb_open:
        with st.container(border=True):
            st.radio(
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
            _gbs = ss("gb_status", "adequately_distended")
            if _gbs not in ("contracted", "operated"):
                st.checkbox("Wall thickening present", key="gb_wall_check")
                if ss("gb_wall_check", False):
                    st.text_input("Wall thickness (mm)", key="gb_wall_mm")
            st.radio("Calculi", ["none", "present"], horizontal=True,
                     key="gb_calculi")
            if ss("gb_calculi", "none") == "present":
                c_a, c_b = st.columns(2)
                with c_a:
                    st.radio("Count",
                             ["single", "few", "multiple", "innumerable"],
                             horizontal=True, format_func=lambda x: x.title(),
                             key="gb_calc_count")
                with c_b:
                    if ss("gb_calc_count", "single") != "innumerable":
                        st.radio("Size", ["small", "large"], horizontal=True,
                                 format_func=lambda x: x.title(),
                                 key="gb_calc_size_cat")
                if ss("gb_calc_count", "single") != "innumerable":
                    st.text_input("Largest size (mm)", key="gb_calc_size")
                    st.checkbox("Calculus at GB neck", key="gb_calc_neck")
                    if ss("gb_calc_neck", False):
                        st.text_input("Neck calculus size (mm)", key="gb_neck_size")
            st.radio("Sludge",
                     ["none", "trace", "significant", "echogenic", "organized"],
                     horizontal=True, format_func=lambda x: x.title(),
                     key="gb_sludge")
            st.checkbox("Sludge ball / polyp present", key="gb_sludge_ball")
            if ss("gb_sludge_ball", False):
                c_a, c_b = st.columns(2)
                with c_a:
                    st.radio("Count", ["single", "few", "multiple"],
                             horizontal=True, key="gb_sb_count")
                with c_b:
                    st.radio("Wall", ["anterior", "posterior"],
                             horizontal=True, format_func=lambda x: x.title(),
                             key="gb_sb_wall")
                st.text_input("Size (mm)", key="gb_sb_size")
            st.checkbox(
                "Comet tail artifacts (adenomyomatosis / cholesterolosis)",
                key="gb_comet_tail")
            if ss("gb_comet_tail", False):
                c_a, c_b = st.columns(2)
                with c_a:
                    st.radio("Count", ["single", "few", "multiple"],
                             horizontal=True, key="gb_ct_count")
                with c_b:
                    st.radio("Wall", ["anterior", "posterior"],
                             horizontal=True, format_func=lambda x: x.title(),
                             key="gb_ct_wall")
            st.checkbox("Thin rim of pericholecystic fluid present",
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
        cbd_open = organ_button("COMMON BILE DUCT", "COMMON BILE DUCT")
    if cbd_open:
        with st.container(border=True):
            st.radio(
                "Status", ["normal", "proximal", "dilated"], horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "proximal": "Proximally dilated",
                                       "dilated": "Dilated throughout"}[x],
                key="cbd_status")
            st.checkbox("Calculus in CBD", key="cbd_calc")
            if ss("cbd_calc", False):
                c_a, c_b = st.columns(2)
                with c_a:
                    st.radio("Count", ["single", "few"], horizontal=True,
                             format_func=lambda x: x.title(),
                             key="cbd_calc_count")
                with c_b:
                    st.text_input("Size (mm, largest if few)",
                                  key="cbd_calc_size")
                st.radio("Location",
                         ["proximal", "mid", "distal", "mid_distal"],
                         horizontal=True,
                         format_func=lambda x: {"proximal": "Proximal",
                                                "mid": "Mid",
                                                "distal": "Distal",
                                                "mid_distal": "Mid/Distal"}[x],
                         key="cbd_calc_location")
            if ss("cbd_status", "normal") in ("proximal", "dilated") \
                    or ss("cbd_calc", False):
                st.radio("IHBR", ["normal", "proximal", "dilated"],
                         horizontal=True,
                         format_func=lambda x: {"normal": "Normal",
                                                "proximal": "Proximally dilated",
                                                "dilated": "Dilated"}[x],
                         key="cbd_ihbr")

    # ------- PANCREAS -------
    pn_open = organ_button("PANCREAS", "PANCREAS")
    if pn_open:
        with st.container(border=True):
            st.radio(
                "Status",
                ["early_evolving", "acute", "won_pseudocyst", "chronic"],
                index=None,
                horizontal=True,
                format_func=lambda x: {
                    "early_evolving": "Early/evolving pancreatitis",
                    "acute": "Acute pancreatitis",
                    "won_pseudocyst": "WON / Pseudocyst",
                    "chronic": "Chronic pancreatitis"}[x],
                key="pn_status")
            _pn = ss("pn_status", None)
            if _pn == "early_evolving":
                st.radio(
                    "Pancreas size",
                    ["normal", "mildly_bulky"],
                    horizontal=True,
                    format_func=lambda x: {"normal": "Normal size",
                                           "mildly_bulky": "Mildly bulky"}[x],
                    key="pn_ee_size")
                st.checkbox("Mild peri-pancreatic fat stranding", key="pn_ee_fat")
                if ss("pn_ee_fat", False):
                    st.radio(
                        "Fat stranding location",
                        ["none", "head_neck", "body", "neck_body", "perisplenic"],
                        horizontal=True,
                        format_func=lambda x: {"none": "None",
                                               "head_neck": "Head & neck",
                                               "body": "Body",
                                               "neck_body": "Neck & body",
                                               "perisplenic": "Peri-splenic"}[x],
                        key="pn_ee_fat_loc")
                st.checkbox("Mild peri-pancreatic free fluid", key="pn_ee_fluid")
                if ss("pn_ee_fluid", False):
                    st.radio(
                        "Free fluid location",
                        ["none", "head_neck", "body", "neck_body", "perisplenic"],
                        horizontal=True,
                        format_func=lambda x: {"none": "None",
                                               "head_neck": "Head & neck",
                                               "body": "Body",
                                               "neck_body": "Neck & body",
                                               "perisplenic": "Peri-splenic"}[x],
                        key="pn_ee_fluid_loc")
            elif _pn == "acute":
                st.radio(
                    "Pancreas size",
                    ["normal", "bulky"],
                    horizontal=True,
                    format_func=lambda x: {"normal": "Normal size",
                                           "bulky": "Bulky"}[x],
                    key="pn_ac_size")
                st.radio(
                    "Echotexture",
                    ["normal", "hypoechoic"],
                    horizontal=True,
                    format_func=lambda x: {"normal": "Normal",
                                           "hypoechoic": "Hypoechoic heterogeneous"}[x],
                    key="pn_ac_echo")
                if ss("pn_ac_echo", "normal") == "hypoechoic":
                    st.radio(
                        "Echotexture location",
                        ["none", "head_neck", "body"],
                        horizontal=True,
                        format_func=lambda x: {"none": "None (all)",
                                               "head_neck": "Head & neck",
                                               "body": "Body"}[x],
                        key="pn_ac_echo_loc")
                st.radio(
                    "Margins",
                    ["normal", "irregular"],
                    horizontal=True,
                    format_func=lambda x: {"normal": "Normal",
                                           "irregular": "Irregular / fuzzy"}[x],
                    key="pn_ac_margins")
                st.caption("Mild to moderate peri-pancreatic fat stranding and "
                           "mild free fluid are automatically included.")
            elif _pn == "won_pseudocyst":
                st.radio(
                    "Type",
                    ["won", "pseudocyst", "won_pseudocyst"],
                    horizontal=True,
                    format_func=lambda x: {"won": "WON",
                                           "pseudocyst": "Pseudocyst",
                                           "won_pseudocyst": "WON + Pseudocyst"}[x],
                    key="pn_wp_type")
                c_a, c_b = st.columns(2)
                with c_a:
                    st.text_input("Dimensions (e.g. 65x54x36)", key="pn_wp_dims")
                with c_b:
                    st.text_input("Volume (cc)", key="pn_wp_vol")
                st.radio(
                    "Location",
                    ["lesser_sac", "overlying_body"],
                    horizontal=True,
                    format_func=lambda x: {"lesser_sac": "In the lesser sac",
                                           "overlying_body": "Overlying the body of the pancreas"}[x],
                    key="pn_wp_location")
            elif _pn == "chronic":
                st.checkbox("Foci of calcification", key="pn_ch_foci")
                st.checkbox("MPD dilation", key="pn_ch_mpd")
                if ss("pn_ch_mpd", False):
                    st.text_input("MPD size (mm)", key="pn_ch_mpd_size")
                st.checkbox("Mild fat stranding", key="pn_ch_fat")

    # ------- SPLEEN -------
    col_sp_main, col_sp_sz = st.columns([5, 1], vertical_alignment="bottom")
    with col_sp_main:
        spleen_open = organ_button("SPLEEN", "SPLEEN")
    with col_sp_sz:
        sp_size_num = st.number_input(
            "Spleen size mm", min_value=0, max_value=500, value=None, step=1,
            key="spleen_size_num", label_visibility="collapsed",
            placeholder="SPLEEN mm")
    sp_size = str(int(sp_size_num)) if (sp_size_num and sp_size_num > 0) else ""
    sp_desc = auto_classify_spleen(sp_size_num or 0, p_age, p_sex)
    spleen_enlarged = sp_desc != "normal"

    if sp_size_num and sp_size_num > 0:
        if sp_desc == "normal":
            st.success(f"✓ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")
        elif sp_desc == "enlarged_for_age":
            st.info(f"→ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")
        else:
            st.warning(f"→ Spleen: **{SPLEEN_STATUS_LABELS[sp_desc]}**")

    if spleen_open:
        with st.container(border=True):
            st.caption("Size status derived automatically from mm value above.")
            st.markdown("**Focal lesion**")
            st.radio(
                "Focal lesion type",
                ["none", "hyperechoic_foci", "hypoechoic_foci"],
                horizontal=True, label_visibility="collapsed",
                format_func=lambda x: {"none": "None",
                                       "hyperechoic_foci": "Hyperechoic foci",
                                       "hypoechoic_foci": "Hypoechoic foci"}[x],
                key="sp_focal")
            if ss("sp_focal", "none") != "none":
                st.radio("Count", ["few", "multiple"], horizontal=True,
                         format_func=lambda x: x.title(), key="sp_focal_count")
            if spleen_enlarged:
                st.text_input(
                    "Portal vein size (mm) — replaces splenic vein line",
                    key="sp_portal_mm",
                    help="≤13 normal, >13 & <14 prominent, ≥14 dilated")
            st.checkbox("Accessory spleen present", key="sp_acc")
            if ss("sp_acc", False):
                c_a, c_b = st.columns(2)
                with c_a:
                    st.text_input("Accessory spleen size (mm)", key="sp_acc_size")
                with c_b:
                    st.radio(
                        "Location", ["hilum", "upper_pole", "lower_pole"],
                        horizontal=True,
                        format_func=lambda x: {"hilum": "Hilum",
                                               "upper_pole": "Upper pole",
                                               "lower_pole": "Lower pole"}[x],
                        key="sp_acc_loc")

    # ------- KIDNEYS -------
    kd_open = organ_button("KIDNEYS", "KIDNEYS")
    if kd_open:
        with st.container(border=True):
            c_a, c_b, c_c = st.columns([2, 2, 1])
            with c_a:
                st.radio("Cortical echogenicity",
                         ["normal", "mildly_raised"],
                         horizontal=True,
                         format_func=lambda x: {
                             "normal": "Normal",
                             "mildly_raised": "Mildly raised"}[x],
                         key="kd_cort_echo")
            with c_b:
                if ss("kd_cort_echo", "normal") == "mildly_raised":
                    st.radio("Laterality",
                             ["bilateral", "right", "left"],
                             horizontal=True,
                             format_func=lambda x: x.title(),
                             key="kd_cort_lat")
            with c_c:
                if ss("kd_cort_echo", "normal") == "mildly_raised":
                    st.checkbox("?Age related", key="kd_age_related")
            st.radio(
                "Negative renal impression line (clinical query, no finding)",
                ["no", "yes"],
                horizontal=True, key="kd_neg_renal")

            ub_status_now = ss("ub_status", "adequately_distended")

            for side in ("right", "left"):
                side_title = f"{side.title()} Kidney"
                with st.expander(side_title, expanded=False):
                    st.radio(
                        "Status",
                        ["normal", "absent_agenesis", "absent_ectopic"],
                        horizontal=True, key=f"kd_{side[0]}_status",
                        format_func=lambda x: {
                            "normal": "Present",
                            "absent_agenesis": "Absent (agenesis/hypoplasia)",
                            "absent_ectopic": "Absent (ectopic)"}[x])
                    if ss(f"kd_{side[0]}_status", "normal") != "normal":
                        st.caption("Findings skipped for absent side.")
                        continue
                    st.text_input(
                        f"{side.title()} kidney size (mm) — e.g. 100x52",
                        key=f"kd_{side[0]}_size_text")
                    st.checkbox("Renal calculi present",
                                key=f"kd_{side[0]}_has_calc")
                    calc_key = f"kd_{side[0]}_calc_count"
                    if ss(f"kd_{side[0]}_has_calc", False):
                        if calc_key not in st.session_state:
                            st.session_state[calc_key] = 1
                        calc_r_cols = st.columns([3, 1])
                        with calc_r_cols[0]:
                            st.caption(
                                f"{side.title()}: "
                                f"{st.session_state[calc_key]} calculi")
                        with calc_r_cols[1]:
                            if st.button("+ Add",
                                         key=f"kd_{side[0]}_add_calc"):
                                st.session_state[calc_key] += 1
                                st.rerun()
                        for i in range(st.session_state[calc_key]):
                            c_a, c_b, c_c = st.columns([1, 2, 1])
                            with c_a:
                                st.text_input(
                                    "Size",
                                    key=f"kd_{side[0]}_calc_size_{i}",
                                    label_visibility="collapsed",
                                    placeholder="mm")
                            with c_b:
                                st.selectbox(
                                    "Pole",
                                    ["upper", "upper_mid", "mid",
                                     "lower_mid", "lower"],
                                    key=f"kd_{side[0]}_calc_pole_{i}",
                                    label_visibility="collapsed",
                                    format_func=lambda x: _format_pole(x))
                            with c_c:
                                if st.button(
                                        "✕",
                                        key=f"kd_{side[0]}_calc_del_{i}"):
                                    st.session_state[calc_key] -= 1
                                    st.rerun()
                    else:
                        st.session_state[calc_key] = 0

                    st.checkbox("Ureteric calculus present",
                                key=f"kd_{side[0]}_has_uc")
                    if ss(f"kd_{side[0]}_has_uc", False):
                        st.radio(
                            "Ureteric calculus count",
                            ["single", "couple", "few", "multiple"],
                            horizontal=True, key=f"kd_{side[0]}_uc_count",
                            format_func=lambda x: x.title())
                        st.selectbox(
                            "Level",
                            ["renal_pelvis", "puj", "proximal_ureter",
                             "mid_ureter", "distal_ureter", "vuj"],
                            key=f"kd_{side[0]}_uc_level",
                            format_func=lambda x: {
                                "renal_pelvis": "Renal pelvis",
                                "puj": "Pelvi-ureteric junction",
                                "proximal_ureter": "Proximal ureter",
                                "mid_ureter": "Mid ureter",
                                "distal_ureter": "Distal ureter",
                                "vuj": "Vesico-ureteric junction"}[x])
                        if ss(f"kd_{side[0]}_uc_count", "single") == "couple":
                            c_a, c_b = st.columns(2)
                            with c_a:
                                st.text_input("Size 1 (mm)",
                                              key=f"kd_{side[0]}_uc_s1")
                            with c_b:
                                st.text_input("Size 2 (mm)",
                                              key=f"kd_{side[0]}_uc_s2")
                        else:
                            st.text_input("Size (mm, largest)",
                                          key=f"kd_{side[0]}_uc_size")
                        st.radio(
                            "Grade of back-pressure",
                            ["none", "no_significant", "minimal", "mild",
                             "moderate"],
                            horizontal=True, key=f"kd_{side[0]}_uc_grade",
                            format_func=lambda x: {
                                "none": "No back-pressure",
                                "no_significant": "No significant",
                                "minimal": "Minimal",
                                "mild": "Mild",
                                "moderate": "Moderate"}[x])

                    st.checkbox("Cyst present",
                                key=f"kd_{side[0]}_has_cyst")
                    if ss(f"kd_{side[0]}_has_cyst", False):
                        st.radio(
                            "Cyst type",
                            ["simple", "cortical"],
                            horizontal=True, key=f"kd_{side[0]}_cyst",
                            format_func=lambda x: x.title())
                        c1, c2 = st.columns(2)
                        with c1:
                            st.text_input("Size (mm)",
                                          key=f"kd_{side[0]}_cyst_size")
                        with c2:
                            st.selectbox(
                                "Location",
                                ["", "upper", "upper_mid", "mid",
                                 "lower_mid", "lower"],
                                key=f"kd_{side[0]}_cyst_loc",
                                format_func=lambda x: (
                                    "—" if x == "" else _format_pole(x)))
                        st.checkbox(
                            "Upgrade Bosniak category (Cat-II / Cat-IIF)",
                            key=f"kd_{side[0]}_upgrade_bosniak")
                        if ss(f"kd_{side[0]}_upgrade_bosniak", False):
                            st.radio(
                                "Bosniak category",
                                ["II", "IIF"],
                                horizontal=True,
                                key=f"kd_{side[0]}_bosniak")

                    st.checkbox("Standalone hydronephrosis",
                                key=f"kd_{side[0]}_has_hydro")
                    if ss(f"kd_{side[0]}_has_hydro", False):
                        st.radio(
                            "Grade",
                            ["minimal", "mild", "moderate"],
                            horizontal=True, key=f"kd_{side[0]}_hydro",
                            format_func=lambda x: x.title())
                        st.checkbox(
                            "No obstructive calculus upto visualized "
                            "distal ureter",
                            key=f"kd_{side[0]}_hydro_nocalc")
                        st.checkbox("?Recently passed calculus",
                                    key=f"kd_{side[0]}_hydro_rpc")
                        if ub_status_now == "empty":
                            st.checkbox(
                                f"{side.title()} distal ureter could not "
                                f"be traced (UB is empty)",
                                key=f"kd_{side[0]}_uc_not_traced")

                    st.checkbox(
                        f"Contralateral "
                        f"({'left' if side == 'right' else 'right'}) "
                        f"normal clause",
                        key=f"kd_{side[0]}_contra")

    # ------- URINARY BLADDER -------
    ub_open = organ_button("URINARY BLADDER", "URINARY BLADDER")
    if ub_open:
        with st.container(border=True):
            st.radio(
                "Status",
                ["adequately_distended", "over", "partially", "empty"],
                horizontal=True,
                format_func=lambda x: x.replace("_", " ").title(),
                key="ub_status")
            st.radio(
                "Sedimentation",
                ["none", "trace", "free_floating", "significant", "extensive"],
                horizontal=True,
                format_func=lambda x: x.replace("_", " ").title(),
                key="ub_sed")

    # ------- UTERUS / OVARIES or PROSTATE -------
    if p_sex == "F":
        u_exp_col, u_sz_col, u_et_col = st.columns(
            [4, 1, 1], vertical_alignment="bottom")
        with u_exp_col:
            uterus_open = organ_button("UTERUS", "UTERUS")
        with u_sz_col:
            st.text_input("UTERUS size (mm)", key="ut_size",
                          placeholder="UTERUS mm",
                          label_visibility="collapsed")
        with u_et_col:
            st.text_input("Endometrium (mm)", key="ut_et",
                          placeholder="Endometrium",
                          label_visibility="collapsed")
        if uterus_open:
            with st.container(border=True):
                st.radio("Status",
                         ["anteverted", "retroverted", "bulky",
                          "operated", "not_visualized"],
                         horizontal=True, key="ut_status")

        o_exp_col, o_r_col, o_l_col = st.columns(
            [4, 1, 1], vertical_alignment="bottom")
        with o_exp_col:
            ovaries_open = organ_button("OVARIES", "OVARIES")
        with o_r_col:
            st.text_input("Rt Ovary (mm)", key="ov_r_size",
                          placeholder="Rt Ovary",
                          label_visibility="collapsed")
        with o_l_col:
            st.text_input("LT Ovary (mm)", key="ov_l_size",
                          placeholder="LT Ovary",
                          label_visibility="collapsed")
        if ovaries_open:
            with st.container(border=True):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Right Ovary**")
                    st.radio("Status",
                             ["normal", "cyst", "not_visualized"],
                             horizontal=True, key="ov_r_status")
                with c2:
                    st.markdown("**Left Ovary**")
                    st.radio("Status",
                             ["normal", "cyst", "not_visualized"],
                             horizontal=True, key="ov_l_status")
    else:
        p_exp_col, p_sz_col = st.columns([5, 1], vertical_alignment="bottom")
        with p_exp_col:
            prostate_open = organ_button("PROSTATE", "PROSTATE")
        with p_sz_col:
            st.text_input("PROSTATE size (cc)", key="pr_cc",
                          placeholder="PROSTATE cc",
                          label_visibility="collapsed")
        if prostate_open:
            with st.container(border=True):
                st.radio("Status",
                         ["normal", "borderline", "bulky", "grade1"],
                         horizontal=True,
                         format_func=lambda x: {
                             "normal": "Normal",
                             "borderline": "Borderline",
                             "bulky": "Bulky",
                             "grade1": "Grade-I BPH"}[x],
                         key="pr_status")

    # ------- BOWEL -------
    bowel_open = organ_button("BOWEL / FREE FLUID", "BOWEL / FREE FLUID")
    if bowel_open:
        with st.container(border=True):
            st.radio("Free fluid",
                     ["none", "minimal", "mild", "moderate"],
                     horizontal=True, key="bw_ff")
            st.radio("Mesenteric LN", ["none", "present"],
                     horizontal=True, key="bw_ln")

    # ------- APPENDIX -------
    ap_open = organ_button("APPENDIX", "APPENDIX")
    if ap_open:
        with st.container(border=True):
            st.radio(
                "Status",
                ["not_assessed", "not_visualized", "normal", "dilated"],
                horizontal=True,
                format_func=lambda x: x.replace("_", " ").title(),
                key="ap_status")
            if ss("ap_status", "not_assessed") in ("normal", "dilated"):
                st.text_input("Diameter (mm)", key="ap_d")


# ============================================================
# Assemble data — ss() reads via the mirror pattern, so collapsing
# an organ never loses values, even across on_change reruns.
# ============================================================

data = new_report(p_sex)
data["patient"] = {"name": p_name, "age": p_age, "sex": p_sex,
                   "date": p_date, "referred_by": p_ref}

_lf = ss("liver_focal", "none")

_abscess_lesions = []
if _lf == "abscess":
    _ac = ss("abscess_count", "single")
    n_abs = 1 if _ac == "single" else int(ss("abscess_n", 2) or 2)
    for i in range(n_abs):
        _abscess_lesions.append({
            "segment": ss(f"abs_seg_{i}", "I"),
            "dim": ss(f"abs_dim_{i}", ""),
            "vol": ss(f"abs_vol_{i}", ""),
        })

data["liver"].update({
    "size_mm": liver_size,
    "size_descriptor": liver_status,
    "outline": ss("liver_outline", "normal"),
    "echotexture": ss("liver_echo", "normal"),
    "steatosis_grade": ss("liver_steatosis", None),
    "focal_lesion": _lf,
    "focal_lesion_text": ss("liver_focal_text", ""),
    "cyst_count": ss("cyst_count", "single"),
    "cyst_single_lobe": ss("cyst_single_lobe", "right"),
    "cyst_single_size_mm": ss("cyst_single_size", ""),
    "cyst_few_largest_mm": ss("cyst_few_largest", ""),
    "cyst_few_lobe": ss("cyst_few_lobe", "right"),
    "hemangioma_count": ss("hemangioma_count", "single"),
    "hemangioma_single_lobe": ss("hemangioma_single_lobe", "right"),
    "hemangioma_single_size_mm": ss("hemangioma_single_size", ""),
    "hemangioma_few_largest_mm": ss("hemangioma_few_largest", ""),
    "hemangioma_few_lobe": ss("hemangioma_few_lobe", "right"),
    "abscess_count": ss("abscess_count", "single"),
    "abscess_lesions": _abscess_lesions,
    "ihbr": ss("liver_ihbr", "normal"),
    "portal_vein": ss("liver_portal", "normal"),
    "portal_vein_mm": ss("liver_portal_mm", "") if ss("liver_portal", "normal") == "dilated" else "",
})

data["gall_bladder"].update({
    "status": ss("gb_status", "adequately_distended"),
    "wall_thickened": bool(ss("gb_wall_check", False)),
    "wall_mm": ss("gb_wall_mm", ""),
    "calculi": ss("gb_calculi", "none"),
    "calculi_count": ss("gb_calc_count", "single"),
    "calculi_size_cat": ss("gb_calc_size_cat", "small"),
    "calculi_size_mm": ss("gb_calc_size", ""),
    "calculi_neck": bool(ss("gb_calc_neck", False)),
    "calculi_neck_size_mm": ss("gb_neck_size", ""),
    "sludge": ss("gb_sludge", "none"),
    "sludge_ball": bool(ss("gb_sludge_ball", False)),
    "sludge_ball_count": ss("gb_sb_count", "single"),
    "sludge_ball_size_mm": ss("gb_sb_size", ""),
    "sludge_ball_wall": ss("gb_sb_wall", "anterior"),
    "comet_tail": bool(ss("gb_comet_tail", False)),
    "comet_tail_count": ss("gb_ct_count", "single"),
    "comet_tail_wall": ss("gb_ct_wall", "anterior"),
    "pericholecystic_fluid": bool(ss("gb_peri_fluid", False)),
})

data["cbd"].update({
    "size_mm": cbd_mm,
    "status": ss("cbd_status", "normal"),
    "calculi": bool(ss("cbd_calc", False)),
    "calculi_count": ss("cbd_calc_count", "single"),
    "calculi_size_mm": ss("cbd_calc_size", ""),
    "calculi_location": ss("cbd_calc_location", "distal"),
    "ihbr": ss("cbd_ihbr", "normal"),
})

data["pancreas"].update({
    "status": ss("pn_status", None) or "normal",
    "ee_size": ss("pn_ee_size", "normal"),
    "ee_fat_stranding": bool(ss("pn_ee_fat", False)),
    "ee_fat_location": ss("pn_ee_fat_loc", "none"),
    "ee_free_fluid": bool(ss("pn_ee_fluid", False)),
    "ee_fluid_location": ss("pn_ee_fluid_loc", "none"),
    "ac_size": ss("pn_ac_size", "normal"),
    "ac_echo": ss("pn_ac_echo", "normal"),
    "ac_echo_location": ss("pn_ac_echo_loc", "none"),
    "ac_margins": ss("pn_ac_margins", "normal"),
    "wp_type": ss("pn_wp_type", "won"),
    "wp_dims": ss("pn_wp_dims", ""),
    "wp_vol": ss("pn_wp_vol", ""),
    "wp_location": ss("pn_wp_location", "lesser_sac"),
    "ch_foci": bool(ss("pn_ch_foci", False)),
    "ch_mpd": bool(ss("pn_ch_mpd", False)),
    "ch_mpd_size": ss("pn_ch_mpd_size", ""),
    "ch_fat": bool(ss("pn_ch_fat", False)),
})

data["spleen"].update({
    "size_mm": sp_size,
    "size_descriptor": sp_desc,
    "focal_lesion": ss("sp_focal", "none"),
    "focal_count": ss("sp_focal_count", "few"),
    "portal_vein_mm": ss("sp_portal_mm", "") if spleen_enlarged else "",
    "accessory_spleen": bool(ss("sp_acc", False)),
    "accessory_size_mm": ss("sp_acc_size", ""),
    "accessory_location": ss("sp_acc_loc", "hilum"),
})

for side in ("right", "left"):
    sd = side[0]
    calcs = []
    if ss(f"kd_{sd}_has_calc", False):
        for i in range(ss(f"kd_{sd}_calc_count", 0)):
            sz = ss(f"kd_{sd}_calc_size_{i}", "")
            pole = ss(f"kd_{sd}_calc_pole_{i}", "mid")
            if sz:
                calcs.append({"size_mm": sz, "pole": pole})

    status = ss(f"kd_{sd}_status", "normal")

    if status == "normal":
        has_uc = ss(f"kd_{sd}_has_uc", False)
        if has_uc:
            uc_count = ss(f"kd_{sd}_uc_count", "single")
            if uc_count == "couple":
                uc_sizes = [
                    ss(f"kd_{sd}_uc_s1", ""),
                    ss(f"kd_{sd}_uc_s2", ""),
                ]
            else:
                uc_sizes = [ss(f"kd_{sd}_uc_size", "")]
            uc_level = ss(f"kd_{sd}_uc_level", "distal_ureter")
            uc_grade = ss(f"kd_{sd}_uc_grade", "none")
        else:
            uc_count = "none"
            uc_sizes = []
            uc_level = "distal_ureter"
            uc_grade = "none"

        if ss(f"kd_{sd}_has_cyst", False):
            cyst = ss(f"kd_{sd}_cyst", "simple")
            cyst_size = ss(f"kd_{sd}_cyst_size", "")
            cyst_loc = ss(f"kd_{sd}_cyst_loc", "")
            if ss(f"kd_{sd}_upgrade_bosniak", False):
                bosniak = ss(f"kd_{sd}_bosniak", "II")
            else:
                bosniak = "I"
        else:
            cyst = "none"
            cyst_size = ""
            cyst_loc = ""
            bosniak = ""

        if ss(f"kd_{sd}_has_hydro", False):
            hydro = ss(f"kd_{sd}_hydro", "mild")
            hydro_nocalc = ss(f"kd_{sd}_hydro_nocalc", False)
            hydro_rpc = ss(f"kd_{sd}_hydro_rpc", False)
            uc_not_traced = ss(f"kd_{sd}_uc_not_traced", False)
        else:
            hydro = "none"
            hydro_nocalc = False
            hydro_rpc = False
            uc_not_traced = False

        contra = ss(f"kd_{sd}_contra", False)

        data["kidneys"][side].update({
            "status": "normal",
            "size_text": ss(f"kd_{sd}_size_text", ""),
            "calculi": calcs,
            "ureter_calculus": {
                "count": uc_count,
                "sizes": [s for s in uc_sizes if s],
                "level": uc_level,
                "grade": uc_grade,
            },
            "ureter_not_traced": uc_not_traced,
            "cyst": cyst,
            "cyst_size_mm": cyst_size,
            "cyst_location": cyst_loc,
            "bosniak": bosniak,
            "hydronephrosis": hydro,
            "hydronephrosis_no_obstructive_calculus": hydro_nocalc,
            "recently_passed_calculus_suspected": hydro_rpc,
            "contralateral_normal_clause": contra,
        })
    else:
        data["kidneys"][side].update({
            "status": status,
            "size_text": "",
            "calculi": [],
            "ureter_calculus": _new_ureter_calculus(),
            "ureter_not_traced": False,
            "cyst": "none",
            "cyst_size_mm": "",
            "cyst_location": "",
            "bosniak": "",
            "hydronephrosis": "none",
            "hydronephrosis_no_obstructive_calculus": False,
            "recently_passed_calculus_suspected": False,
            "contralateral_normal_clause": False,
        })

data["kidneys"]["cortical_echogenicity"] = ss("kd_cort_echo", "normal")
data["kidneys"]["cortical_echogenicity_laterality"] = ss("kd_cort_lat", "bilateral")
data["kidneys"]["age_related_echogenicity"] = bool(ss("kd_age_related", False))
data["kidneys"]["negative_renal_line"] = (ss("kd_neg_renal", "no") == "yes")

_r_uc = data["kidneys"]["right"]["ureter_calculus"]
_l_uc = data["kidneys"]["left"]["ureter_calculus"]
data["kidneys"]["bilateral_ureter_calculi"] = (
    data["kidneys"]["right"]["status"] == "normal"
    and data["kidneys"]["left"]["status"] == "normal"
    and _r_uc["count"] != "none"
    and _l_uc["count"] != "none"
)

_force_ub_empty = (data["kidneys"]["right"]["ureter_not_traced"]
                   or data["kidneys"]["left"]["ureter_not_traced"])
data["urinary_bladder"].update({
    "status": ss("ub_status", "adequately_distended"),
    "sedimentation": ss("ub_sed", "none"),
    "force_empty_suboptimal": _force_ub_empty,
})

if p_sex == "F":
    data["uterus"].update({
        "status": ss("ut_status", "anteverted"),
        "size": ss("ut_size", ""),
        "endometrial_thickness_mm": ss("ut_et", ""),
    })
    data["ovaries"].update({
        "right_status": ss("ov_r_status", "normal"),
        "right_size": ss("ov_r_size", ""),
        "left_status": ss("ov_l_status", "normal"),
        "left_size": ss("ov_l_size", ""),
    })
else:
    data["prostate"].update({
        "status": ss("pr_status", "normal"),
        "size_cc": ss("pr_cc", ""),
    })

data["bowel"].update({
    "free_fluid": ss("bw_ff", "none"),
    "mesenteric_ln": ss("bw_ln", "none"),
})
data["appendix"].update({
    "status": ss("ap_status", "not_assessed"),
    "diameter_mm": ss("ap_d", ""),
})


# ============================================================
# PREVIEW / ADDENDUM / IMPRESSION
# ============================================================

with col_prev:
    st.subheader("📄 Live Preview")
    preview_placeholder = st.empty()

with col_prev:
    st.subheader("✏️ Additional Body Findings (manual)")
    st.caption("Optional free text. Rendered in bold italic just before "
               "IMPRESSION in the preview and DOCX. First letter of each "
               "sentence is auto-capitalized.")

    def _cap_addendum_cb():
        cur = st.session_state.get("additional_body_findings_box", "")
        st.session_state["additional_body_findings_box"] = capitalize_sentences(cur)

    addendum_text = st.text_area(
        "Additional body findings",
        height=120,
        label_visibility="collapsed",
        key="additional_body_findings_box",
        placeholder="e.g. A small umbilical hernia is noted in the anterior "
                    "abdominal wall.",
        on_change=_cap_addendum_cb)

data["additional_body_findings"] = addendum_text or ""


def render_preview_findings(data):
    p = data["patient"]
    out = []
    out.append("         ULTRASOUND WHOLE ABDOMEN")
    out.append("")
    sex, age = p["sex"], p["age"]
    age_years = parse_age(age)
    is_ped = age_years is not None and age_years < 18
    p_status = data["pancreas"].get("status", "normal")
    spleen_enlarged = data["spleen"]["size_descriptor"] != "normal"
    ub_status = data["urinary_bladder"].get("status", "adequately_distended")
    secs = [liver_sentence(data["liver"], sex, age,
                            spleen_enlarged=spleen_enlarged),
            gall_bladder_sentence(data["gall_bladder"]),
            cbd_sentence(data["cbd"]),
            pancreas_sentence(data["pancreas"]),
            spleen_sentence(data["spleen"]),
            kidneys_sentence(data["kidneys"], ub_status=ub_status),
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
    addendum = (data.get("additional_body_findings") or "").strip()
    if addendum:
        out.append(addendum)
        out.append("")
    out.append("IMPRESSION:")
    for line in data["impression"]["lines"]:
        out.append(f"  - {line}")
    return "\n".join(out)


auto_impression = generate_impression(data, p_sex, p_age)
data["impression"]["lines"] = auto_impression

preview_text = render_preview_findings(data)

with preview_placeholder:
    _safe = html.escape(preview_text).replace("\n", "<br>")
    st.markdown(
        f'<div class="preview-box">{_safe}</div>',
        unsafe_allow_html=True,
    )


with col_prev:
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
        final_data["additional_body_findings"] = addendum_text or ""
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
                final_data["additional_body_findings"] = addendum_text or ""
                save_report(final_data)
                if p_ref.strip():
                    add_referrer(p_ref)
                st.success("Saved to local database.")
