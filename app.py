"""
Radiology Report Generator — USG Whole Abdomen
v3.1.1-stable

v3.1.1:
  - DOCX spacing fixed (Issue #10): explicit empty paragraphs at 5.5pt (half line)
    and 11pt (one line) replace space_after. Layout:
        * patient table -> 0.5 line -> title
        * title -> 1.5 lines -> first organ
        * 1 line between each non-empty organ section
        * last organ -> 1.5 lines -> IMPRESSION heading
        * IMPRESSION heading -> 0.5 line -> first bullet
        * last bullet -> 0.5 line -> disclaimer box
    Empty paragraphs carry a run at the required size so they occupy real
    vertical space; all paragraph space_after/space_before remain Pt(0).

v3.1.0: (see git history)
v3.0.0: (see git history)

Frozen rules:
- Impression text ALL CAPS except:
    * "Adv- ... Correlation."                -> bold + italic, mixed case
    * anything inside parentheses "( ... )"  -> bold + italic, mixed case
    * the token "vs"                         -> bold, lowercase
    * EXCEPTION: "BOSNIAK CAT-" + Roman numeral stays ALL CAPS even
      inside parentheses.
- Body findings bold, italic, mixed case, unless an explicit override is given.
- Body text uses lowercase "mm"; impressions use uppercase "MM".
- No space after "?" anywhere.
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
FONT_SIZE_HALF = 5.5
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

def _mirror(widget_key):
    mirror_key = f"_mirror_{widget_key}"
    if widget_key in st.session_state:
        st.session_state[mirror_key] = st.session_state[widget_key]
        return st.session_state[widget_key]
    return st.session_state.get(mirror_key)


def ss(widget_key, default):
    v = _mirror(widget_key)
    return v if v is not None else default


def _q(text):
    if not isinstance(text, str):
        return text
    return re.sub(r"\?\s+", "?", text)


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


def parse_kidney_length_mm(size_text):
    if not size_text:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", str(size_text))
    return float(m.group(1)) if m else None


def normalize_kidney_size_text(size_text):
    if not size_text:
        return ""
    s = str(size_text).strip()
    if not s:
        return ""
    s = re.sub(r"\s*[xX]\s*", "x", s)
    s = re.sub(r"\s*[mM]{2}\s*$", "", s)
    return f"{s}mm"


def _first_number(text):
    m = re.search(r"(\d+(?:\.\d+)?)", str(text or ""))
    if not m:
        return None
    try:
        return float(m.group(1))
    except (ValueError, TypeError):
        return None


def _lower_mm_in_text(t):
    if not isinstance(t, str):
        return t
    return re.sub(r"\bMM\b", "mm", t)


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
    "none": None, "no_significant": "NO SIGNIFICANT",
    "minimal": "MINIMAL", "mild": "MILD", "moderate": "MODERATE",
}

URETER_GRADES_SENTENCE = {
    "no_significant": "no significant", "minimal": "minimal",
    "mild": "mild", "moderate": "moderate",
}

POLE_LABELS = {
    "upper": "upper-pole", "upper_mid": "upper-mid pole",
    "mid": "mid pole", "lower_mid": "lower-mid pole", "lower": "lower-pole",
}

SMALL_KIDNEY_MM = 82
SMALL_KIDNEY_MIN_AGE = 15

ECHO_GRADE_WORD = {
    "mildly_raised": "mildly raised",
    "moderately_raised": "moderately raised",
    "significantly_raised": "significantly raised",
}
ECHO_GRADE_WORD_UP = {
    "mildly_raised": "MILDLY RAISED",
    "moderately_raised": "MODERATELY RAISED",
    "significantly_raised": "SIGNIFICANTLY RAISED",
}
CMD_WORD = {"preserved": "preserved", "hazy": "hazy", "lost": "lost"}
CMD_WORD_UP = {"preserved": "PRESERVED", "hazy": "HAZY", "lost": "LOST"}


# ============================================================
# PROSTATE CONSTANTS
# ============================================================

PROSTATE_GRADE_BANDS = [
    (30, 36, "borderline bulky", "BORDERLINE PROSTATOMEGALY"),
    (36, 46, "mildly bulky", "GRADE-I PROSTATOMEGALY"),
    (46, 51, "mild to moderately bulky", "GRADE-I/II PROSTATOMEGALY"),
    (51, 71, "moderately bulky", "GRADE-II PROSTATOMEGALY"),
    (71, float("inf"), "moderate to grossly bulky", "GRADE-III/IV PROSTATOMEGALY"),
]


def prostate_band_from_volume(cc):
    try:
        v = float(cc)
    except (ValueError, TypeError):
        return None
    if v < 30:
        return None
    for lo, hi, body_desc, imp_label in PROSTATE_GRADE_BANDS:
        if lo <= v < hi:
            return body_desc, imp_label
    return None


# ============================================================
# UTERUS CONSTANTS
# ============================================================

FIBROID_TYPE_LABELS = {
    "intramural": ("an intramural fibroid", "FIGO-4"),
    "intramural_subserosal": ("an intramural>subserosal fibroid", "FIGO-4/5"),
    "intramural_submucosal": ("an intramural>submucosal fibroid", "FIGO-4/3"),
    "submucosal_intramural": ("a submucosal>intramural fibroid", "FIGO-2/3"),
    "subserosal_intramural":
        ("a partially extrusive subserosal>intramural fibroid", "FIGO-5/4"),
}
FIBROID_TYPE_LABELS_UP = {
    "intramural": ("AN INTRAMURAL FIBROID", "FIGO-4"),
    "intramural_subserosal": ("AN INTRAMURAL>SUBSEROSAL FIBROID", "FIGO-4/5"),
    "intramural_submucosal": ("AN INTRAMURAL>SUBMUCOSAL FIBROID", "FIGO-4/3"),
    "submucosal_intramural": ("A SUBMUCOSAL>INTRAMURAL FIBROID", "FIGO-2/3"),
    "subserosal_intramural":
        ("A PARTIALLY EXTRUSIVE SUBSEROSAL>INTRAMURAL FIBROID", "FIGO-5/4"),
}
PARTIALLY_EXTRUSIVE_TYPES = {"intramural_subserosal", "subserosal_intramural"}


def classify_uterus_size(length_mm):
    if length_mm is None:
        return "normal"
    if length_mm < 90:
        return "normal"
    if length_mm <= 100:
        return "bulky"
    if length_mm < 120:
        return "moderately_bulky"
    return "grossly_bulky"


UTERUS_SIZE_BODY = {
    "normal": "normal in size",
    "bulky": "bulky in size",
    "moderately_bulky": "moderately bulky in size",
    "grossly_bulky": "grossly bulky in size",
}
UTERUS_SIZE_IMPRESSION = {
    "normal": "",
    "bulky": "BULKY UTERUS",
    "moderately_bulky": "MODERATELY BULKY UTERUS",
    "grossly_bulky": "GROSSLY BULKY UTERUS",
}
UTERUS_SIZE_IMPRESSION_ADJ = {
    "normal": "",
    "bulky": "BULKY",
    "moderately_bulky": "MODERATELY BULKY",
    "grossly_bulky": "GROSSLY BULKY",
}

ENDOMETRIAL_COLLECTION_LABELS = {
    "none": "No collection seen in the endometrial cavity.",
    "mild": "Mild collection seen in the endometrial cavity.",
    "moderate": "Moderate collection seen in the endometrial cavity.",
}
ENDOMETRIAL_COLLECTION_HET = {
    "mild": "Mild heterogenous collection seen in the endometrial cavity.",
    "moderate": "Moderate heterogenous collection seen in the endometrial cavity.",
}
ENDOMETRIAL_COLLECTION_IMPRESSION = {
    "none": None,
    "mild": "MILD COLLECTION IN THE ENDOMETRIAL CANAL.",
    "moderate": "MODERATE COLLECTION IN THE ENDOMETRIAL CANAL.",
}
ENDOMETRIAL_COLLECTION_HET_IMPRESSION = {
    "mild": "MILD HETEROGENOUS COLLECTION IN THE ENDOMETRIAL CANAL.",
    "moderate": "MODERATE HETEROGENOUS COLLECTION IN THE ENDOMETRIAL CANAL.",
}


def classify_endometrium(mm):
    if mm is None:
        return None
    if mm < 2:
        return "thinned_out"
    if mm <= 15:
        return None
    if mm < 20:
        return "thickened"
    return "significantly_thickened"


ADENOMYOSIS_FEATURES = [
    "globular_shape",
    "asymmetric_posterior",
    "venetian_blind_sign",
    "interface_indistinct",
    "interface_barely_perceptible_fundus",
    "interface_lost",
    "subendometrial_cysts",
]


# ============================================================
# OVARY CONSTANTS
# ============================================================

OVARY_LARGE_MM = 50
OVARY_BULKY_CC = 10.0

OVARY_CYST_TYPES = {
    "simple_cyst": "simple cyst",
    "hemorrhagic_cyst": "hemorrhagic cyst",
    "hemorrhagic_follicle": "hemorrhagic follicle",
    "follicular_cyst": "follicular cyst",
}
OVARY_CYST_TYPES_UP = {
    "simple_cyst": "SIMPLE CYST",
    "hemorrhagic_cyst": "HEMORRHAGIC CYST",
    "hemorrhagic_follicle": "HEMORRHAGIC FOLLICLE",
    "follicular_cyst": "FOLLICULAR CYST",
}
OVARY_CYST_TYPES_UP_PLURAL = {
    "simple_cyst": "SIMPLE CYSTS",
    "hemorrhagic_cyst": "HEMORRHAGIC CYSTS",
    "hemorrhagic_follicle": "HEMORRHAGIC FOLLICLES",
    "follicular_cyst": "FOLLICULAR CYSTS",
}

PCOS_OUTCOME_CONCLUSIVE = "conclusive"
PCOS_OUTCOME_PARTIAL = "partial"
PCOS_OUTCOME_NOT_CONCLUSIVE = "not_conclusive"

PCOS_OUTCOME_PHRASES = {
    PCOS_OUTCOME_CONCLUSIVE: (
        "NEEDS CLINICO-LAB CORRELATION TO RULE OUT POLYCYSTIC OVARIAN "
        "SPECTRUM CHANGES (PCOS). Adv - LH/FSH & AMH Correlation."
    ),
    PCOS_OUTCOME_PARTIAL: (
        "PARTIALLY FITS ON THE POLYCYSTIC OVARIAN SPECTRUM. "
        "Adv- LH/FSH & AMH Correlation."
    ),
    PCOS_OUTCOME_NOT_CONCLUSIVE: (
        "FEATURES DO NOT APPEAR CONCLUSIVE FOR POLYCYSTIC OVARIAN SPECTRUM "
        "CHANGES (PCOS). Adv - LH/FSH & AMH Correlation."
    ),
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

def _new_ureter_calculus():
    return {"count": "none", "sizes": [], "level": "distal_ureter", "grade": "none"}


def _new_kidney_side():
    return {
        "status": "normal", "size_text": "", "calculi": [],
        "ureter_calculus": _new_ureter_calculus(), "ureter_not_traced": False,
        "cyst_type": "none", "cyst_count": "single", "cyst_size_mm": "",
        "cyst_location": "", "cyst_septa": "none", "cyst_calc": "none",
        "cyst_bosniak": "", "hydronephrosis": "none",
        "hydronephrosis_no_obstructive_calculus": False,
        "recently_passed_calculus_suspected": False,
        "nephrocalcinosis": "none", "contralateral_normal_clause": False,
        "is_small": False,
    }


def _new_urinary_bladder():
    return {
        "status": None, "catheterized": False, "pre_void_cc": "",
        "post_void_cc": "", "mass_calculus": False,
        "sedimentation_free_floating": False,
        "sedimentation_settled_debris": False,
        "sedimentation_trace": False,
        "sedimentation_extensive": False,
        "uti_suspected": False, "wall_thickening": False,
        "wall_irregular": False, "wall_mm": "",
        "force_empty_suboptimal": False,
    }


def _new_prostate():
    return {
        "size_cc": "", "median_lobe": False, "median_lobe_size_mm": "",
        "median_lobe_boo": False, "cyst": False, "cyst_size": "",
        "cyst_left_hemi": False, "cyst_right_hemi": False, "not_visualized": False,
        "partial_visualization": False,
    }


def _new_fibroid_lesion():
    return {"location": "fundal", "size": ""}


def _new_uterus():
    return {
        "status": "anteverted",
        "size": "",
        "retroflexed": False,
        "gravid": False,
        "gravid_type": "CRL",
        "gravid_length": "",
        "gravid_weeks": "",
        "gravid_days": "",
        "low_lying": False,
        "cervix_visualized": "partially",
        "myometrium": "homogenous",
        "prominent_vascular_channels": False,
        "endometrial_thickness_mm": "",
        "endometrial_collection": "none",
        "endometrial_collection_het": False,
        "rpoc": False, "rpoc_size": "", "rpoc_echo": "hypo",
        "cervix_elongated": False, "cervix_bulky": False,
        "cervix_bulky_size_mm": "",
        "nabothian": "none", "cervix_collection": "none",
        "fibroid": False, "fibroid_count": "single",
        "fibroid_types": [], "fibroid_lesions": [], "fibroid_figo_manual": "",
        "adenomyosis": False, "adenomyosis_features": [],
        "adenomyosis_confidence": "auto",
        "adenomyomas": False, "adenomyomas_count": "couple",
        "adenomyomas_echo": "hyperechoic",
        "adenomyomas_location": "posterior", "adenomyomas_size": "",
        "adenomyomas_avascular": True,
        "neg_rpoc": False, "neg_sol": False,
    }


def _new_ovary_side():
    return {
        "status": "normal",
        "size_text": "",
        "volume_cc": "",
        "findings": [],
    }


def _new_ovaries():
    return {
        "right": _new_ovary_side(),
        "left": _new_ovary_side(),
        "pcos": False,
        "pcos_feature2": False,
        "pcos_feature3": False,
        "pcos_feature4": False,
        "pcos_feature4_variable": False,
        "pcos_feature4_peripheral": False,
        "pcos_feature4_random": False,
        "pcos_outcome": "",
        "afc": "normal",
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
            "sludge": "none", "sludge_ball": False, "sludge_ball_count": "single",
            "sludge_ball_size_mm": "", "sludge_ball_wall": "anterior",
            "comet_tail": False, "comet_tail_count": "single",
            "comet_tail_wall": "anterior", "pericholecystic_fluid": False,
        },
        "cbd": {"size_mm": "", "status": "normal", "calculi": False,
                "calculi_count": "single", "calculi_size_mm": "",
                "calculi_location": "distal", "ihbr": "normal"},
        "pancreas": {
            "status": "normal", "ee_size": "normal",
            "ee_fat_stranding": False, "ee_fat_location": "none",
            "ee_free_fluid": False, "ee_fluid_location": "none",
            "ac_size": "normal", "ac_echo": "normal",
            "ac_echo_location": "none", "ac_margins": "normal",
            "wp_type": "won", "wp_dims": "", "wp_vol": "",
            "wp_location": "lesser_sac",
            "ch_foci": False, "ch_mpd": False, "ch_mpd_size": "", "ch_fat": False,
            "peripancreatic_ln": False, "periportal_ln": False,
        },
        "spleen": {
            "size_mm": "", "size_descriptor": "normal",
            "focal_lesion": "none", "focal_count": "few",
            "portal_vein_mm": "",
            "accessory_spleen": False, "accessory_size_mm": "",
            "accessory_location": "hilum",
        },
        "kidneys": {
            "right": _new_kidney_side(), "left": _new_kidney_side(),
            "cortical_echogenicity": "normal", "cortical_cmd": "preserved",
            "cortical_echogenicity_laterality": "bilateral",
            "age_related_echogenicity": False, "negative_renal_line": False,
            "bilateral_ureter_calculi": False,
        },
        "urinary_bladder": _new_urinary_bladder(),
        "uterus": _new_uterus(),
        "ovaries": _new_ovaries(),
        "prostate": _new_prostate(),
        "bowel": {"wall_thickening": False, "free_fluid": "none",
                  "mesenteric_ln": "none", "ln_size_category": "sad_lt_7",
                  "ln_location": "bilateral", "ln_largest": "",
                  "ln_character": "discrete", "pleural_effusion": "none"},
        "appendix": {"status": "not_assessed", "diameter_mm": ""},
        "additional_body_findings": "",
        "impression": {"lines": []},
    }


# ============================================================
# SEGMENTS
# ============================================================

def seg(text, bold=False, underline=False, italic=None):
    if isinstance(text, str):
        text = _lower_mm_in_text(text)
        text = _q(text)
    return (text, bold, underline, italic)


# ============================================================
# LIVER
# ============================================================

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
            text = (f"A hyperechoic small SOL({size}mm) seen in the {lobe} "
                    f"hepatic lobe of liver.")
            return [seg(" "), seg(text, True)]
        else:
            size = d.get("hemangioma_few_largest_mm", "")
            lobe = d.get("hemangioma_few_lobe", "right")
            text = (f"Few hyperechoic small SOLs seen in the liver, largest of "
                    f"these measuring {size}mm in {lobe} lobe.")
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


def liver_sentence(d, sex, age, spleen_enlarged=False, cbd_ihbr="normal"):
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    outline = d.get("outline", "normal")
    echo = d["echotexture"]
    grade = d.get("steatosis_grade")
    s = [seg("LIVER", True, True)]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}mm)"))
    elif desc == "borderline":
        s += [seg(" is "), seg("borderline enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "mild":
        s += [seg(" is "), seg("mildly enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("moderately enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "gross":
        s += [seg(" is "), seg("grossly enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("enlarged for age in size", True), seg(f" ({size}mm)")]

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
        elif echo == "raised":
            s.append(seg(" raised echotexture", True))
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
        elif echo == "raised":
            s.append(seg("raised echotexture", True))
            s.append(seg("."))

    s.extend(liver_focal_sentence(d))

    if cbd_ihbr == "normal":
        if d["ihbr"] == "normal":
            s.append(seg(" Intra hepatic biliary radicals are normal."))
        else:
            s += [seg(" Intra hepatic biliary radicals are "),
                  seg("dilated", True), seg(".")]

    if not spleen_enlarged:
        if d["portal_vein"] == "normal":
            s.append(seg(" Portal vein is normal in course and caliber."))
        else:
            s += [seg(" Portal vein is "), seg("dilated", True)]
            if d["portal_vein_mm"]:
                s.append(seg(f" ({d['portal_vein_mm']}mm)"))
            s.append(seg("."))
    return s


# ============================================================
# GALL BLADDER
# ============================================================

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

    if status != "contracted":
        if d.get("wall_thickened") and d.get("wall_mm"):
            try:
                w = float(d["wall_mm"])
            except ValueError:
                w = 0
            if w <= 8:
                s += [seg(" "), seg(f"wall is mildly thickened upto {d['wall_mm']}mm.", True)]
            else:
                s += [seg(" "), seg(f"wall is significantly thickened upto {d['wall_mm']}mm.", True)]
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
                s += [seg(" "), seg(f"A calculus measuring {size}mm, seen impacted "
                                     f"at the GB neck.", True)]
            else:
                s += [seg(" "), seg(f"A calculus measuring {size}mm is seen in the "
                                     f"GB lumen.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            if neck:
                s += [seg(" "), seg(f"{word} {size_cat} calculi seen in the GB lumen, "
                                     f"with a calculus measuring {neck_size}mm "
                                     f"impacted at the GB neck.", True)]
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
            s += [seg(" "), seg(f"A comet tail artifact seen arising from the {wall} "
                                 f"GB wall.", True)]
        else:
            word = "Few" if count == "few" else "Multiple"
            s += [seg(" "), seg(f"{word} comet tail artifacts seen arising from the "
                                 f"{wall} GB wall.", True)]

    if d.get("pericholecystic_fluid"):
        s += [seg(" "), seg("Thin rim of pericholecystic fluid seen.", True)]
    return s


# ============================================================
# CBD
# ============================================================

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
            s.append(seg(f" is normal in caliber({size}mm)"))
        else:
            s.append(seg(" is normal in caliber"))
    elif status == "proximal":
        s.append(seg(" is "))
        s.append(seg("proximally dilated in caliber", True))
        if size:
            s.append(seg(f" upto {size}mm", True))
    elif status == "dilated":
        s.append(seg(" is "))
        s.append(seg("dilated in caliber", True))
        if size:
            s.append(seg(f" upto {size}mm", True))

    if calculi and calculi_size:
        if calculi_count == "single":
            s.append(seg(f", with suggestion of a calculus measuring {calculi_size}mm "
                         f"in the {loc_word}", True))
        else:
            s.append(seg(f", with suggestion of few calculi in the {loc_word}, "
                         f"largest of these measuring {calculi_size}mm", True))
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
            s += [seg(" "), seg("Intra hepatic biliary radicals are proximally "
                                 "dilated.", True)]
        elif ihbr == "dilated":
            s += [seg(" "), seg("Intra hepatic biliary radicals are dilated.", True)]
    return s


# ============================================================
# PANCREAS
# ============================================================

def pancreas_sentence(d):
    s = [seg("PANCREAS", True, True)]
    status = d.get("status", "normal")
    if status == "normal":
        s.append(seg(" is normal in size, outline and echotexture. No focal lesion "
                     "is seen. No evidence of calcification is seen."))
        return s

    if status == "early_evolving":
        ee_size = d.get("ee_size", "normal")
        fat = d.get("ee_fat_stranding", False)
        fluid = d.get("ee_free_fluid", False)
        fat_loc = d.get("ee_fat_location", "none")
        fluid_loc = d.get("ee_fluid_location", "none")
        if ee_size == "mildly_bulky":
            s += [seg(" is "), seg("mildly bulky in size", True),
                  seg(" with normal echotexture.")]
        else:
            s.append(seg(" is normal in size, outline and echotexture. No focal "
                         "lesion is seen. No evidence of calcification is seen."))
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
                s += [seg(" "), seg("Mild peri-splenic fat stranding and free "
                                     "fluid seen.", True)]
            elif loc:
                s += [seg(" "), seg(f"Mild peri-pancreatic fat stranding and free "
                                     f"fluid seen {loc}.", True)]
            else:
                s += [seg(" "), seg("Mild peri-pancreatic fat stranding and free "
                                     "fluid seen.", True)]
        elif fat:
            loc = loc_phrase_map.get(fat_loc, "")
            if fat_loc == "perisplenic":
                s += [seg(" "), seg("Mild peri-splenic fat stranding is seen.", True)]
            elif loc:
                s += [seg(" "), seg(f"Mild peri-pancreatic fat stranding is seen "
                                     f"{loc}.", True)]
            else:
                s += [seg(" "), seg("Mild peri-pancreatic fat stranding is seen.", True)]
        elif fluid:
            loc = loc_phrase_map.get(fluid_loc, "")
            if fluid_loc == "perisplenic":
                s += [seg(" "), seg("Mild peri-splenic free fluid is seen.", True)]
            elif loc:
                s += [seg(" "), seg(f"Mild peri-pancreatic free fluid is seen {loc}.",
                                     True)]
            else:
                s += [seg(" "), seg("Mild peri-pancreatic free fluid is seen.", True)]
        return s

    if status == "acute":
        ac_size = d.get("ac_size", "normal")
        ac_echo = d.get("ac_echo", "normal")
        ac_echo_loc = d.get("ac_echo_location", "none")
        ac_margins = d.get("ac_margins", "normal")
        if ac_size == "bulky":
            s += [seg(" is "), seg("bulky in size", True)]
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
                s.append(seg(" with normal echotexture and irregular/fuzzy margins",
                             True))
                s.append(seg("."))
        else:
            if ac_echo == "normal" and ac_margins == "normal":
                s.append(seg(" is normal in size, outline and echotexture. No focal "
                             "lesion is seen."))
            elif ac_echo == "hypoechoic":
                if ac_echo_loc == "head_neck":
                    s += [seg(" appears "),
                          seg("hypoechoic and heterogeneous in the head and neck "
                              "region", True)]
                elif ac_echo_loc == "body":
                    s += [seg(" appears "),
                          seg("hypoechoic and heterogeneous in the body region", True)]
                else:
                    s += [seg(" echotexture appears "),
                          seg("hypoechoic and heterogeneous", True)]
                if ac_margins == "irregular":
                    s.append(seg(" with irregular/fuzzy margins", True))
                s.append(seg("."))
            elif ac_margins == "irregular":
                s += [seg(" shows "), seg("irregular/fuzzy margins", True), seg(".")]
        s += [seg(" "), seg("Mild to moderate peri-pancreatic fat stranding and mild "
                             "free fluid is seen.", True)]
        return s

    if status == "won_pseudocyst":
        wp_type = d.get("wp_type", "won")
        dims = d.get("wp_dims", "")
        vol = d.get("wp_vol", "")
        loc = d.get("wp_location", "lesser_sac")
        loc_phrase = ("in the lesser sac" if loc == "lesser_sac"
                      else "overlying the body of the pancreas")
        dims_vol = f"({dims}mm; Vol= {vol}cc)"
        if wp_type == "won":
            s += [seg(" is "),
                  seg("irregularly defined and obscured by a thick-walled collection",
                      True),
                  seg(f"{dims_vol} {loc_phrase} with debris within and significant "
                      f"peripancreatic fat stranding and mild free fluid.")]
        elif wp_type == "pseudocyst":
            s += [seg(" appears "),
                  seg("hypotrophied with irregular margins with a loculated collection",
                      True),
                  seg(f"{dims_vol} with clear contents {loc_phrase} and mild "
                      f"peripancreatic fat stranding with mild free fluid.")]
        else:
            s += [seg(" appears "),
                  seg("hypotrophied with irregular margins & a loculated mildly thick "
                      "walled collection", True),
                  seg(f"{dims_vol} with internal echoes seen {loc_phrase} and mild "
                      f"peripancreatic fat stranding with mild free fluid.")]
        return s

    if status == "chronic":
        foci = d.get("ch_foci", False)
        mpd = d.get("ch_mpd", False)
        mpd_size = d.get("ch_mpd_size", "")
        fat = d.get("ch_fat", False)
        s += [seg(" is "), seg("poorly defined and appears hypotrophied", True)]
        if foci and mpd and fat:
            s += [seg(", studded with foci of calcification and dilated MPD"
                      f"(upto {mpd_size} mm). "),
                  seg("Associated mild fat stranding is also present.", True)]
        elif foci and mpd:
            s += [seg(", studded with foci of calcification & dilated MPD"
                      f"(upto {mpd_size} mm). ", True),
                  seg("However, no associated fat stranding appreciated.")]
        elif foci and fat:
            s += [seg(", studded with foci of calcification & reveals associated "
                      "mild fat stranding. ", True),
                  seg("MPD is, however, not dilated.")]
        elif mpd and fat:
            s.append(seg(f", with dilated MPD(upto {mpd_size} mm) & reveals "
                         f"associated mild fat stranding.", True))
        elif foci:
            s += [seg(", studded with foci of calcification. ", True),
                  seg("MPD is, however, not dilated.")]
        elif mpd:
            s += [seg(f", with dilated MPD(upto {mpd_size} mm). ", True),
                  seg("However, no foci of calcification or associated fat stranding "
                      "appreciated.")]
        elif fat:
            s += [seg(", associated with mild fat stranding. ", True),
                  seg("However, MPD is not dilated & no foci of calcification "
                      "appreciated.")]
        else:
            s.append(seg("."))
        return s
    return s


# ============================================================
# SPLEEN
# ============================================================

def spleen_sentence(d):
    s = [seg("SPLEEN", True, True)]
    size = d["size_mm"] or "___"
    desc = d["size_descriptor"]
    if desc == "normal":
        s.append(seg(f" is normal in size ({size}mm)"))
    elif desc == "borderline":
        s += [seg(" is "), seg("borderline enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "mild":
        s += [seg(" is "), seg("mildly enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "mild_to_moderate":
        s += [seg(" is "), seg("mild to moderately enlarged in size", True),
              seg(f" ({size}mm)")]
    elif desc == "moderate":
        s += [seg(" is "), seg("moderately enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "moderate_to_gross":
        s += [seg(" is "), seg("moderately to grossly enlarged in size", True),
              seg(f" ({size}mm)")]
    elif desc == "gross":
        s += [seg(" is "), seg("grossly enlarged in size", True), seg(f" ({size}mm)")]
    elif desc == "enlarged_for_age":
        s += [seg(" is "), seg("enlarged for age in size", True), seg(f" ({size}mm)")]

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
            pv_text = f" Portal vein is prominent in caliber ({pv_mm}mm)."
        elif pv_class == "dilated":
            pv_text = f" Portal vein is dilated in caliber ({pv_mm}mm)."
        else:
            pv_text = " Portal vein is normal in course and caliber"
            if pv_mm:
                pv_text += f" ({pv_mm}mm)"
            pv_text += "."
        s.append(seg(pv_text, True))
    else:
        s.append(seg(" Splenic vein is normal in course and caliber."))

    if d.get("accessory_spleen"):
        asize = d.get("accessory_size_mm", "")
        aloc = d.get("accessory_location", "hilum")
        loc_phrase = {"hilum": "at the splenic hilum",
                      "upper_pole": "at the upper pole",
                      "lower_pole": "at the lower pole"}.get(
            aloc, "at the splenic hilum")
        size_phrase = f"({asize}mm) " if asize else ""
        s += [seg(" "), seg(f"An accessory spleen {size_phrase}is seen "
                             f"{loc_phrase}.", True)]
    return s


# ============================================================
# KIDNEYS
# ============================================================

def _kidney_side_is_present(k):
    return k["status"] == "normal"


def _size_bracket(k, presence):
    r_present, l_present = presence
    r = normalize_kidney_size_text(k["right"].get("size_text", ""))
    l = normalize_kidney_size_text(k["left"].get("size_text", ""))
    parts = []
    if r_present and r:
        parts.append(f"RK={r}")
    if l_present and l:
        parts.append(f"LK={l}")
    if not parts:
        return ""
    return "[" + ";".join(parts) + "]"


def _format_pole(pole_key):
    return POLE_LABELS.get(pole_key, pole_key.replace("_", "-"))


def _kidney_side_has_other_finding(k):
    if k["calculi"]:
        return True
    if k["ureter_calculus"]["count"] != "none":
        return True
    if k.get("cyst_type", "none") != "none":
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
    side_phrase = f"{side_low} kidney"
    if n == 1:
        c = calcs[0]
        out.append(seg(
            f"A calculus measuring {c['size_mm']}mm is seen at the "
            f"{_format_pole(c['pole'])} of {side_phrase}", True))
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
            return f"{sizes[0]}mm at the {pole_txt}"
        elif len(sizes) == 2:
            return f"{sizes[0]}mm & {sizes[1]}mm, both at the {pole_txt}"
        else:
            joined = ", ".join(f"{s}mm" for s in sizes[:-1])
            return f"{joined} & {sizes[-1]}mm, all at the {pole_txt}"

    parts = [group_text(g) for g in groups]
    lead = "A couple of calculi" if n == 2 else "Few calculi"
    if len(parts) == 1:
        joined = parts[0]
    elif len(parts) == 2:
        joined = f"{parts[0]} & {parts[1]}"
    else:
        joined = ", ".join(parts[:-1]) + " & " + parts[-1]
    out.append(seg(f"{lead} seen in the {side_phrase}, largest of these "
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
    side_low = side_key
    grade = uc["grade"]
    sizes = uc["sizes"]
    s = [seg(" ")]
    if uc["count"] == "single":
        size_txt = sizes[0] if sizes else ""
        s.append(seg(f"A calculus measuring {size_txt}mm seen {prep} the "
                     f"{side_low} {body_level}", True))
    elif uc["count"] == "couple":
        a = sizes[0] if len(sizes) > 0 else ""
        b = sizes[1] if len(sizes) > 1 else ""
        s.append(seg(f"A couple of calculi seen {prep} the {side_low} "
                     f"{body_level}, measuring {a}mm & {b}mm", True))
    elif uc["count"] in ("few", "multiple"):
        size_txt = sizes[0] if sizes else ""
        word = "Few" if uc["count"] == "few" else "Multiple"
        s.append(seg(f"{word} calculi seen {prep} the {side_low} "
                     f"{body_level}, largest of these measuring {size_txt}mm", True))
    if grade == "none":
        s.append(seg(".", True))
    elif grade == "no_significant":
        s.append(seg(f", however causing no significant {term.lower()}", True))
        s.append(seg(".", True))
    else:
        g_word = URETER_GRADES_SENTENCE[grade]
        s.append(seg(f", causing {side_low} sided {g_word} {term.lower()}", True))
        s.append(seg(".", True))
    return s


def _ureter_calculus_bilateral_body(k):
    r = k["right"]["ureter_calculus"]
    l = k["left"]["ureter_calculus"]
    if r["count"] == "none" or l["count"] == "none":
        return []

    def max_size(uc):
        nums = [_parse_mm_value(s) for s in uc["sizes"]]
        nums = [n for n in nums if n is not None]
        return max(nums) if nums else 0

    if max_size(l) > max_size(r):
        first_side, second_side = "left", "right"
    else:
        first_side, second_side = "right", "left"

    def side_phrase(side_key):
        uc = k[side_key]["ureter_calculus"]
        _, _, _, _, bi = URETER_LEVELS[uc["level"]]
        rt_lt = "Rt" if side_key == "right" else "Lt"
        sizes_join = " & ".join(f"{s}mm" for s in uc["sizes"] if s)
        return f"{rt_lt} {bi} = {sizes_join}"

    p1 = side_phrase(first_side)
    p2 = side_phrase(second_side)

    def consequence(side_key):
        uc = k[side_key]["ureter_calculus"]
        _, term, _, _, _ = URETER_LEVELS[uc["level"]]
        grade = uc["grade"]
        if grade == "none":
            return None
        if grade == "no_significant":
            return f"no significant {term.lower()} seen on the {side_key}"
        g_word = URETER_GRADES_SENTENCE[grade]
        return f"{side_key} {g_word} {term.lower()}"

    c1 = consequence(first_side)
    c2 = consequence(second_side)
    s = [seg(" "), seg(f"Bilateral ureteric calculi are present ({p1} & {p2})", True)]
    if c1 and c2:
        s.append(seg(f" causing {c1} & {c2}", True))
    elif c1 and not c2:
        s.append(seg(f" causing {c1}, while {c2}", True))
    elif c2 and not c1:
        s.append(seg(f" causing {c2}, while {c1}", True))
    s.append(seg(".", True))
    return s


def _cyst_descriptor_phrase(k):
    septa = k.get("cyst_septa", "none")
    calc = k.get("cyst_calc", "none")
    grade_iif = (septa == "thick") or (calc == "nodular")
    if grade_iif:
        septa_txt = "thick septa" if septa in ("thin", "thick") else None
        calc_txt = "mural nodular calcification" if calc in ("arc", "nodular") else None
    else:
        septa_txt = "thin septa" if septa == "thin" else None
        calc_txt = "mural arc-like calcification" if calc == "arc" else None
    if septa_txt and calc_txt:
        return f"{septa_txt} and {calc_txt}"
    if septa_txt:
        return septa_txt
    if calc_txt:
        return calc_txt
    return None


def _cyst_count_word(count):
    return {"single": "", "few": "Few", "multiple": "Multiple"}.get(count, "")


def _cyst_impression_count_word(count):
    return {"single": "A", "few": "FEW", "multiple": "MULTIPLE"}.get(count, "A")


def _cyst_bosniak(k):
    if k["cyst_type"] == "simple":
        return "I"
    septa = k.get("cyst_septa", "none")
    calc = k.get("cyst_calc", "none")
    if septa == "thick" or calc == "nodular":
        return "IIF"
    return "II"


def _cyst_body(k, side_low):
    ctype = k.get("cyst_type", "none")
    if ctype == "none":
        return []
    count = k.get("cyst_count", "single")
    size = k.get("cyst_size_mm", "")
    loc = k.get("cyst_location", "")
    loc_txt = _format_pole(loc) if loc else ""
    side_phrase = f"{side_low} kidney"
    if ctype == "simple":
        if count == "single":
            text = (f"A simple cortical cyst measuring {size}mm is seen at the "
                    f"{loc_txt} of {side_phrase}.")
        else:
            word = _cyst_count_word(count)
            text = (f"{word} simple cortical cysts seen in the {side_phrase}, "
                    f"largest of these measuring {size}mm at the {loc_txt}.")
    else:
        desc = _cyst_descriptor_phrase(k)
        if not desc:
            return []
        if count == "single":
            text = (f"A complex cyst with {desc}, seen at the {loc_txt} of "
                    f"{side_phrase}.")
        else:
            word = _cyst_count_word(count)
            text = (f"{word} complex cortical cysts with {desc} seen in the "
                    f"{side_phrase}, largest of these measuring {size}mm at the "
                    f"{loc_txt}.")
    return [seg(" "), seg(text, True)]


def _standalone_hydro_body(k, side_low, ub_status):
    if k["hydronephrosis"] == "none":
        return []
    grade = k["hydronephrosis"]
    s = [seg(" "), seg(f"{side_low.capitalize()} {grade} hydroureteronephrosis "
                        f"is present", True)]
    if k.get("ureter_not_traced"):
        s += [seg(f", {side_low} distal ureter could however not be traced", True),
              seg(" ", True), seg("(UB is empty)", True)]
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


def kidneys_sentence(d, ub_status="adequately_distended", age_text=None):
    s = []
    r, l = d["right"], d["left"]
    r_present = _kidney_side_is_present(r)
    l_present = _kidney_side_is_present(l)

    age_years = parse_age(age_text) if age_text is not None else None
    age_ok = (age_years is not None and age_years > SMALL_KIDNEY_MIN_AGE)

    def is_small(side):
        if not age_ok:
            return False
        length = parse_kidney_length_mm(side.get("size_text", ""))
        return length is not None and length <= SMALL_KIDNEY_MM

    r_small = is_small(r) if r_present else False
    l_small = is_small(l) if l_present else False
    r["is_small"] = r_small
    l["is_small"] = l_small

    echo_grade = d.get("cortical_echogenicity", "normal")
    echo_raised = echo_grade in ("mildly_raised", "moderately_raised",
                                  "significantly_raised")
    echo_lat = d.get("cortical_echogenicity_laterality", "bilateral")
    cmd = d.get("cortical_cmd", "preserved")

    if not r_present or not l_present:
        for side_key, side in (("right", r), ("left", l)):
            if side["status"] == "absent_agenesis":
                s += [seg(f"{side_key.upper()} RENAL FOSSA", True, True),
                      seg(" is empty"), seg(" "),
                      seg("(agenesis/hypoplasia)", True), seg("."), seg("\n")]
            elif side["status"] == "absent_ectopic":
                s += [seg(f"{side_key.upper()} RENAL FOSSA", True, True),
                      seg(" is empty. "),
                      seg(f"Ectopic {side_key} kidney", True),
                      seg(f" seen lying in the {side_key} pelvic region."),
                      seg("\n")]
        for side_key, side in (("right", r), ("left", l)):
            if not _kidney_side_is_present(side):
                continue
            s.append(seg(f"{side_key.upper()} KIDNEY", True, True))
            bracket = _size_bracket(d, (side_key == "right", side_key == "left"))
            if side.get("is_small"):
                s += [seg(" appears "),
                      seg("relatively small/contracted in size", True)]
                if bracket:
                    s.append(seg(f" {bracket}"))
                if echo_raised and echo_lat in (side_key, "bilateral"):
                    grade_word = ECHO_GRADE_WORD[echo_grade]
                    side_word = "" if echo_lat == "bilateral" else f"{side_key} "
                    s.append(seg(f" & outline with {grade_word} {side_word}renal "
                                 f"cortical echogenicity with {CMD_WORD[cmd]} "
                                 f"corticomedullary differentiation.", True))
                else:
                    s.append(seg(", however with normal outline and echogenicity. "
                                 "Corticomedullary differentiation is maintained."))
            else:
                if echo_raised and echo_lat in (side_key, "bilateral"):
                    if bracket:
                        s.append(seg(f" is normal in size{bracket} & outline with "))
                    else:
                        s.append(seg(" is normal in size & outline with "))
                    grade_word = ECHO_GRADE_WORD[echo_grade]
                    side_word = "" if echo_lat == "bilateral" else f"{side_key} "
                    s += [seg(f"{grade_word} {side_word}renal cortical echogenicity",
                              True),
                          seg(f" with {CMD_WORD[cmd]} corticomedullary "
                              f"differentiation.")]
                else:
                    if bracket:
                        s.append(seg(f" is normal in size{bracket}, outline and "))
                    else:
                        s.append(seg(" is normal in size, outline and "))
                    s.append(seg("echogenicity. Corticomedullary differentiation "
                                 "is maintained."))
            s.extend(_renal_calculi_body(side, side_key))
            s.extend(_cyst_body(side, side_key))
            s.extend(_standalone_hydro_body(side, side_key, ub_status))
            s.extend(_nephrocalcinosis_body(side, side_key))
            if side["ureter_calculus"]["count"] != "none":
                s.extend(_ureter_calculus_body(side, side_key))
        return s

    both_small = r_small and l_small
    only_r_small = r_small and not l_small
    only_l_small = l_small and not r_small
    r_has_other = _kidney_side_has_other_finding(r)
    l_has_other = _kidney_side_has_other_finding(l)

    unified_small_possible = (both_small and not echo_raised
                              and not r_has_other and not l_has_other)
    if unified_small_possible:
        s.append(seg("Both kidneys appear "))
        s.append(seg("relatively small/contracted in size", True))
        bracket = _size_bracket(d, (True, True))
        if bracket:
            s.append(seg(f" {bracket}"))
        s.append(seg(", however with normal outline and echogenicity. "
                     "Corticomedullary differentiation is maintained."))
        return s

    needs_split = (only_r_small or only_l_small
                   or (both_small and (echo_raised or r_has_other or l_has_other)))
    if needs_split:
        order = ["left", "right"] if only_l_small else ["right", "left"]
        for side_key in order:
            side = r if side_key == "right" else l
            s.append(seg(f"{side_key.upper()} KIDNEY", True, True))
            bracket = _size_bracket(d, (side_key == "right", side_key == "left"))
            echo_on_this = echo_raised and echo_lat in (side_key, "bilateral")
            small = (r_small if side_key == "right" else l_small)
            if small:
                s += [seg(" appears "),
                      seg("relatively small/contracted in size", True)]
                if bracket:
                    s.append(seg(f" {bracket}"))
                if echo_on_this:
                    grade_word = ECHO_GRADE_WORD[echo_grade]
                    side_word = "" if echo_lat == "bilateral" else f"{side_key} "
                    s.append(seg(f" & outline with {grade_word} {side_word}renal "
                                 f"cortical echogenicity with {CMD_WORD[cmd]} "
                                 f"corticomedullary differentiation.", True))
                else:
                    s.append(seg(", however with normal outline and echogenicity. "
                                 "Corticomedullary differentiation is maintained."))
            else:
                if echo_on_this:
                    if bracket:
                        s.append(seg(f" is normal in size{bracket} & outline with "))
                    else:
                        s.append(seg(" is normal in size & outline with "))
                    grade_word = ECHO_GRADE_WORD[echo_grade]
                    side_word = "" if echo_lat == "bilateral" else f"{side_key} "
                    s += [seg(f"{grade_word} {side_word}renal cortical echogenicity",
                              True),
                          seg(f" with {CMD_WORD[cmd]} corticomedullary "
                              f"differentiation.")]
                else:
                    if bracket:
                        s.append(seg(f" is normal in size{bracket}, outline and "))
                    else:
                        s.append(seg(" is normal in size, outline and "))
                    s.append(seg("echogenicity. Corticomedullary differentiation "
                                 "is maintained."))
            s.extend(_renal_calculi_body(side, side_key))
            s.extend(_cyst_body(side, side_key))
            s.extend(_standalone_hydro_body(side, side_key, ub_status))
            s.extend(_nephrocalcinosis_body(side, side_key))
            if side["ureter_calculus"]["count"] != "none":
                s.extend(_ureter_calculus_body(side, side_key))
        return s

    s.append(seg("BOTH KIDNEYS", True, True))
    bracket = _size_bracket(d, (True, True))
    if echo_raised:
        if bracket:
            s.append(seg(f" are normal in size{bracket} & outline with "))
        else:
            s.append(seg(" are normal in size & outline with "))
        grade_word = ECHO_GRADE_WORD[echo_grade]
        side_word = "" if echo_lat == "bilateral" else f"{echo_lat} "
        s += [seg(f"{grade_word} {side_word}renal cortical echogenicity", True),
              seg(f" with {CMD_WORD[cmd]} corticomedullary differentiation.")]
    else:
        if bracket:
            s.append(seg(f" are normal in size{bracket}, outline and "))
        else:
            s.append(seg(" are normal in size, outline and "))
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


# ============================================================
# URINARY BLADDER
# ============================================================

def _ub_suboptimal_tail():
    return [seg(" "), seg("(suboptimal pelvic assessment)", bold=False, italic=True)]


def _ub_sedimentation_body(d):
    out = []
    ff = bool(d.get("sedimentation_free_floating", False))
    sd = bool(d.get("sedimentation_settled_debris", False))
    trace = bool(d.get("sedimentation_trace", False))
    extensive = bool(d.get("sedimentation_extensive", False))

    if trace:
        out += [seg(" "), seg("Trace sedimentation seen in the UB lumen.", True)]
    if ff and sd:
        out += [seg(" "), seg("Settled debris in the dependent part of the UB "
                              "lumen along with free floating sedimentation.", True)]
    elif ff:
        out += [seg(" "), seg("Free floating sedimentation in the UB lumen.", True)]
    elif sd:
        out += [seg(" "), seg("Settled debris in the dependent part of the UB "
                              "lumen.", True)]
    if extensive:
        out += [seg(" "), seg("Extensive sedimentation seen in the UB lumen.", True)]
    return out


def urinary_bladder_sentence(d):
    forced = d.get("force_empty_suboptimal", False)
    s = [seg("URINARY BLADDER", True, True)]
    status = d.get("status")
    cath = d.get("catheterized", False)
    pre = str(d.get("pre_void_cc", "") or "").strip()
    post = str(d.get("post_void_cc", "") or "").strip()

    if forced:
        s.append(seg(" is empty"))
        s.extend(_ub_suboptimal_tail())
        s.append(seg("."))
        if not d.get("mass_calculus"):
            s.append(seg(" No mass or calculus seen."))
        return s

    if status is None and not cath:
        s.append(seg(" is adequately distended."))
    elif status is None and cath:
        s.append(seg(" is adequately distended and catheterized."))
    elif status == "over":
        s += [seg(" is "), seg("over-distended", True)]
        if pre:
            s.append(seg(f"({pre}CC)", True))
        if cath:
            s.append(seg(" and catheterized", True))
        s.append(seg("."))
    elif status == "empty":
        s.append(seg(" is empty"))
        if cath:
            s.append(seg(" & catheterized"))
        s.extend(_ub_suboptimal_tail())
        s.append(seg("."))
    elif status == "partially_empty":
        s.append(seg(" is partially empty"))
        if cath:
            s.append(seg(" & catheterized"))
        s.extend(_ub_suboptimal_tail())
        s.append(seg("."))
    else:
        s.append(seg(" is adequately distended."))

    if not d.get("mass_calculus"):
        s.append(seg(" No mass or calculus seen."))

    if d.get("wall_thickening") and d.get("wall_mm"):
        try:
            w = float(d["wall_mm"])
        except (ValueError, TypeError):
            w = 0
        prefix = "Mild" if w <= 8 else "Significant"
        irregular = d.get("wall_irregular", False)
        frag = f"{prefix} "
        if irregular:
            frag += "irregular "
        frag += f"urinary bladder wall thickening upto {d['wall_mm']}mm is seen."
        s += [seg(" "), seg(frag, True)]

    s.extend(_ub_sedimentation_body(d))

    if pre:
        s.append(seg("\n"))
        s.append(seg(f"URINARY BLADDER PRE-VOID VOLUME({pre}CC)", True))
        s.append(seg("\n"))
        if post:
            s.append(seg(f"POST-VOID RESIDUE -- {post}CC", True))
        else:
            s.append(seg("POST VOID RESIDUE - ", True))
    return s


def _pvr_pct(d):
    pre = _parse_mm_value(d.get("pre_void_cc", ""))
    post = _parse_mm_value(d.get("post_void_cc", ""))
    if pre and post and pre > 0:
        return post / pre * 100.0
    return None


def _pvr_significant(d):
    pct = _pvr_pct(d)
    if pct is None:
        return False
    return pct >= 15.0


def _pvr_insignificant(d):
    pct = _pvr_pct(d)
    if pct is None:
        return False
    return pct < 15.0


def _ub_impression_lines(d, prostate_line=None):
    out = []
    ff = bool(d.get("sedimentation_free_floating", False))
    sd = bool(d.get("sedimentation_settled_debris", False))
    trace = bool(d.get("sedimentation_trace", False))
    extensive = bool(d.get("sedimentation_extensive", False))
    uti = bool(d.get("uti_suspected", False))
    wall = d.get("wall_thickening") and d.get("wall_mm")

    sed_bits = []
    if trace:
        sed_bits.append("TRACE SEDIMENTATION")
    if ff and sd:
        sed_bits.append("SETTLED DEBRIS IN THE DEPENDENT PART OF THE URINARY "
                        "BLADDER LUMEN ALONG WITH FREE FLOATING SEDIMENTATION")
    elif ff:
        sed_bits.append("FREE FLOATING SEDIMENTATION")
    elif sd:
        sed_bits.append("SETTLED DEBRIS IN THE DEPENDENT PART OF THE URINARY "
                        "BLADDER LUMEN")
    if extensive:
        sed_bits.append("EXTENSIVE SEDIMENTATION")

    if wall:
        try:
            w = float(d["wall_mm"])
        except (ValueError, TypeError):
            w = 0
        prefix = "MILD" if w <= 8 else "SIGNIFICANT"
        irregular = "IRREGULAR " if d.get("wall_irregular", False) else ""
        base = (f"{prefix} {irregular}URINARY BLADDER WALL THICKENING"
                f"(UPTO {d['wall_mm']}MM)")
        if ff and sd:
            tail = "-- ?ACUTE ON CHRONIC CYSTITIS"
        else:
            tail = "-- ?CHRONIC CYSTITIS"
        body = base
        if sed_bits:
            body += " WITH " + " & ".join(sed_bits) + " IN THE UB LUMEN"
        body += f" {tail}"
        out.append(body + " Adv- Urine R/M Correlation.")
        return out

    if sed_bits:
        sed_line = (" ".join(sed_bits) + " IN THE UB LUMEN")
        if uti:
            sed_line += " - ?UTI"
        out.append(sed_line + ". Adv- Urine R/M Correlation.")

    pre = str(d.get("pre_void_cc", "") or "").strip()
    post = str(d.get("post_void_cc", "") or "").strip()
    status = d.get("status")

    if status == "over" and pre:
        out.append(f"OVERDISTENDED URINARY BLADDER(PRE-VOID VOLUME={pre}CC).")
    elif pre and not post:
        out.append(f"URINARY BLADDER PRE-VOID VOLUME({pre}CC)")

    if pre and post:
        if _pvr_significant(d):
            pvr_txt = "SIGNIFICANT POST VOID RESIDUE"
            pct = _pvr_pct(d)
            if pct is not None and pct >= 80.0:
                pvr_txt += "(>80%)"
            if prostate_line:
                out.append(f"{pvr_txt} WITH {prostate_line.rstrip('.')}.")
            else:
                out.append(pvr_txt + ".")
        elif _pvr_insignificant(d):
            if prostate_line:
                prostate_wo_dot = prostate_line.rstrip('.')
                out.append(f"{prostate_wo_dot}, HOWEVER THE POST-VOID RESIDUE "
                           f"IS INSIGNIFICANT.")
            else:
                out.append("INSIGNIFICANT POST-VOID RESIDUE.")
    return out


# ============================================================
# PROSTATE
# ============================================================

def _prostate_band(d):
    return prostate_band_from_volume(d.get("size_cc", ""))


def _prostate_cyst_body(d):
    if not d.get("cyst"):
        return None
    size = d.get("cyst_size", "")
    side = ""
    if d.get("cyst_left_hemi") and not d.get("cyst_right_hemi"):
        side = "left hemi-prostate "
    elif d.get("cyst_right_hemi") and not d.get("cyst_left_hemi"):
        side = "right hemi-prostate "
    size_txt = f"({size}mm)" if size else ""
    return f"A simple cystic SOL{size_txt} seen in the {side}prostate."


def _prostate_median_lobe_body_frag(d):
    if not d.get("median_lobe", False):
        return None
    size = d.get("median_lobe_size_mm", "")
    boo = d.get("median_lobe_boo", False)
    size_txt = f"({size}mm)" if size else ""
    boo_txt = ", impinging upon the bladder outlet" if boo else ""
    return f"median lobe hypertrophy{size_txt}{boo_txt}"


def prostate_sentence(d, pediatric=False, ub_status=None):
    s = [seg("PROSTATE", True, True)]
    if d.get("not_visualized"):
        s.append(seg(" is not visualized."))
        return s
    if d.get("partial_visualization"):
        s.append(seg(" is partially visualized."))
        return s

    if pediatric and not _prostate_band(d):
        s.append(seg(" is age-appropriate."))
        return s

    band = _prostate_band(d)
    median_frag = _prostate_median_lobe_body_frag(d)
    cyst_frag = _prostate_cyst_body(d)
    cc = d.get("size_cc", "")

    if band is None:
        if median_frag:
            s.append(seg(" is normal in size"))
            s += [seg(" with "), seg(median_frag, True),
                  seg(", however the attenuation and margins are normal.")]
        else:
            s.append(seg(" is normal in size and attenuation. No diffuse or "
                         "focal lesion seen."))
        if cyst_frag:
            s += [seg(" "), seg(cyst_frag, True)]
        return s

    body_desc, _ = band
    s += [seg(" is "), seg(f"{body_desc} in size", True)]
    if cc:
        s.append(seg(f"({cc}CC)"))
    if median_frag:
        s += [seg(" with "), seg(median_frag, True),
              seg(", however the attenuation and margins are normal.")]
    else:
        s.append(seg(", with normal attenuation and margins."))
    if cyst_frag:
        s += [seg(" "), seg(cyst_frag, True)]
    return s


def _prostate_impression_line(d):
    if d.get("not_visualized") or d.get("partial_visualization"):
        return None
    band = _prostate_band(d)
    cc = d.get("size_cc", "")
    base = None
    if band:
        _, imp_label = band
        base = f"{imp_label}({cc}CC)" if cc else imp_label

    median = d.get("median_lobe", False)
    median_size = d.get("median_lobe_size_mm", "")
    boo = d.get("median_lobe_boo", False)
    median_frag = None
    if median:
        frag = "MEDIAN LOBE HYPERTROPHY"
        if median_size:
            frag += f"({median_size}MM)"
        if boo:
            frag += " IMPINGING UPON BLADDER OUTLET"
        median_frag = frag

    cyst_frag = None
    if d.get("cyst"):
        size = d.get("cyst_size", "")
        side_bits = []
        if d.get("cyst_left_hemi"):
            side_bits.append("LEFT HEMI-PROSTATE")
        if d.get("cyst_right_hemi"):
            side_bits.append("RIGHT HEMI-PROSTATE")
        side_txt = ""
        if len(side_bits) == 1:
            side_txt = side_bits[0] + " "
        elif len(side_bits) == 2:
            side_txt = "BILATERAL HEMI-PROSTATE "
        size_txt = f"({size}MM)" if size else ""
        cyst_frag = (f"A SIMPLE CYSTIC SOL{size_txt} IN THE {side_txt}PROSTATE")

    if base and median_frag:
        base = base + " WITH " + median_frag
    elif not base and median_frag:
        base = median_frag

    if base and cyst_frag:
        return base + " WITH " + cyst_frag + "."
    if base:
        return base + "."
    if cyst_frag:
        return cyst_frag + "."
    return None


# ============================================================
# UTERUS
# ============================================================

def _uterus_size_body_prefix(d):
    length = _first_number(d.get("size", ""))
    cls = classify_uterus_size(length)
    return cls, UTERUS_SIZE_BODY.get(cls, "normal in size")


def _uterus_myometrium_clause(d):
    m = d.get("myometrium", "homogenous")
    if m == "homogenous":
        return ("Myometrium appears homogenous and no focal lesion is seen.", False, False)
    if m == "mildly_heterogeneous":
        return ("Myometrium appears mildly heterogenous and no focal lesion is seen.", False, False)
    if m == "heterogeneous":
        return ("Myometrium appears heterogenous and no focal lesion is seen.", False, False)
    return ("", False, False)


def _fibroid_lesion_sort_key(lesion):
    return -(_first_number(lesion.get("size", "")) or 0)


def _fibroid_body_fragment(d):
    if not d.get("fibroid"):
        return None
    count = d.get("fibroid_count", "single")
    types = d.get("fibroid_types", []) or ["intramural"]
    lesions = sorted(d.get("fibroid_lesions", []), key=_fibroid_lesion_sort_key)
    while len(lesions) < len(types):
        lesions.append(_new_fibroid_lesion())
    lesions = lesions[:len(types)]

    body_parts = []
    for i, t in enumerate(types):
        lesion = lesions[i]
        loc = lesion.get("location", "fundal")
        size = lesion.get("size", "")
        is_partial_ext = t in PARTIALLY_EXTRUSIVE_TYPES
        if t == "submucosal_intramural":
            body_parts.append(
                f"a hypoechoic heterogenous SOL measuring {size}mm in the "
                f"{loc} myometrium, impinging upon the endometrium")
        elif is_partial_ext:
            body_parts.append(
                f"a partially extrusive hypoechoic heterogenous SOL measuring "
                f"{size}mm in the {loc} myometrium")
        else:
            body_parts.append(
                f"a hypoechoic heterogenous SOL in the {loc} myometrium "
                f"measuring {size}mm")

    if len(body_parts) == 1:
        joined = body_parts[0]
    elif len(body_parts) == 2:
        joined = f"{body_parts[0]} and {body_parts[1]}"
    else:
        joined = ", ".join(body_parts[:-1]) + " and " + body_parts[-1]

    if count == "single":
        return [seg(" with "), seg(joined, True)]
    else:
        count_word = {"few": "few", "multiple": "multiple"}.get(count, "few")
        desc_list = []
        for i, t in enumerate(types):
            lesion = lesions[i]
            size = lesion.get("size", "")
            loc = lesion.get("location", "fundal")
            suffix = ", impinging upon the endometrium" if t == "submucosal_intramural" else ""
            desc_list.append(f"{size}mm in {loc} myometrium{suffix}")
        if len(desc_list) == 1:
            joined_desc = desc_list[0]
        elif len(desc_list) == 2:
            joined_desc = f"{desc_list[0]} and {desc_list[1]}"
        else:
            joined_desc = ", ".join(desc_list[:-1]) + " and " + desc_list[-1]
        return [seg(" with "), seg(f"{count_word} hypoechoic heterogenous SOLs "
                                    f"in the myometrium, largest of these "
                                    f"measuring {joined_desc}", True)]


def _fibroid_impression_fragment(d):
    if not d.get("fibroid"):
        return None
    count = d.get("fibroid_count", "single")
    types = d.get("fibroid_types", []) or ["intramural"]
    lesions = sorted(d.get("fibroid_lesions", []), key=_fibroid_lesion_sort_key)
    while len(lesions) < len(types):
        lesions.append(_new_fibroid_lesion())
    lesions = lesions[:len(types)]

    if len(types) > 1:
        figo_manual = d.get("fibroid_figo_manual", "")
        if count == "single":
            size0 = lesions[0].get("size", "") if lesions else ""
            loc0 = lesions[0].get("location", "fundal") if lesions else "fundal"
            figo_part = f"(FIGO- {figo_manual})" if figo_manual else "(FIGO- 4-5)"
            return (f"SUGGESTION OF A FIBROID({size0}MM) IN THE {loc0.upper()} "
                    f"MYOMETRIUM {figo_part}")
        else:
            figo_part = f"(FIGO - {figo_manual})" if figo_manual else "(FIGO - 4-5)"
            count_word = "FEW" if count == "few" else "MULTIPLE"
            return (f"SUGGESTION OF {count_word} FIBROIDS {figo_part} AS "
                    f"DESCRIBED ABOVE")

    t = types[0]
    label, figo = FIBROID_TYPE_LABELS_UP.get(t, FIBROID_TYPE_LABELS_UP["intramural"])
    if count == "single":
        size = lesions[0].get("size", "") if lesions else ""
        loc = lesions[0].get("location", "fundal") if lesions else "fundal"
        return (f"SUGGESTION OF {label}({size}MM) IN THE {loc.upper()} "
                f"MYOMETRIUM({figo})")
    else:
        count_word = "FEW" if count == "few" else "MULTIPLE"
        locs = [lesion.get("location", "fundal").upper() for lesion in lesions]
        seen = set()
        locs_u = []
        for loc in locs:
            if loc not in seen:
                locs_u.append(loc)
                seen.add(loc)
        if len(locs_u) == 1:
            locs_txt = f"THE {locs_u[0]} MYOMETRIUM"
        elif len(locs_u) == 2:
            locs_txt = f"THE {locs_u[0]} & {locs_u[1]} MYOMETRIUM"
        else:
            locs_txt = "THE " + ", ".join(locs_u[:-1]) + " & " + locs_u[-1] + " MYOMETRIUM"
        return (f"SUGGESTION OF {count_word} FIBROIDS IN {locs_txt} AS DESCRIBED "
                f"ABOVE({figo})")


def _adenomyosis_features_body(d):
    feats = d.get("adenomyosis_features", []) or []
    frags = []
    for f in ADENOMYOSIS_FEATURES:
        if f not in feats:
            continue
        if f == "globular_shape":
            frags.append("with globular shape")
        elif f == "asymmetric_posterior":
            frags.append("with asymmetrically bulky posterior myometrium")
        elif f == "venetian_blind_sign":
            frags.append("diffusely heterogeneous echo pattern(venetian blind sign "
                         "is positive) of myometrium")
        elif f == "interface_indistinct":
            frags.append("with indistinct endo-myometrial interface")
        elif f == "interface_barely_perceptible_fundus":
            frags.append("endo-myometrial interface is indistinct and barely "
                         "perceptible at fundus level")
        elif f == "interface_lost":
            frags.append("loss of endo-myometrial interface")
        elif f == "subendometrial_cysts":
            frags.append("with few subendometrial cysts abutting the endometrium")
    return frags


def _adenomyosis_confidence(d):
    conf = d.get("adenomyosis_confidence", "auto")
    if conf in ("likely", "early"):
        return conf
    feats = d.get("adenomyosis_features", []) or []
    return "likely" if len(feats) >= 2 else "early"


def _adenomyosis_body_segments(d):
    if not d.get("adenomyosis"):
        return []
    frags = _adenomyosis_features_body(d)
    conf = _adenomyosis_confidence(d)
    qualifier = "likely adenomyosis" if conf == "likely" else "?early adenomyosis"
    body = " ".join(frags)
    out = []
    if body:
        out += [seg(" "), seg(body, True)]
    out += [seg(f" -- {qualifier}.", True)]
    return out


def _adenomyosis_impression_fragment(d):
    if not d.get("adenomyosis"):
        return None
    frags = [f.upper() for f in _adenomyosis_features_body(d)]
    conf = _adenomyosis_confidence(d)
    qualifier = ("LIKELY UTERINE ADENOMYOSIS" if conf == "likely"
                 else "?EARLY UTERINE ADENOMYOSIS")
    if not frags:
        return qualifier
    return f"{' & '.join(frags)} ---- {qualifier}"


def _adenomyomas_body_fragment(d):
    if not d.get("adenomyomas"):
        return None
    count = d.get("adenomyomas_count", "couple")
    echo = d.get("adenomyomas_echo", "hyperechoic")
    loc = d.get("adenomyomas_location", "posterior")
    avasc = "avascular " if d.get("adenomyomas_avascular") else ""
    count_word = {"single": "A", "couple": "A couple of",
                  "few": "Few", "multiple": "Multiple"}.get(count, "A couple of")
    plural = "SOL" if count == "single" else "SOLs"
    size_txt = ""
    if count == "single":
        size = d.get("adenomyomas_size", "")
        if size:
            size_txt = f"({size}mm) "
    text = (f"{count_word} {avasc}{echo} small {plural} {size_txt}seen abutting "
            f"the {loc} endometrium -- likely adenomyoma")
    if count != "single":
        text += "s"
    text += "."
    return text


def _adenomyomas_impression_fragment(d):
    if not d.get("adenomyomas"):
        return None
    count = d.get("adenomyomas_count", "couple")
    echo = d.get("adenomyomas_echo", "hyperechoic").upper()
    loc = d.get("adenomyomas_location", "posterior").upper()
    avasc = "AVASCULAR " if d.get("adenomyomas_avascular") else ""
    count_word = {"single": "A", "couple": "A COUPLE OF",
                  "few": "FEW", "multiple": "MULTIPLE"}.get(count, "A COUPLE OF")
    plural = "SOL" if count == "single" else "SOLS"
    size_txt = ""
    if count == "single":
        size = d.get("adenomyomas_size", "")
        if size:
            size_txt = f"({size}MM) "
    frag = (f"{count_word} {avasc}{echo} SMALL {plural} {size_txt}SEEN ABUTTING "
            f"THE {loc} UTERINE ENDOMETRIUM -- LIKELY ADENOMYOMA")
    if count != "single":
        frag += "S"
    return frag


def uterus_sentence(d, pediatric=False, age_years=None):
    status = d.get("status", "anteverted")

    if pediatric:
        s = [seg("UTERUS & BOTH OVARIES", True, True),
             seg(" are age-appropriate.")]
        return s

    if status == "not_visualized":
        return [seg("UTERUS", True, True), seg(" is not visualized.")]

    if status == "operated":
        return [seg("UTERUS", True, True), seg(" is operated.")]

    s = [seg("UTERUS", True, True)]
    size = d.get("size", "")
    length = _first_number(size)
    cls = classify_uterus_size(length)
    descriptor = ""
    if cls == "bulky":
        descriptor = "bulky in size"
    elif cls == "moderately_bulky":
        descriptor = "moderately bulky in size"
    elif cls == "grossly_bulky":
        descriptor = "grossly bulky in size"

    retroflexed = d.get("retroflexed", False)
    gravid = d.get("gravid", False)
    low_lying = d.get("low_lying", False)

    if gravid:
        gravid_type = d.get("gravid_type", "CRL")
        length_v = d.get("gravid_length", "")
        wk = d.get("gravid_weeks", "")
        dy = d.get("gravid_days", "")
        s.append(seg(f" is gravid(Single & alive: {gravid_type}={length_v} mm; "
                     f"GA={wk} weeks {dy} days)."))
    else:
        orientation = "retro-flexed" if retroflexed else "anteverted"
        if status == "partially_visualized":
            s += [seg(" is partially visualized and appears "),
                  seg(orientation, True)]
        else:
            s += [seg(" is "), seg(orientation, True)]

        if descriptor and size:
            s += [seg(" and "), seg(descriptor, True), seg(f"({size}mm)")]
        elif size:
            s.append(seg(f" and normal in size({size}mm)"))
        else:
            s.append(seg(" and normal in size"))

        if low_lying:
            cervix_state = d.get("cervix_visualized", "partially")
            cervix_word = ("partially visualized" if cervix_state == "partially"
                           else "not visualized")
            grade_word = "GRADE-II" if cervix_state == "partially" else "GRADE-II/III"
            s += [seg(" and "),
                  seg("low-lying with positive transverse uterus sign, cervix is "
                      f"{cervix_word} in the retro-vesical space on TAS - likely "
                      f"{grade_word.lower()} uterine prolapse", True)]
        s.append(seg(" with normal shape and echopattern."))

    myo_text, myo_b, myo_i = _uterus_myometrium_clause(d)
    if myo_text:
        s.append(seg(" " + myo_text, myo_b, italic=myo_i if myo_b else None))

    aden_segs = _adenomyosis_body_segments(d)
    if aden_segs:
        s.extend(aden_segs)

    if d.get("prominent_vascular_channels"):
        s += [seg(" "), seg("Prominent vascular channels are seen along the "
                             "peripheral uterine myometrium.", True)]

    fib_segs = _fibroid_body_fragment(d)
    if fib_segs:
        if myo_text and not myo_text.endswith("."):
            s.append(seg("."))
        s.extend(fib_segs)
        s.append(seg("."))

    adenoma_text = _adenomyomas_body_fragment(d)
    if adenoma_text:
        s += [seg(" "), seg(adenoma_text, True)]

    endo_mm = d.get("endometrial_thickness_mm", "")
    endo_cls = classify_endometrium(_first_number(endo_mm))
    if endo_mm:
        endo_descriptor = ""
        if endo_cls == "thinned_out":
            endo_descriptor = " (thinned out endometrium)"
        elif endo_cls == "thickened":
            endo_descriptor = " (thickened endometrium)"
        elif endo_cls == "significantly_thickened":
            endo_descriptor = " (significantly thickened endometrium)"
        s.append(seg(f" Endometrial echo is central and regular in "
                     f"thickness({endo_mm}mm){endo_descriptor}."))
    coll = d.get("endometrial_collection", "none")
    het = d.get("endometrial_collection_het", False)
    if coll == "none" and not het:
        s.append(seg(" No collection seen in the endometrial canal."))
    elif coll in ("mild", "moderate"):
        if het:
            s.append(seg(" " + ENDOMETRIAL_COLLECTION_HET[coll]))
        else:
            s.append(seg(" " + ENDOMETRIAL_COLLECTION_LABELS[coll]))

    if d.get("rpoc"):
        size_txt = d.get("rpoc_size", "")
        echo = d.get("rpoc_echo", "hypo")
        echo_word = "hypo-echoic" if echo == "hypo" else "hyper-echoic"
        first_dim = _first_number(size_txt.split("x")[0] if size_txt else "")
        if first_dim is not None and first_dim > 15:
            s += [seg(" "),
                  seg(f"A moderate sized ({size_txt}mm) {echo_word} structure "
                      f"seen abutting the fundal endometrium showing color flow "
                      f"& aliasing on CD study.", True)]
        else:
            s += [seg(" "),
                  seg(f"A small ({size_txt}mm) {echo_word} SOL seen abutting "
                      f"the fundal endometrium showing color flow & aliasing on "
                      f"CD study.", True)]

    if d.get("cervix_elongated"):
        s += [seg(" "), seg("Cervix appears elongated.", True)]
    if d.get("cervix_bulky"):
        sz = d.get("cervix_bulky_size_mm", "")
        sz_txt = f"({sz}mm)" if sz else ""
        s += [seg(" "), seg(f"Cervix appears bulky{sz_txt}.", True)]
    nab = d.get("nabothian", "none")
    if nab != "none":
        word = {"one": "A", "few": "Few", "multiple": "Multiple"}.get(nab, "Few")
        s += [seg(" "), seg(f"{word} nabothian cysts seen in the cervix.", True)]
    cx_coll = d.get("cervix_collection", "none")
    if cx_coll != "none":
        s += [seg(" "), seg(f"{cx_coll.title()} collection seen in the cervix.", True)]

    return s


def _uterus_impression_lines(d, age_years):
    lines = []
    if d.get("status") in ("not_visualized", "operated"):
        return lines

    size = d.get("size", "")
    length = _first_number(size)
    cls = classify_uterus_size(length)
    leading = UTERUS_SIZE_IMPRESSION.get(cls, "")

    if d.get("gravid"):
        gravid_type = d.get("gravid_type", "CRL")
        length_v = d.get("gravid_length", "")
        wk = d.get("gravid_weeks", "")
        dy = d.get("gravid_days", "")
        return [f"GRAVID UTERUS (SINGLE & ALIVE: {gravid_type}={length_v} MM; "
                f"GA={wk} WEEKS {dy} DAYS)."]

    low_lying = d.get("low_lying", False)
    if low_lying:
        cervix_state = d.get("cervix_visualized", "partially")
        cervix_word = ("PARTIALLY VISUALIZED" if cervix_state == "partially"
                       else "NOT VISUALIZED")
        grade_word = "GRADE-II" if cervix_state == "partially" else "GRADE-II/III"
        adj = UTERUS_SIZE_IMPRESSION_ADJ.get(cls, "")
        leading = f"{adj} LOW-LYING UTERUS" if adj else "LOW-LYING UTERUS"
        lines.append(f"{leading} WITH POSITIVE TRANSVERSE UTERUS SIGN, CERVIX "
                     f"IS {cervix_word} IN THE RETRO-VESICAL SPACE ON TAS -- "
                     f"LIKELY {grade_word} UTERINE PROLAPSE.")

    aden_frag = _adenomyosis_impression_fragment(d) if d.get("adenomyosis") else None
    fib_frag = _fibroid_impression_fragment(d)
    adenoma_frag = _adenomyomas_impression_fragment(d)

    myo = d.get("myometrium", "homogenous")
    myo_prefix = ""
    if myo == "mildly_heterogeneous":
        myo_prefix = "MILDLY HETEROGENOUS MYOMETRIUM"
    elif myo == "heterogeneous":
        myo_prefix = "HETEROGENOUS MYOMETRIUM"

    base_prefix_parts = []
    if leading and not low_lying:
        base_prefix_parts.append(leading.rstrip("."))

    combined_prefix = ""
    if base_prefix_parts:
        combined_prefix = base_prefix_parts[0]
    if myo_prefix and not low_lying:
        combined_prefix = (f"{combined_prefix} WITH {myo_prefix}"
                           if combined_prefix else myo_prefix)

    fragments_to_attach = []
    if aden_frag:
        fragments_to_attach.append(aden_frag)
    if fib_frag:
        fragments_to_attach.append(fib_frag)
    if adenoma_frag:
        fragments_to_attach.append(adenoma_frag)

    if low_lying:
        if fib_frag and not aden_frag and not adenoma_frag:
            lines[-1] = lines[-1] + " ALSO THERE IS " + fib_frag + "."
        elif fib_frag and (aden_frag or adenoma_frag):
            if aden_frag:
                lines.append(f"ALSO THERE IS {aden_frag}.")
            if fib_frag:
                lines.append(f"ALSO THERE IS {fib_frag}.")
            if adenoma_frag:
                lines.append(f"ALSO THERE IS {adenoma_frag}.")
        elif aden_frag:
            lines[-1] = lines[-1] + " ALSO THERE IS " + aden_frag + "."
        elif adenoma_frag:
            lines[-1] = lines[-1] + " ALSO THERE IS " + adenoma_frag + "."
        return lines

    if not fragments_to_attach:
        if combined_prefix:
            lines.append(combined_prefix + ".")
        return lines

    assembled = combined_prefix
    for frag in fragments_to_attach:
        if assembled:
            assembled = assembled + " WITH " + frag
        else:
            assembled = "THERE IS " + frag

    if not combined_prefix and assembled and not assembled.startswith("THERE IS"):
        assembled = "THERE IS " + assembled

    if combined_prefix and not assembled.startswith(combined_prefix):
        assembled = combined_prefix + " WITH " + assembled

    lines.append(assembled + ".")
    return lines


# ============================================================
# OVARIES
# ============================================================

def _ovary_present(side):
    return side.get("status", "normal") == "normal"


def _ovary_is_bulky(side):
    v = _parse_mm_value(side.get("volume_cc", ""))
    return v is not None and v >= OVARY_BULKY_CC


def _ovary_side_label(side_key):
    return "RIGHT" if side_key == "right" else "LEFT"


def _ovary_findings_count_word(count, up=False):
    mapping = {"single": ("", "A"),
               "few": ("Few", "FEW"),
               "multiple": ("Multiple", "MULTIPLE")}
    lower, upper = mapping.get(count, ("", "A"))
    return upper if up else lower


def _ovary_finding_phrase(finding, side_low, up=False):
    ftype = finding.get("type", "")
    if ftype not in OVARY_CYST_TYPES:
        return None
    size = finding.get("size_mm", "")
    count = finding.get("count", "single")
    large = False
    if size:
        m = re.search(r"(\d+(?:\.\d+)?)", str(size))
        if m:
            large = float(m.group(1)) >= OVARY_LARGE_MM

    type_word_body = OVARY_CYST_TYPES[ftype]
    type_word_up_sing = OVARY_CYST_TYPES_UP[ftype]
    type_word_up_plural = OVARY_CYST_TYPES_UP_PLURAL[ftype]

    if up:
        lead = _ovary_findings_count_word(count, up=True)
        large_word = "LARGE " if large else ""
        type_word = type_word_up_sing if count == "single" else type_word_up_plural
        side_up = _ovary_side_label(side_low)
        return f"{lead} {side_up} OVARIAN {large_word}{type_word}({size}MM)"
    else:
        large_word = "large " if large else ""
        if count == "single":
            return (f"A {large_word}{type_word_body}({size}mm) seen in the "
                    f"{side_low} ovary.")
        else:
            word = _ovary_findings_count_word(count)
            plural = type_word_body + "s"
            return (f"{word} {large_word}{plural} seen in the {side_low} ovary, "
                    f"largest of these measuring {size}mm.")


def _ovary_grouped_impression_lines(d, up_ctx=None):
    lines = []
    r = d["right"]
    l = d["left"]
    r_findings = [f for f in r.get("findings", []) if f.get("type") in OVARY_CYST_TYPES]
    l_findings = [f for f in l.get("findings", []) if f.get("type") in OVARY_CYST_TYPES]

    r_types = [f.get("type") for f in r_findings]
    l_types = [f.get("type") for f in l_findings]
    shared_types = [t for t in r_types if t in l_types]

    handled = set()
    for t in shared_types:
        if t in handled:
            continue
        handled.add(t)
        r_f = next(f for f in r_findings if f.get("type") == t)
        l_f = next(f for f in l_findings if f.get("type") == t)
        r_size = r_f.get("size_mm", "")
        l_size = l_f.get("size_mm", "")
        type_up_plural = OVARY_CYST_TYPES_UP_PLURAL[t]
        lines.append(f"BILATERAL OVARIAN {type_up_plural}({r_size}MM IN RIGHT "
                     f"OVARY & {l_size}MM IN LEFT OVARY).")

    for side_key, findings in (("right", r_findings), ("left", l_findings)):
        for f in findings:
            if f.get("type") in handled:
                continue
            phrase = _ovary_finding_phrase(f, side_key, up=True)
            if phrase:
                lines.append(phrase + ".")
    return lines


def _pcos_volume_pattern(d):
    r_bulky = _ovary_is_bulky(d["right"])
    l_bulky = _ovary_is_bulky(d["left"])
    if r_bulky and l_bulky:
        return "both_bulky"
    if r_bulky and not l_bulky:
        return "right_bulky"
    if l_bulky and not r_bulky:
        return "left_bulky"
    return "both_normal"


def _pcos_feature_phrase_body(d):
    parts = []
    if d.get("pcos_feature4"):
        var = d.get("pcos_feature4_variable", False)
        periph = d.get("pcos_feature4_peripheral", False)
        rand = d.get("pcos_feature4_random", False)
        if var and periph:
            parts.append("predominantly peripheral distribution of variable sized follicles")
        elif var and rand:
            parts.append("random distribution of variable sized follicles")
        elif var:
            parts.append("variable sized follicles")
        elif periph:
            parts.append("predominantly peripheral distribution of follicles")
        elif rand:
            parts.append("random distribution of follicles")
    else:
        if d.get("pcos_feature2"):
            parts.append("multiple small sized follicles arranged peripherally")
        if d.get("pcos_feature3"):
            parts.append("centrally echogenic stroma")
    if len(parts) == 2:
        return " & ".join(parts)
    return parts[0] if parts else ""


def _pcos_feature_phrase_up(d):
    return _pcos_feature_phrase_body(d).upper()


def _pcos_feature_count(d):
    count = 0
    pattern = _pcos_volume_pattern(d)
    if pattern in ("both_bulky", "right_bulky", "left_bulky"):
        count += 1
    if d.get("pcos_feature4"):
        count += 1
    else:
        if d.get("pcos_feature2"):
            count += 1
        if d.get("pcos_feature3"):
            count += 1
    return count


def _pcos_outcome(d):
    if d.get("pcos_feature4"):
        return PCOS_OUTCOME_PARTIAL
    no_subs = (not d.get("pcos_feature2")
               and not d.get("pcos_feature3")
               and not d.get("pcos_feature4"))
    if no_subs:
        return PCOS_OUTCOME_CONCLUSIVE
    count = _pcos_feature_count(d)
    if count >= 3:
        return PCOS_OUTCOME_CONCLUSIVE
    if count == 2:
        return PCOS_OUTCOME_PARTIAL
    return PCOS_OUTCOME_NOT_CONCLUSIVE


def _pcos_body_sentence(d):
    if not d.get("pcos"):
        return []
    pattern = _pcos_volume_pattern(d)
    feature_phrase = _pcos_feature_phrase_body(d)
    if not feature_phrase:
        feature_phrase = ("multiple small sized follicles arranged peripherally "
                          "& centrally echogenic stroma")
    if feature_phrase and not feature_phrase.endswith("."):
        feature_phrase = feature_phrase + " bilaterally"

    if pattern == "both_bulky":
        return [seg("Both ovaries appear "), seg("bulky in size", True),
                seg(f" with {feature_phrase}.")]
    if pattern == "right_bulky":
        return [seg("Right ovary appears "), seg("bulky in size", True),
                seg(f" while left ovary is normal, with {feature_phrase}.")]
    if pattern == "left_bulky":
        return [seg("Left ovary appears "), seg("bulky in size", True),
                seg(f" while right ovary is normal, with {feature_phrase}.")]
    return [seg("Both ovaries appear normal in size, however with "
                f"{feature_phrase}.")]


def _pcos_impression_line(d):
    if not d.get("pcos"):
        return None
    pattern = _pcos_volume_pattern(d)
    feature_phrase = _pcos_feature_phrase_up(d)
    if not feature_phrase:
        feature_phrase = ("MULTIPLE SMALL SIZED FOLLICLES ARRANGED PERIPHERALLY "
                          "& CENTRALLY ECHOGENIC STROMA")
    feature_phrase = feature_phrase + " BILATERALLY"

    if pattern == "both_bulky":
        head = f"BULKY OVARIES(>10CC) WITH {feature_phrase}"
    elif pattern == "right_bulky":
        head = (f"BULKY RIGHT OVARY(>10CC) WHILE LEFT OVARY IS NORMAL, "
                f"WITH {feature_phrase}")
    elif pattern == "left_bulky":
        head = (f"BULKY LEFT OVARY(>10CC) WHILE RIGHT OVARY IS NORMAL, "
                f"WITH {feature_phrase}")
    else:
        head = f"NORMAL OVARIAN VOLUME(<10CC), HOWEVER WITH {feature_phrase}"

    outcome = _pcos_outcome(d)
    tail = PCOS_OUTCOME_PHRASES[outcome]
    return f"{head}. {tail}"


def ovaries_sentence(d, age_years=None, uterus_operated=False):
    s = []
    r = d["right"]
    l = d["left"]
    r_present = _ovary_present(r)
    l_present = _ovary_present(l)

    postmeno = (age_years is not None and age_years >= 48)

    if not r_present and not l_present:
        if postmeno and not uterus_operated:
            return [seg("BOTH OVARIES", True, True),
                    seg(" are not visualized(likely atrophic).")]
        return [seg("BOTH OVARIES", True, True), seg(" are not visualized.")]

    pcos_body = _pcos_body_sentence(d) if d.get("pcos") else []

    if r_present and l_present:
        both_r_findings = r.get("findings", [])
        both_l_findings = l.get("findings", [])
        size_bracket = _ovary_size_bracket(r, l)
        if pcos_body:
            s.extend(pcos_body)
        else:
            s += [seg("BOTH OVARIES", True, True),
                  seg(" appears normal in size and echo pattern.")]
            if size_bracket:
                s += [seg(" "), seg(size_bracket, True)]
        for side_key, findings in (("right", both_r_findings),
                                    ("left", both_l_findings)):
            for f in findings:
                phrase = _ovary_finding_phrase(f, side_key, up=False)
                if phrase:
                    s += [seg(" "), seg(phrase, True)]
        return s

    if r_present:
        s += _ovary_single_side_block(r, "right", pcos_body)
    if l_present:
        s += _ovary_single_side_block(l, "left", pcos_body)
    return s


def _ovary_size_bracket(r, l):
    r_size = r.get("size_text", "").strip()
    l_size = l.get("size_text", "").strip()
    r_vol = str(r.get("volume_cc", "") or "").strip()
    l_vol = str(l.get("volume_cc", "") or "").strip()
    parts = []
    if r_size or r_vol:
        chunk = f"RO={r_size}mm" if r_size else "RO=___"
        if r_vol:
            chunk += f", VOL={r_vol}cc"
        parts.append(chunk)
    if l_size or l_vol:
        chunk = f"LO={l_size}mm" if l_size else "LO=___"
        if l_vol:
            chunk += f", VOL={l_vol}cc"
        parts.append(chunk)
    if not parts:
        return ""
    return "[" + " & ".join(parts) + "]."


def _ovary_single_side_block(side, side_key, pcos_body):
    out = []
    if pcos_body:
        out.extend(pcos_body)
    else:
        out += [seg(f"{side_key.upper()} OVARY", True, True),
                seg(" appears normal in size and echo pattern.")]
    for f in side.get("findings", []):
        phrase = _ovary_finding_phrase(f, side_key, up=False)
        if phrase:
            out += [seg(" "), seg(phrase, True)]
    return out


# ============================================================
# BOWEL / APPENDIX
# ============================================================

def bowel_sentence(d, sex, pancreas_status="normal"):
    acute = (pancreas_status == "acute")
    s = []
    if sex == "F":
        if not d["wall_thickening"]:
            s.append(seg("No obvious bowel wall thickening or lymphadenitis "
                         "appreciated.", True))
        else:
            s.append(seg("Bowel wall thickening seen.", True))
        if not acute:
            s.append(seg(" "))
            if d["free_fluid"] == "none":
                s.append(seg("No free fluid seen in the peritoneal cavity."))
            else:
                s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free "
                             f"fluid seen.", True))
    else:
        if not acute:
            if d["free_fluid"] == "none":
                s.append(seg("No free fluid is seen in the peritoneal cavity."))
            else:
                s.append(seg(f"{d['free_fluid'].replace('_', ' ').title()} free "
                             f"fluid seen.", True))
            s.append(seg(" "))
        if not d["wall_thickening"]:
            s.append(seg("No obvious bowel wall thickening or lymphadenitis "
                         "appreciated.", True))
        else:
            s.append(seg("Bowel wall thickening seen.", True))
    if d["mesenteric_ln"] == "present":
        cat = {"sad_lt_7": "Small(SAD<7mm)", "sad_gt_7": "Enlarged(SAD>7mm)",
               "sad_gt_10": "Enlarged(SAD>10mm)"}.get(d["ln_size_category"],
                                                       "Mesenteric")
        loc = d["ln_location"].replace("_", " ").title()
        largest = (f" with largest of these measuring {d['ln_largest']}"
                   if d["ln_largest"] else "")
        s += [seg(" "), seg(f"{cat} mesenteric lymph nodes are seen in the "
                             f"{loc} region{largest}", True), seg(".", True)]
    pe = d["pleural_effusion"]
    if pe != "none":
        text = {"trace_right": "Trace right pleural effusion is seen",
                "trace_left": "Trace left pleural effusion is seen",
                "mild_right": "Mild right pleural effusion is seen",
                "mild_bilateral": "Trace left & mild right pleural effusion seen"
                }.get(pe, "")
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
            s.append(seg(f" ({d['diameter_mm']}mm in diameter)", True))
        s.append(seg(".", True))
    elif d["status"] == "dilated":
        s += [seg("Appendix is dilated upto", True),
              seg(f" {d['diameter_mm']}mm", True),
              seg(" - ?Evolving appendicitis vs physiological.", True)]
    return s


# ============================================================
# IMPRESSION — helpers
# ============================================================

def _ureter_impression_term(level):
    return URETER_LEVELS[level][1]


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
    return f"A {side_up} {adj} CALCULUS({sz}MM){_ureter_impression_consequence(uc, side_up)}"


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
    return f"A COUPLE OF CALCULI({a}MM & {b}MM) {place}{_ureter_impression_consequence(uc, side_up)}"


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
    return f"{word} CALCULI, LARGEST MEASURING {sz}MM, {place}{_ureter_impression_consequence(uc, side_up)}"


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
    return _renal_calculi_clause(k, side_key)


def _cyst_impression_clause(k, side_key, small_prefix=False):
    ctype = k.get("cyst_type", "none")
    if ctype == "none":
        return None
    side_up = _ureter_side_label(side_key)
    count = k.get("cyst_count", "single")
    lead = _cyst_impression_count_word(count)
    bosniak = _cyst_bosniak(k)
    if ctype == "simple":
        core = (f"{lead} SIMPLE {side_up} RENAL CORTICAL CYST"
                if count == "single"
                else f"{lead} SIMPLE {side_up} RENAL CORTICAL CYSTS")
    else:
        core = (f"{lead} COMPLEX {side_up} RENAL CORTICAL CYST"
                if count == "single"
                else f"{lead} COMPLEX {side_up} RENAL CORTICAL CYSTS")
    bosniak_txt = f" (BOSNIAK CAT-{bosniak})"
    if small_prefix:
        core_no_article = core
        for art in ("A ", "FEW ", "MULTIPLE "):
            if core_no_article.startswith(art):
                core_no_article = core_no_article[len(art):]
                break
        return f"RELATIVELY SMALL {side_up} KIDNEY WITH {core_no_article}{bosniak_txt}"
    return core + bosniak_txt


def _spleen_impression_line(spleen):
    desc = spleen["size_descriptor"]
    spleen_focal = spleen.get("focal_lesion", "none")
    spleen_count = spleen.get("focal_count", "few")
    if desc == "enlarged_for_age":
        spleno_term = "SPLENOMEGALY FOR AGE"
        is_enlarged = True
    elif desc in SPLEEN_IMPRESSION_LABELS:
        size_str = spleen.get("size_mm", "")
        if size_str:
            spleno_term = f"{SPLEEN_IMPRESSION_LABELS[desc]}({size_str}MM)"
        else:
            spleno_term = SPLEEN_IMPRESSION_LABELS[desc]
        is_enlarged = True
    else:
        spleno_term = ""
        is_enlarged = False

    out = []
    if is_enlarged:
        pv_mm = spleen.get("portal_vein_mm", "")
        pv_class = classify_portal_vein(pv_mm) if pv_mm else "unknown"
        pv_paren = f"({pv_mm}MM)" if pv_mm else ""
        if pv_class == "prominent":
            pv_clause = f"PORTAL VEIN PROMINENT IN CALIBER{pv_paren}"
        elif pv_class == "dilated":
            pv_clause = f"PORTAL VEIN DILATED IN CALIBER{pv_paren}"
        else:
            pv_clause = f"NORMAL CALIBER PORTAL VEIN{pv_paren}"

        focal_clause = None
        if spleen_focal == "hyperechoic_foci":
            if spleen_count == "multiple":
                focal_clause = ("MULTIPLE HYPERECHOIC FOCI SCATTERED ACROSS "
                                "SPLENIC PARENCHYMA")
            else:
                focal_clause = ("FEW HYPERECHOIC FOCI SCATTERED ACROSS SPLENIC "
                                "PARENCHYMA")
        elif spleen_focal == "hypoechoic_foci":
            if spleen_count == "multiple":
                focal_clause = ("MULTIPLE HYPOECHOIC FOCI SCATTERED ACROSS "
                                "SPLENIC PARENCHYMA")
            else:
                focal_clause = ("FEW HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
                                "PARENCHYMA")

        pv_htn = pv_class in ("prominent", "dilated")
        has_foci = focal_clause is not None
        if pv_htn and has_foci:
            tail = (" - FEATURES SUGGESTIVE OF PORTAL HYPERTENSION WITH "
                    "?OLD GRANULOMATOUS ETIOLOGY.")
        elif pv_htn:
            tail = " - ?PORTAL HYPERTENSION."
        elif has_foci:
            tail = " - ?OLD GRANULOMATOUS ETIOLOGY."
        else:
            tail = "."

        if has_foci:
            return [f"{spleno_term} WITH {pv_clause} & {focal_clause}{tail}"]
        else:
            return [f"{spleno_term} WITH {pv_clause}{tail}"]
    else:
        if spleen_focal == "hyperechoic_foci":
            out.append("MULTIPLE HYPERECHOIC FOCI SCATTERED ACROSS SPLENIC "
                       "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY."
                       if spleen_count == "multiple"
                       else "FEW HYPERECHOIC FOCI SCATTERED ACROSS SPLENIC "
                            "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY.")
        elif spleen_focal == "hypoechoic_foci":
            out.append("MULTIPLE HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
                       "PARENCHYMA - ?OLD GRANULOMATOUS ETIOLOGY."
                       if spleen_count == "multiple"
                       else "FEW HYPOECHOIC FOCI SCATTERED ACROSS SPLENIC "
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

    echo = liver.get("echotexture", "normal")
    grade = liver.get("steatosis_grade") or ""
    echo_descriptor = None
    if echo == "coarse":
        echo_descriptor = "COARSE HEPATIC ECHOTEXTURE"
    elif echo == "raised":
        echo_descriptor = "RAISED HEPATIC ECHOTEXTURE"
    elif echo == "low":
        echo_descriptor = "LOW HEPATIC ECHOTEXTURE"
    elif echo == "increased":
        if grade == "Severe+++":
            echo_descriptor = ("SIGNIFICANT FATTY INFILTRATION - "
                               "?NON-ALCOHOLIC STEATO-HEPATITIS")
        elif grade:
            echo_descriptor = f"HEPATIC STEATOSIS({grade.upper()})"
        else:
            echo_descriptor = "HEPATIC STEATOSIS"

    lf = []
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
                lf.append(f"AN IRREGULAR MARGINATED ILL-DEFINED AVASCULAR SOL IN "
                          f"THE LIVER MEASURING VOL= {l0['vol']}CC IN SEGMENT "
                          f"{l0['segment']} - LIKELY LIVER ABSCESS")
            else:
                parts = [f"VOL= {l['vol']}CC IN SEGMENT {l['segment']}"
                         for l in lesions]
                keyword = "FEW" if count == "few" else "MULTIPLE"
                lf.append(f"{keyword} IRREGULAR MARGINATED ILL-DEFINED AVASCULAR "
                          f"SOLS IN THE LIVER, LARGEST OF THESE MEASURING "
                          f"{' & '.join(parts)} - LIKELY LIVER ABSCESSES")

    out = []
    if hepatomegaly and echo_descriptor and not lf:
        out.append(f"{hepatomegaly} WITH {echo_descriptor}. Adv- LFT Correlation.")
        return out
    if hepatomegaly and echo_descriptor and lf:
        out.append(f"{hepatomegaly} WITH {echo_descriptor} & "
                   + " AND ".join(lf) + ". Adv- LFT Correlation.")
        return out
    if not hepatomegaly and echo_descriptor and not lf:
        out.append(f"{echo_descriptor}. Adv- LFT Correlation.")
        return out
    if not hepatomegaly and echo_descriptor and lf:
        out.append(f"{echo_descriptor} & " + " AND ".join(lf)
                   + ". Adv- LFT Correlation.")
        return out
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
        liver_extras.append(f"HEPATIC STEATOSIS ({grade.upper()})"
                            if grade else "HEPATIC STEATOSIS")
    pv_mm = spleen.get("portal_vein_mm", "")
    pv_class = classify_portal_vein(pv_mm) if pv_mm else "unknown"
    pv_clause = None
    if pv_class == "normal":
        pv_clause = f"NORMAL CALIBER PORTAL VEIN ({pv_mm}MM)"
    elif pv_class == "prominent":
        pv_clause = f"PORTAL VEIN PROMINENT IN CALIBER ({pv_mm}MM)"
    elif pv_class == "dilated":
        return None
    parts = []
    if liver_extras:
        parts.append(" WITH " + " AND ".join(liver_extras))
    if pv_clause:
        parts.append(" & " + pv_clause)
    return base + "".join(parts) + ". Adv- LFT Correlation."


def _echogenicity_impression_lines(k):
    grade = k.get("cortical_echogenicity", "normal")
    if grade == "normal":
        return []
    lat = k.get("cortical_echogenicity_laterality", "bilateral")
    lat_up = {"bilateral": "BILATERAL", "right": "RIGHT",
              "left": "LEFT"}.get(lat, "BILATERAL")
    cmd = k.get("cortical_cmd", "preserved")
    grade_up = ECHO_GRADE_WORD_UP[grade]
    age_related = bool(k.get("age_related_echogenicity"))
    if grade == "mildly_raised":
        if age_related:
            return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY - "
                    f"?AGE RELATED. Adv- KFT Correlation."]
        return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY WITH "
                f"PRESERVED CORTICOMEDULLARY DIFFERENTIATION. "
                f"Adv- KFT Correlation."]
    if grade == "moderately_raised":
        if cmd == "preserved":
            return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY, THE "
                    f"CORTICOMEDULLARY DIFFERENTIATION IS HOWEVER PRESERVED - "
                    f"?MEDICAL RENAL DISEASE GRADE-II vs AKI. "
                    f"Adv- KFT Correlation."]
        return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY WITH HAZY "
                f"CORTICOMEDULLARY DIFFERENTIATION - ?MEDICAL RENAL DISEASE "
                f"GRADE-II/III vs AKI. Adv- KFT Correlation."]
    if cmd == "hazy":
        return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY WITH HAZY "
                f"CORTICOMEDULLARY DIFFERENTIATION - ?MEDICAL RENAL DISEASE "
                f"GRADE-II/III vs AKI. Adv- KFT Correlation."]
    return [f"{grade_up} {lat_up} RENAL CORTICAL ECHOGENICITY WITH LOST "
            f"CORTICOMEDULLARY DIFFERENTIATION - ?MEDICAL RENAL DISEASE "
            f"GRADE-III/IV vs AKI. Adv- KFT Correlation."]


# ============================================================
# IMPRESSION GENERATOR
# ============================================================

def generate_impression(d, sex, age):
    lines = []
    age_years = parse_age(age)
    is_pediatric = age_years is not None and age_years < 18
    small_rule_active = age_years is not None and age_years > SMALL_KIDNEY_MIN_AGE

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
                lines.append("MILD PERI-SPLENIC FAT STRANDING & FREE FLUID."
                             + tail())
            elif fat_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FAT STRANDING & FREE FLUID "
                             f"AROUND THE {loc_full[fat_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FAT STRANDING & FREE FLUID."
                             + tail())
        elif fat:
            if fat_loc == "perisplenic":
                lines.append("MILD PERI-SPLENIC FAT STRANDING." + tail())
            elif fat_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FAT STRANDING AROUND "
                             f"THE {loc_full[fat_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FAT STRANDING." + tail())
        elif fluid:
            if fluid_loc == "perisplenic":
                lines.append("MILD FREE FLUID IN THE PERI-SPLENIC REGION."
                             + tail())
            elif fluid_loc != "none":
                lines.append(f"MILD PERI-PANCREATIC FREE FLUID AROUND THE "
                             f"{loc_full[fluid_loc]}." + tail())
            else:
                lines.append("MILD PERI-PANCREATIC FREE FLUID." + tail())
        if ee_size == "mildly_bulky":
            if lines and (fat or fluid):
                last = lines.pop()
                lines.append("MILDLY BULKY PANCREAS WITH " + last)
            else:
                lines.append("MILDLY BULKY PANCREAS, HOWEVER NO PERI-PANCREATIC "
                             "FAT STRANDING OR FREE FLUID." + tail())

    elif p_status == "acute":
        ac_size = p.get("ac_size", "normal")
        ac_echo = p.get("ac_echo", "normal")
        ac_margins = p.get("ac_margins", "normal")
        if ac_size == "bulky":
            lines.append("MILD ASCITES WITH FEATURES SUGGESTIVE OF ACUTE "
                         "EDEMATOUS PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif ac_echo == "hypoechoic" or ac_margins == "irregular":
            lines.append("MILD ASCITES WITH FEATURES SUGGESTIVE OF ACUTE "
                         "NECROTIZING PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        else:
            lines.append("MILD ASCITES WITH MILD TO MODERATE PERI-PANCREATIC "
                         "FAT STRANDING AND MILD PERI-PANCREATIC FREE FLUID - "
                         "?ACUTE PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")

    elif p_status == "won_pseudocyst":
        wp_type = p.get("wp_type", "won")
        vol = p.get("wp_vol", "")
        loc = p.get("wp_location", "lesser_sac")
        loc_caps = ("IN THE LESSER SAC" if loc == "lesser_sac"
                    else "OVERLYING THE BODY OF THE PANCREAS")
        if wp_type == "won":
            lines.append(
                f"IRREGULARLY DEFINED PANCREAS OBSCURED BY A THICK-WALLED "
                f"COLLECTION {loc_caps} (VOL= {vol}CC) WITH DEBRIS WITHIN AND "
                f"SIGNIFICANT PERIPANCREATIC FAT STRANDING AND MILD FREE FLUID "
                f"- LIKELY WALLED-OFF NECROSIS (WON) AS A SEQUELAE TO ACUTE "
                f"PANCREATITIS. Adv- S.Amylase/Lipase Correlation.")
        elif wp_type == "pseudocyst":
            lines.append(
                f"HYPOTROPHIED PANCREAS WITH IRREGULAR MARGINS WITH A LOCULATED "
                f"COLLECTION (VOL= {vol}CC) WITH CLEAR CONTENTS {loc_caps} AND "
                f"MILD PERIPANCREATIC FAT STRANDING WITH MILD FREE FLUID - "
                f"LIKELY PANCREATIC PSEUDOCYST AS A SEQUELAE TO ACUTE "
                f"PANCREATITIS. Adv- S.Amylase/Lipase Correlation.")
        else:
            lines.append(
                f"HYPOTROPHIED PANCREAS WITH IRREGULAR MARGINS & A LOCULATED "
                f"MILDLY THICK WALLED COLLECTION (VOL= {vol}CC) WITH INTERNAL "
                f"ECHOES {loc_caps} AND MILD PERIPANCREATIC FAT STRANDING WITH "
                f"MILD FREE FLUID - LIKELY WON/PANCREATIC PSEUDOCYST AS A "
                f"SEQUELAE TO ACUTE PANCREATITIS. "
                f"Adv- S.Amylase/Lipase Correlation.")

    elif p_status == "chronic":
        foci = p.get("ch_foci", False)
        mpd = p.get("ch_mpd", False)
        fat = p.get("ch_fat", False)
        if foci and mpd and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC "
                         "PANCREATITIS. Adv- S.Amylase/Lipase Correlation.")
        elif foci and mpd:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif foci and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC "
                         "PANCREATITIS. Adv- S.Amylase/Lipase Correlation.")
        elif mpd and fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS "
                         "/ SEQUELAE TO ACUTE PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif foci:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif mpd:
            lines.append("FEATURES SUGGESTIVE OF CHRONIC PANCREATITIS. "
                         "Adv- S.Amylase/Lipase Correlation.")
        elif fat:
            lines.append("FEATURES SUGGESTIVE OF ?ACUTE ON CHRONIC PANCREATITIS "
                         "/ SEQUELAE TO ACUTE PANCREATITIS. "
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
        lines.append(f"CHOLELITHIASIS WITH CHOLEDOCHOLITHIASIS AND DILATED CBD "
                     f"UPTO {cbd_size}MM" + ihbr_clause() +
                     " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and gb_has_calculi and not cbd_dilated:
        lines.append("CHOLELITHIASIS WITH CHOLEDOCHOLITHIASIS AND NORMAL "
                     "CALIBER CBD" + ihbr_clause() +
                     " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and cbd_dilated:
        lines.append(f"CHOLEDOCHOLITHIASIS WITH DILATED CBD UPTO {cbd_size}MM"
                     + ihbr_clause() + " Adv- MRCP/CECT Abdomen Correlation.")
    elif cbd_has_calc and not cbd_dilated:
        lines.append("CHOLEDOCHOLITHIASIS WITH NORMAL CALIBER CBD"
                     + ihbr_clause() + " Adv- MRCP/CECT Abdomen Correlation.")

    gb_calc_already_reported = cbd_has_calc and gb_has_calculi and cbd_dilated
    if not gb_calc_already_reported:
        if gb_over and gb_has_calculi and (gb_wall_thickened or gb_peri_fluid):
            lines.append("OVERDISTENDED GALL BLADDER WITH CHOLELITHIASIS & "
                         "FEATURES SUGGESTIVE OF ACUTE CHOLECYSTITIS.")
        elif gb_over and gb_has_calculi:
            lines.append("OVERDISTENDED GALL BLADDER WITH CHOLELITHIASIS. "
                         "HOWEVER NO PERICHOLECYSTIC FLUID OR GB WALL "
                         "THICKENING APPRECIATED.")
        elif gb_contracted and gb_has_calculi:
            lines.append("CHOLELITHIASIS WITH ?CHRONIC CHOLECYSTITIS.")
        elif gb_has_calculi and (gb_wall_thickened or gb_peri_fluid):
            lines.append("CHOLELITHIASIS WITH ?ACUTE CHOLECYSTITIS.")
        elif gb_has_calculi:
            lines.append("CHOLELITHIASIS WITH NO APPRECIABLE PERICHOLECYSTIC "
                         "FLUID OR GB WALL THICKENING.")
        elif gb_wall_thickened and gb_peri_fluid:
            lines.append(f"GB WALL THICKENING {wall_descriptor}, SEEN UP TO "
                         f"{gb['wall_mm']}MM, WITH THIN RIM OF PERICHOLECYSTIC "
                         f"FLUID - ?ACALCULUS CHOLECYSTITIS. "
                         f"Adv- LFT, Lab & Clinical Correlation.")
        elif gb_wall_thickened and not gb_has_calculi:
            lines.append("ISOLATED GB WALL THICKENING WITHOUT ANY CALCULUS - "
                         "?ACALCULUS CHOLECYSTITIS. "
                         "Adv- LFT, Lab and Clinical Correlation.")
        elif gb_peri_fluid and not gb_has_calculi:
            lines.append("THIN RIM OF PERICHOLECYSTIC FLUID - "
                         "?SIGNIFICANCE. Adv- LFT, Lab and Clinical Correlation.")
        elif gb_sludge and not gb_has_calculi:
            lines.append(f"{gb['sludge'].upper()} SLUDGE SEEN IN THE GALLBLADDER "
                         f"LUMEN. Adv- Review scan after a month.")
        elif gb_sludge_ball and not gb_has_calculi:
            count = gb.get("sludge_ball_count", "single")
            wall = gb.get("sludge_ball_wall", "anterior").upper()
            if count == "single":
                lines.append(f"A SLUDGE BALL/POLYP SEEN IMPACTED AT THE {wall} GB "
                             f"WALL. Adv- Review scan after a month.")
            else:
                word = "FEW" if count == "few" else "MULTIPLE"
                lines.append(f"{word} SLUDGE BALLS/POLYPS SEEN IMPACTED AT THE "
                             f"{wall} GB WALL. Adv- Review scan after a month.")

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
        lines.extend(_liver_impression_line(liver))
        lines.extend(_spleen_impression_line(spleen))

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
            lines.append(f"EMPTY {side_key.upper()} RENAL FOSSA - ECTOPIC "
                         f"KIDNEY LYING IN {side_key.upper()} PELVIS.")

    r_uc = r["ureter_calculus"]
    l_uc = l["ureter_calculus"]
    bilateral_flag = k.get("bilateral_ureter_calculi", False)
    renal_lines = []

    cyst_clauses = []
    for side_key, side in (("right", r), ("left", l)):
        if _kidney_side_is_present(side):
            small = (small_rule_active
                     and parse_kidney_length_mm(side.get("size_text", "")) is not None
                     and parse_kidney_length_mm(side.get("size_text", "")) <= SMALL_KIDNEY_MM)
            clause = _cyst_impression_clause(
                side, side_key,
                small_prefix=small and side.get("cyst_type", "none") != "none")
            if clause:
                cyst_clauses.append(clause)

    if small_rule_active:
        r_len = parse_kidney_length_mm(r.get("size_text", "")) if r_present else None
        l_len = parse_kidney_length_mm(l.get("size_text", "")) if l_present else None
        r_small = r_len is not None and r_len <= SMALL_KIDNEY_MM
        l_small = l_len is not None and l_len <= SMALL_KIDNEY_MM
        echo_raised_any = k.get("cortical_echogenicity", "normal") != "normal"
        r_other = _kidney_side_has_other_finding(r) or echo_raised_any
        l_other = _kidney_side_has_other_finding(l) or echo_raised_any
        if r_present and l_present and r_small and l_small and not r_other and not l_other:
            renal_lines.append("BILATERALLY SMALL/CONTRACTED KIDNEYS - "
                               "?MEDICAL RENAL DISEASE. Adv- KFT Correlation.")
        elif r_present and l_present and r_small and not l_small and not r_other:
            renal_lines.append("SMALL/CONTRACTED RIGHT KIDNEY - ?SIGNIFICANCE. "
                               "Adv- KFT Correlation.")
        elif r_present and l_present and l_small and not r_small and not l_other:
            renal_lines.append("SMALL/CONTRACTED LEFT KIDNEY - ?SIGNIFICANCE. "
                               "Adv- KFT Correlation.")

    if bilateral_flag and r_uc["count"] != "none" and l_uc["count"] != "none":
        def side_desc(side_key, uc):
            rt_lt = "Rt" if side_key == "right" else "Lt"
            _, _, _, _, bi = URETER_LEVELS[uc["level"]]
            return f"{rt_lt} {bi} = {' & '.join(f'{s}mm' for s in uc['sizes'] if s)}"

        def max_size(uc):
            nums = [_parse_mm_value(s) for s in uc["sizes"]]
            nums = [n for n in nums if n is not None]
            return max(nums) if nums else 0

        first, second = (("left", "right") if max_size(l_uc) > max_size(r_uc)
                          else ("right", "left"))
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
                r_renal = _renal_calculi_clause(r, "right") if r_present else None
                l_renal = _renal_calculi_clause(l, "left") if l_present else None
                renal_clause = None
                if r_renal and l_renal:
                    renal_clause = "BILATERAL RENAL CALCULI"
                elif r_renal:
                    renal_clause = r_renal
                elif l_renal:
                    renal_clause = l_renal
                if renal_clause:
                    line = line.rstrip(".") + " & " + renal_clause + "."
                renal_lines.append(line)

    has_ureter_line = any(
        ("URETERIC CALCULUS" in ln or "URETERIC CALCULI" in ln
         or "URETER CALCULUS" in ln or "RENAL PELVIS CALCULUS" in ln
         or "PELVI-URETERIC JUNCTION CALCULUS" in ln
         or "VESICO-URETERIC JUNCTION CALCULUS" in ln
         or "BILATERAL URETERIC CALCULI" in ln)
        for ln in renal_lines)
    if not has_ureter_line:
        r_calcs = r["calculi"] if r_present else []
        l_calcs = l["calculi"] if l_present else []
        no_hydro = (r["hydronephrosis"] == "none" and l["hydronephrosis"] == "none")
        suffix = (", HOWEVER NO HYDRONEPHROSIS SEEN AT THE TIME OF SCAN."
                  if no_hydro else ".")
        if r_calcs and l_calcs:
            renal_lines.append(f"BILATERAL RENAL CALCULI{suffix}")
        elif r_calcs:
            renal_lines.append(f"{_renal_calculi_standalone(r, 'right')}{suffix}")
        elif l_calcs:
            renal_lines.append(f"{_renal_calculi_standalone(l, 'left')}{suffix}")

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
        rpc = (" - ?RECENTLY PASSED CALCULUS"
               if side.get("recently_passed_calculus_suspected") else "")
        renal_lines.append(f"{g} {side_up} HYDROURETERONEPHROSIS IS PRESENT"
                           f"{extra}{rpc}.")

    echo_lines = _echogenicity_impression_lines(k)
    for ln in echo_lines:
        if small_rule_active:
            r_len = parse_kidney_length_mm(r.get("size_text", "")) if r_present else None
            l_len = parse_kidney_length_mm(l.get("size_text", "")) if l_present else None
            r_small = r_len is not None and r_len <= SMALL_KIDNEY_MM
            l_small = l_len is not None and l_len <= SMALL_KIDNEY_MM
            lat = k.get("cortical_echogenicity_laterality", "bilateral")
            if lat == "bilateral" and r_small and l_small:
                ln = "RELATIVELY SMALL BOTH KIDNEYS WITH " + ln
            elif lat == "right" and r_small:
                ln = "RELATIVELY SMALL RIGHT KIDNEY WITH " + ln
            elif lat == "left" and l_small:
                ln = "RELATIVELY SMALL LEFT KIDNEY WITH " + ln
        renal_lines.append(ln)

    for side_key, side in (("right", r), ("left", l)):
        if not _kidney_side_is_present(side):
            continue
        if not side.get("contralateral_normal_clause"):
            continue
        other = "left" if side_key == "right" else "right"
        if side["calculi"]:
            for i in range(len(renal_lines) - 1, -1, -1):
                if any(tok in renal_lines[i] for tok in
                       ("RENAL CALCULUS", "RENAL CALCULI", "NEPHROLITHIASIS")):
                    base = renal_lines[i].rstrip(".")
                    base = base.replace(", HOWEVER NO HYDRONEPHROSIS SEEN AT THE "
                                        "TIME OF SCAN", "")
                    base = base.rstrip("., ")
                    renal_lines[i] = (base + f", HOWEVER NO CALCULUS/HYDRONEPHROSIS "
                                             f"SEEN ON THE {other.upper()} KIDNEY AT "
                                             f"THE TIME OF SCAN.")
                    break

    if k.get("negative_renal_line"):
        renal_lines.append("NO EVIDENCE OF HYDRONEPHROTIC CHANGES/CALCULUS SEEN "
                           "AT THE TIME OF SCAN.")

    if renal_lines and cyst_clauses:
        joined_main = " ".join(renal_lines)
        if not joined_main.endswith("."):
            joined_main += "."
        cyst_phrase = " ".join(c + "." for c in cyst_clauses)
        lines.append(joined_main + " " + cyst_phrase)
    elif renal_lines:
        lines.extend(renal_lines)
    elif cyst_clauses:
        for c in cyst_clauses:
            lines.append(c + ".")

    ub = d["urinary_bladder"]
    pr = d["prostate"]

    prostate_line = _prostate_impression_line(pr)
    ub_lines = _ub_impression_lines(ub, prostate_line=prostate_line)
    for ln in ub_lines:
        lines.append(ln)

    if sex == "F":
        uterus_lines = _uterus_impression_lines(d["uterus"], age_years)
        for ln in uterus_lines:
            lines.append(ln)

        if d["uterus"].get("neg_rpoc"):
            lines.append("NO OBVIOUS EVIDENCE OF ANY INTRA/EXTRA-UTERINE "
                         "PREGNANCY OR RPOC APPRECIATED ON TAS AT THE TIME OF SCAN.")
        if d["uterus"].get("neg_sol"):
            lines.append("NO OBVIOUS EVIDENCE OF ANY FOCAL SOL, RPOC OR INCREASED "
                         "VASCULARITY APPRECIATED IN THE UTERUS ON TAS.")

        o = d["ovaries"]
        if age_years is not None and age_years < 48:
            pcos_line = _pcos_impression_line(o)
            if pcos_line:
                lines.append(pcos_line)
        cyst_lines = _ovary_grouped_impression_lines(o)
        for cl in cyst_lines:
            lines.append(cl)

    b = d["bowel"]
    if b["mesenteric_ln"] == "present":
        loc = b["ln_location"].replace("_", " ").upper()
        lines.append(f"MESENTERIC LYMPH NODES IN THE {loc} REGION - "
                     f"?SIGNIFICANCE. Adv- Lab & Clinical Correlation.")
    if p_status != "acute" and b["free_fluid"] not in ("none",):
        lines.append(f"{b['free_fluid'].replace('_', ' ').upper()} FREE FLUID SEEN.")
    if b["pleural_effusion"] != "none":
        lines.append("PLEURAL EFFUSION SEEN - ?ETIOLOGY.")

    if not lines:
        return (list(IMPRESSION_NORMAL_FEMALE) if (sex == "F" and not is_pediatric)
                else list(IMPRESSION_NORMAL_MALE))

    lines.append(IMPRESSION_REST_UNREMARKABLE)
    return [_q(l) for l in lines]


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


def _add_blank(doc, size_pt):
    """v3.1.1: empty paragraph at a specific font size so it occupies real
    vertical space. space_after / space_before remain 0."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run("")
    run.font.name = FONT_BODY
    run.font.size = Pt(size_pt)
    return p


def _add_half_line(doc):
    return _add_blank(doc, FONT_SIZE_HALF)


def _add_full_line(doc):
    return _add_blank(doc, FONT_SIZE_BODY)


def _add_one_and_half_line(doc):
    _add_full_line(doc)
    _add_half_line(doc)


_BOSNIAK_RE = re.compile(r"BOSNIAK\s+CAT-[IVX]+", re.IGNORECASE)
_VS_RE = re.compile(r"\bvs\b", re.IGNORECASE)


def _split_adv(text):
    idx = text.find("Adv-")
    if idx == -1:
        return text, ""
    return text[:idx].rstrip(), text[idx:]


def _tokenize_parentheticals(text):
    out, buf, depth = [], [], 0
    for ch in text:
        if ch == "(":
            if depth == 0 and buf:
                out.append(("".join(buf), False))
                buf = []
            depth += 1
            buf.append(ch)
        elif ch == ")":
            buf.append(ch)
            depth -= 1
            if depth == 0:
                out.append(("".join(buf), True))
                buf = []
        else:
            buf.append(ch)
    if buf:
        out.append(("".join(buf), depth > 0))
    return out


def _split_vs(text):
    parts, last = [], 0
    for m in _VS_RE.finditer(text):
        if m.start() > last:
            parts.append((text[last:m.start()], False))
        parts.append(("vs", True))
        last = m.end()
    if last < len(text):
        parts.append((text[last:], False))
    return parts


def _split_bosniak(text):
    parts, last = [], 0
    for m in _BOSNIAK_RE.finditer(text):
        if m.start() > last:
            parts.append((text[last:m.start()], False))
        parts.append((m.group(0).upper(), True))
        last = m.end()
    if last < len(text):
        parts.append((text[last:], False))
    return parts


def _add_impression_line_runs(para, line):
    main, adv = _split_adv(line)
    for chunk, in_parens in _tokenize_parentheticals(main):
        if in_parens:
            for sub, _is_b in _split_bosniak(chunk):
                _add_run(para, sub, bold=True, italic=True)
        else:
            for sub, is_vs in _split_vs(chunk):
                if is_vs:
                    _add_run(para, "vs", bold=True, italic=False)
                else:
                    _add_run(para, sub.upper(), bold=True, italic=False)
    if adv:
        _add_run(para, " " + adv, bold=True, italic=True)


def _render_segments(para, segs):
    for tup in segs:
        text = tup[0]
        bold = tup[1]
        underline = tup[2]
        italic_override = tup[3] if len(tup) > 3 else None
        if italic_override is None:
            italic = bold and not underline
        else:
            italic = italic_override
        if "\n" in text:
            parts = text.split("\n")
            for i, part in enumerate(parts):
                if i > 0:
                    para.add_run().add_break()
                _add_run(para, part, bold=bold, underline=underline, italic=italic)
        else:
            _add_run(para, text, bold=bold, underline=underline, italic=italic)


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
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # --- Patient demographic table ---
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
        _add_run(cell.paragraphs[0], text, bold=True)

    # 0.5 line below patient table
    _add_half_line(doc)

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_after = Pt(0)
    title_para.paragraph_format.space_before = Pt(0)
    _add_run(title_para, "ULTRASOUND WHOLE ABDOMEN", bold=True, underline=True,
             size=FONT_SIZE_TITLE, color=TITLE_COLOR)

    # 1.5 lines between title and first organ
    _add_one_and_half_line(doc)

    sex, age = p["sex"], p["age"]
    age_years = parse_age(age)
    is_ped = age_years is not None and age_years < 18
    is_ped_female = (sex == "F" and age_years is not None and age_years < 14)
    p_status = data["pancreas"].get("status", "normal")
    spleen_enlarged = data["spleen"]["size_descriptor"] != "normal"
    ub_status = data["urinary_bladder"].get("status") or "adequately_distended"

    sections = [
        liver_sentence(data["liver"], sex, age, spleen_enlarged=spleen_enlarged,
                       cbd_ihbr=data["cbd"].get("ihbr", "normal")),
        gall_bladder_sentence(data["gall_bladder"]),
        cbd_sentence(data["cbd"]),
        pancreas_sentence(data["pancreas"]),
        spleen_sentence(data["spleen"]),
        kidneys_sentence(data["kidneys"], ub_status=ub_status, age_text=age),
        urinary_bladder_sentence(data["urinary_bladder"]),
    ]
    if sex == "F":
        sections.append(uterus_sentence(data["uterus"], pediatric=is_ped_female,
                                         age_years=age_years))
        uterus_operated = data["uterus"].get("status") == "operated"
        sections.append(ovaries_sentence(data["ovaries"], age_years=age_years,
                                          uterus_operated=uterus_operated))
    else:
        sections.append(prostate_sentence(data["prostate"], pediatric=is_ped,
                                           ub_status=ub_status))
    sections.append(bowel_sentence(data["bowel"], sex, pancreas_status=p_status))
    if data["appendix"]["status"] != "not_assessed":
        sections.append(appendix_sentence(data["appendix"]))

    rendered_count = 0
    for segs in sections:
        if not segs:
            continue
        # 1 line between organs (before each subsequent organ)
        if rendered_count > 0:
            _add_full_line(doc)
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        _render_segments(para, segs)
        rendered_count += 1

    addendum = (data.get("additional_body_findings") or "").strip()
    if addendum:
        if rendered_count > 0:
            _add_full_line(doc)
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        for i, ln in enumerate(addendum.split("\n")):
            if i > 0:
                para.add_run().add_break()
            _add_run(para, ln, bold=True, italic=True)

    # 1.5 lines before IMPRESSION
    _add_one_and_half_line(doc)

    imp_head = doc.add_paragraph()
    imp_head.paragraph_format.space_after = Pt(0)
    imp_head.paragraph_format.space_before = Pt(0)
    _add_run(imp_head, "IMPRESSION:", bold=True, underline=True,
             size=FONT_SIZE_BODY, color=TITLE_COLOR)

    # 0.5 line between IMPRESSION heading and first bullet
    _add_half_line(doc)

    for line in data["impression"]["lines"]:
        para = doc.add_paragraph(style="List Bullet")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        _add_impression_line_runs(para, line)

    # 0.5 line between last bullet and disclaimer
    _add_half_line(doc)

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
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        display: flex !important; flex-direction: column !important;
        justify-content: flex-end !important;
    }
    @media (max-width: 900px) {
        div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            min-width: 0 !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            flex-wrap: wrap !important;
            gap: 0.4rem 1rem !important;
        }
        div[data-testid="stRadio"] label {
            min-height: 44px !important;
            padding: 6px 10px !important;
            display: flex !important;
            align-items: center !important;
            cursor: pointer !important;
            user-select: none !important;
        }
        div[data-testid="stRadio"] label > div:first-child {
            width: 22px !important; height: 22px !important;
            min-width: 22px !important;
            margin-right: 8px !important;
        }
        div[data-testid="stRadio"] label > div:first-child > div {
            width: 22px !important; height: 22px !important;
        }
        div[data-testid="stCheckbox"] label {
            min-height: 44px !important;
            padding: 6px 6px !important;
            cursor: pointer !important;
            user-select: none !important;
        }
    }
    div[data-testid="stRadio"] label,
    div[data-testid="stCheckbox"] label { cursor: pointer !important; }
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
        fill: #111111 !important; stroke: #111111 !important;
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
    liver_size = (str(int(liver_size_num))
                  if (liver_size_num and liver_size_num > 0) else "")
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
                "Echotexture",
                ["normal", "increased", "raised", "coarse", "low"],
                horizontal=True,
                format_func=lambda x: {"normal": "Normal",
                                       "increased": "Increased (steatosis)",
                                       "raised": "Raised",
                                       "coarse": "Coarse",
                                       "low": "Low"}[x],
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
                 "operated"],
                horizontal=True,
                format_func=lambda x: {"adequately_distended": "Adequate",
                                       "over": "Over-distended",
                                       "partially": "Partially contracted",
                                       "contracted": "Contracted",
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
                index=None, horizontal=True,
                format_func=lambda x: {
                    "early_evolving": "Early/evolving pancreatitis",
                    "acute": "Acute pancreatitis",
                    "won_pseudocyst": "WON / Pseudocyst",
                    "chronic": "Chronic pancreatitis"}[x],
                key="pn_status")
            _pn = ss("pn_status", None)
            if _pn == "early_evolving":
                st.radio("Pancreas size", ["normal", "mildly_bulky"],
                         horizontal=True,
                         format_func=lambda x: {"normal": "Normal size",
                                                "mildly_bulky": "Mildly bulky"}[x],
                         key="pn_ee_size")
                st.checkbox("Mild peri-pancreatic fat stranding", key="pn_ee_fat")
                if ss("pn_ee_fat", False):
                    st.radio("Fat stranding location",
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
                    st.radio("Free fluid location",
                             ["none", "head_neck", "body", "neck_body", "perisplenic"],
                             horizontal=True,
                             format_func=lambda x: {"none": "None",
                                                    "head_neck": "Head & neck",
                                                    "body": "Body",
                                                    "neck_body": "Neck & body",
                                                    "perisplenic": "Peri-splenic"}[x],
                             key="pn_ee_fluid_loc")
            elif _pn == "acute":
                st.radio("Pancreas size", ["normal", "bulky"], horizontal=True,
                         format_func=lambda x: {"normal": "Normal size",
                                                "bulky": "Bulky"}[x],
                         key="pn_ac_size")
                st.radio("Echotexture", ["normal", "hypoechoic"],
                         horizontal=True,
                         format_func=lambda x: {"normal": "Normal",
                                                "hypoechoic": "Hypoechoic heterogeneous"}[x],
                         key="pn_ac_echo")
                if ss("pn_ac_echo", "normal") == "hypoechoic":
                    st.radio("Echotexture location",
                             ["none", "head_neck", "body"], horizontal=True,
                             format_func=lambda x: {"none": "None (all)",
                                                    "head_neck": "Head & neck",
                                                    "body": "Body"}[x],
                             key="pn_ac_echo_loc")
                st.radio("Margins", ["normal", "irregular"], horizontal=True,
                         format_func=lambda x: {"normal": "Normal",
                                                "irregular": "Irregular / fuzzy"}[x],
                         key="pn_ac_margins")
                st.caption("Mild to moderate peri-pancreatic fat stranding and "
                           "mild free fluid are automatically included.")
            elif _pn == "won_pseudocyst":
                st.radio("Type",
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
                st.radio("Location",
                         ["lesser_sac", "overlying_body"], horizontal=True,
                         format_func=lambda x: {"lesser_sac": "In the lesser sac",
                                                "overlying_body": "Overlying the body of the pancreas"}[x],
                         key="pn_wp_location")
            elif _pn == "chronic":
                st.checkbox("Foci of calcification", key="pn_ch_foci")
                st.checkbox("MPD dilation", key="pn_ch_mpd")
                if ss("pn_ch_mpd", False):
                    st.text_input("MPD size (mm)", key="pn_ch_mpd_size")
                st.checkbox("Mild fat stranding", key="pn_ch_fat")

            # TODO(v3.x): peri-pancreatic + peri-portal lymph nodes.

    # ------- SPLEEN -------
    col_sp_main, col_sp_sz = st.columns([5, 1], vertical_alignment="bottom")
    with col_sp_main:
        spleen_open = organ_button("SPLEEN", "SPLEEN")
    with col_sp_sz:
        sp_size_num = st.number_input(
            "Spleen size mm", min_value=0, max_value=500, value=None, step=1,
            key="spleen_size_num", label_visibility="collapsed",
            placeholder="SPLEEN mm")
    sp_size = (str(int(sp_size_num))
               if (sp_size_num and sp_size_num > 0) else "")
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
                    st.radio("Location", ["hilum", "upper_pole", "lower_pole"],
                             horizontal=True,
                             format_func=lambda x: {"hilum": "Hilum",
                                                    "upper_pole": "Upper pole",
                                                    "lower_pole": "Lower pole"}[x],
                             key="sp_acc_loc")

    # ------- KIDNEYS -------
    kd_open = organ_button("KIDNEYS", "KIDNEYS")
    if kd_open:
        with st.container(border=True):
            st.radio(
                "Cortical echogenicity",
                ["normal", "mildly_raised", "moderately_raised",
                 "significantly_raised"],
                horizontal=True,
                format_func=lambda x: {
                    "normal": "Normal",
                    "mildly_raised": "Mildly raised",
                    "moderately_raised": "Moderately raised",
                    "significantly_raised": "Significantly raised"}[x],
                key="kd_cort_echo")
            _echo = ss("kd_cort_echo", "normal")
            if _echo in ("mildly_raised", "moderately_raised",
                         "significantly_raised"):
                c_a, c_b, c_c = st.columns([2, 2, 1])
                with c_a:
                    st.radio("Laterality", ["bilateral", "right", "left"],
                             horizontal=True, format_func=lambda x: x.title(),
                             key="kd_cort_lat")
                with c_b:
                    if _echo == "mildly_raised":
                        cmd_opts = ["preserved"]
                    elif _echo == "moderately_raised":
                        cmd_opts = ["preserved", "hazy"]
                    else:
                        cmd_opts = ["hazy", "lost"]
                    st.radio("Corticomedullary differentiation",
                             cmd_opts, horizontal=True,
                             format_func=lambda x: x.title(),
                             key="kd_cort_cmd")
                with c_c:
                    if _echo == "mildly_raised":
                        st.checkbox("?Age related", key="kd_age_related")
            st.radio(
                "Negative renal impression line (clinical query, no finding)",
                ["no", "yes"], horizontal=True, key="kd_neg_renal")

            ub_status_now = ss("ub_status", None) or "adequately_distended"

            for side in ("right", "left"):
                with st.expander(f"{side.title()} Kidney", expanded=False):
                    st.radio("Status",
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
                    _sz_txt = ss(f"kd_{side[0]}_size_text", "")
                    _len = parse_kidney_length_mm(_sz_txt)
                    _age_yr = parse_age(p_age)
                    if (_len is not None and _age_yr is not None
                            and _age_yr > SMALL_KIDNEY_MIN_AGE):
                        if _len <= SMALL_KIDNEY_MM:
                            st.warning(f"→ Relatively small/contracted "
                                       f"({_len}mm ≤ {SMALL_KIDNEY_MM}mm)")
                        else:
                            st.success(f"✓ Normal length ({_len}mm)")

                    st.checkbox("Renal calculi present",
                                key=f"kd_{side[0]}_has_calc")
                    calc_key = f"kd_{side[0]}_calc_count"
                    if ss(f"kd_{side[0]}_has_calc", False):
                        if calc_key not in st.session_state:
                            st.session_state[calc_key] = 1
                        calc_r_cols = st.columns([3, 1])
                        with calc_r_cols[0]:
                            st.caption(f"{side.title()}: "
                                       f"{st.session_state[calc_key]} calculi")
                        with calc_r_cols[1]:
                            if st.button("+ Add", key=f"kd_{side[0]}_add_calc"):
                                st.session_state[calc_key] += 1
                                st.rerun()
                        for i in range(st.session_state[calc_key]):
                            c_a, c_b, c_c = st.columns([1, 2, 1])
                            with c_a:
                                st.text_input("Size",
                                              key=f"kd_{side[0]}_calc_size_{i}",
                                              label_visibility="collapsed",
                                              placeholder="mm")
                            with c_b:
                                st.selectbox("Pole",
                                             ["upper", "upper_mid", "mid",
                                              "lower_mid", "lower"],
                                             key=f"kd_{side[0]}_calc_pole_{i}",
                                             label_visibility="collapsed",
                                             format_func=lambda x: _format_pole(x))
                            with c_c:
                                if st.button("✕", key=f"kd_{side[0]}_calc_del_{i}"):
                                    st.session_state[calc_key] -= 1
                                    st.rerun()
                    else:
                        st.session_state[calc_key] = 0

                    st.checkbox("Ureteric calculus present",
                                key=f"kd_{side[0]}_has_uc")
                    if ss(f"kd_{side[0]}_has_uc", False):
                        st.radio("Ureteric calculus count",
                                 ["single", "couple", "few", "multiple"],
                                 horizontal=True, key=f"kd_{side[0]}_uc_count",
                                 format_func=lambda x: x.title())
                        st.selectbox("Level",
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
                        st.radio("Grade of back-pressure",
                                 ["none", "no_significant", "minimal", "mild",
                                  "moderate"],
                                 horizontal=True, key=f"kd_{side[0]}_uc_grade",
                                 format_func=lambda x: {
                                     "none": "No back-pressure",
                                     "no_significant": "No significant",
                                     "minimal": "Minimal",
                                     "mild": "Mild",
                                     "moderate": "Moderate"}[x])

                    st.checkbox("Cyst present", key=f"kd_{side[0]}_has_cyst")
                    if ss(f"kd_{side[0]}_has_cyst", False):
                        st.radio("Cyst type", ["simple", "complex"],
                                 horizontal=True, key=f"kd_{side[0]}_cyst_type",
                                 format_func=lambda x: {
                                     "simple": "Simple cyst",
                                     "complex": "Complex cyst"}[x])
                        _ctype = ss(f"kd_{side[0]}_cyst_type", "simple")
                        c1, c2 = st.columns(2)
                        with c1:
                            st.text_input("Size (mm; largest if >1)",
                                          key=f"kd_{side[0]}_cyst_size")
                        with c2:
                            st.selectbox("Location",
                                         ["", "upper", "upper_mid", "mid",
                                          "lower_mid", "lower"],
                                         key=f"kd_{side[0]}_cyst_loc",
                                         format_func=lambda x: (
                                             "—" if x == "" else _format_pole(x)))
                        st.radio("Count", ["single", "few", "multiple"],
                                 horizontal=True, key=f"kd_{side[0]}_cyst_count",
                                 format_func=lambda x: x.title())
                        if _ctype == "complex":
                            st.radio("Septa", ["none", "thin", "thick"],
                                     horizontal=True,
                                     format_func=lambda x: x.title(),
                                     key=f"kd_{side[0]}_cyst_septa")
                            st.radio("Calcification",
                                     ["none", "arc", "nodular"],
                                     horizontal=True,
                                     format_func=lambda x: {
                                         "none": "None",
                                         "arc": "Mural arc-like",
                                         "nodular": "Mural nodular"}[x],
                                     key=f"kd_{side[0]}_cyst_calc")

                    st.checkbox("Standalone hydronephrosis",
                                key=f"kd_{side[0]}_has_hydro")
                    if ss(f"kd_{side[0]}_has_hydro", False):
                        st.radio("Grade", ["minimal", "mild", "moderate"],
                                 horizontal=True, key=f"kd_{side[0]}_hydro",
                                 format_func=lambda x: x.title())
                        st.checkbox("No obstructive calculus upto visualized "
                                    "distal ureter",
                                    key=f"kd_{side[0]}_hydro_nocalc")
                        st.checkbox("?Recently passed calculus",
                                    key=f"kd_{side[0]}_hydro_rpc")
                        if ub_status_now == "empty":
                            st.checkbox(f"{side.title()} distal ureter could "
                                        f"not be traced (UB is empty)",
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
            c1, c2, c3 = st.columns(3)
            with c1:
                st.checkbox("Partially empty", key="ub_partially_empty")
            with c2:
                st.checkbox("Empty", key="ub_empty")
            with c3:
                st.checkbox("Over-distended", key="ub_over")
            st.checkbox("Catheterized", key="ub_catheterized")

            st.text_input("Pre-void volume (CC)", key="ub_pre_void")

            st.checkbox("No mass / calculus seen (default checked)",
                        value=True, key="ub_no_mass")
            st.checkbox("Wall thickening present", key="ub_wall_check")
            if ss("ub_wall_check", False):
                c_a, c_b = st.columns(2)
                with c_a:
                    st.text_input("Wall thickness (mm)", key="ub_wall_mm")
                with c_b:
                    st.checkbox("Irregular wall", key="ub_wall_irregular")

            st.markdown("**Sedimentation**")
            st.checkbox("Trace sedimentation", key="ub_sed_trace")
            st.checkbox("Free floating sedimentation",
                        key="ub_sed_free_floating")
            st.checkbox("Settled debris", key="ub_sed_settled_debris")
            st.checkbox("Extensive sedimentation", key="ub_sed_extensive")
            if (ss("ub_sed_trace", False) or ss("ub_sed_free_floating", False)
                    or ss("ub_sed_settled_debris", False)
                    or ss("ub_sed_extensive", False)):
                st.checkbox("?UTI (append to advice)", key="ub_uti")

            if str(ss("ub_pre_void", "") or "").strip():
                st.text_input("Post-void residue (CC)", key="ub_post_void")

    # ------- UTERUS / OVARIES (F) or PROSTATE (M) -------
    if p_sex == "F":
        _age_years = parse_age(p_age)
        _is_ped_female = (_age_years is not None and _age_years < 14)

        ut_open = organ_button("UTERUS", "UTERUS")
        if ut_open:
            with st.container(border=True):
                st.radio("Status",
                         ["anteverted", "partially_visualized", "operated",
                          "not_visualized"],
                         horizontal=True,
                         format_func=lambda x: {
                             "anteverted": "Anteverted",
                             "partially_visualized": "Partially visualized",
                             "operated": "Operated",
                             "not_visualized": "Not visualized"}[x],
                         key="ut_status")
                _uts = ss("ut_status", "anteverted")

                if _uts == "operated":
                    st.caption("Ovaries auto-fill as not visualized. "
                               "Override below if needed.")

                if _uts in ("anteverted", "partially_visualized"):
                    st.checkbox("Retro-flexed", key="ut_retroflexed")
                    st.checkbox("Gravid uterus", key="ut_gravid")
                    if ss("ut_gravid", False):
                        c_a, c_b, c_c, c_d = st.columns([1, 1, 1, 1])
                        with c_a:
                            st.radio("Type", ["CRL", "GS"], horizontal=True,
                                     key="ut_gravid_type")
                        with c_b:
                            st.text_input("Length (mm)", key="ut_gravid_length")
                        with c_c:
                            st.text_input("GA weeks", key="ut_gravid_weeks")
                        with c_d:
                            st.text_input("GA days", key="ut_gravid_days")
                    st.checkbox("Low-lying uterus", key="ut_low_lying")
                    if ss("ut_low_lying", False):
                        st.radio("Cervix",
                                 ["partially", "not_visualized"],
                                 horizontal=True,
                                 format_func=lambda x: {
                                     "partially": "Partially visualized",
                                     "not_visualized": "Not visualized"}[x],
                                 key="ut_cervix_vis")

                if _uts in ("anteverted", "partially_visualized"):
                    st.text_input("Uterus size (mm) — e.g. 92x33",
                                  key="ut_size")

                    st.radio("Myometrium",
                             ["homogenous", "mildly_heterogeneous",
                              "heterogeneous"],
                             horizontal=True,
                             format_func=lambda x: {
                                 "homogenous": "Homogenous",
                                 "mildly_heterogeneous": "Mildly heterogenous",
                                 "heterogeneous": "Heterogenous"}[x],
                             key="ut_myometrium")
                    st.checkbox("Prominent vascular channels",
                                key="ut_prom_vasc")

                    st.checkbox("Adenomyosis features", key="ut_adenomyosis")
                    if ss("ut_adenomyosis", False):
                        st.caption("Tick features. ≥2 features → 'likely adenomyosis'; "
                                   "1 feature → '?early adenomyosis'.")
                        st.checkbox("Globular shape", key="ut_aden_globular")
                        st.checkbox("Asymmetrically bulky posterior myometrium",
                                    key="ut_aden_asymm")
                        st.checkbox("Venetian blind sign", key="ut_aden_venetian")
                        st.checkbox("Endo-myometrial interface indistinct",
                                    key="ut_aden_iface_indistinct")
                        st.checkbox("Interface barely perceptible at fundus",
                                    key="ut_aden_iface_fundus")
                        st.checkbox("Interface lost", key="ut_aden_iface_lost")
                        st.checkbox("Subendometrial cysts", key="ut_aden_cysts")
                        st.radio("Confidence", ["auto", "likely", "early"],
                                 horizontal=True, key="ut_aden_conf")

                    st.checkbox("Fibroid", key="ut_fibroid")
                    if ss("ut_fibroid", False):
                        st.radio("Count", ["single", "few", "multiple"],
                                 horizontal=True, key="ut_fibroid_count",
                                 format_func=lambda x: x.title())
                        st.caption("Tick all applicable types.")
                        st.checkbox("Intramural", key="ut_fib_type_intramural")
                        st.checkbox("Intramural > subserosal",
                                    key="ut_fib_type_intra_subser")
                        st.checkbox("Intramural > submucosal",
                                    key="ut_fib_type_intra_submuc")
                        st.checkbox("Submucosal > intramural",
                                    key="ut_fib_type_submuc_intra")
                        st.checkbox("Subserosal > intramural",
                                    key="ut_fib_type_subser_intra")
                        st.caption("Enter size and location per ticked type, "
                                   "largest first.")
                        _type_keys = [
                            ("ut_fib_type_intramural", "intramural", "Intramural"),
                            ("ut_fib_type_intra_subser", "intramural_subserosal",
                             "Intramural > subserosal"),
                            ("ut_fib_type_intra_submuc", "intramural_submucosal",
                             "Intramural > submucosal"),
                            ("ut_fib_type_submuc_intra", "submucosal_intramural",
                             "Submucosal > intramural"),
                            ("ut_fib_type_subser_intra", "subserosal_intramural",
                             "Subserosal > intramural"),
                        ]
                        for k, _, lbl in _type_keys:
                            if ss(k, False):
                                c_a, c_b = st.columns([1, 2])
                                with c_a:
                                    st.text_input(f"{lbl} size", key=f"{k}_size")
                                with c_b:
                                    st.selectbox(
                                        f"{lbl} location",
                                        ["fundal", "anterior", "posterior"],
                                        key=f"{k}_loc",
                                        format_func=lambda x: x.title())
                        _n_types = sum(1 for k, _, _ in _type_keys if ss(k, False))
                        if _n_types > 1:
                            st.text_input("FIGO range (manual, e.g. 3-5)",
                                          key="ut_fib_figo_manual")

                    st.checkbox("Adenomyomas", key="ut_adenomyomas")
                    if ss("ut_adenomyomas", False):
                        st.radio("Count",
                                 ["single", "couple", "few", "multiple"],
                                 horizontal=True, key="ut_adenoma_count",
                                 format_func=lambda x: x.title())
                        st.radio("Echo",
                                 ["hyperechoic", "hypoechoic", "hypo-isoechoic"],
                                 horizontal=True, key="ut_adenoma_echo",
                                 format_func=lambda x: x.title())
                        st.radio("Location", ["anterior", "posterior", "fundal"],
                                 horizontal=True, key="ut_adenoma_loc",
                                 format_func=lambda x: x.title())
                        if ss("ut_adenoma_count", "couple") == "single":
                            st.text_input("Size (mm, e.g. 15x12)",
                                          key="ut_adenoma_size")
                        st.checkbox("Avascular (default on)",
                                    value=True, key="ut_adenoma_avascular")

                    st.text_input("Endometrial thickness (mm)", key="ut_endo_mm")
                    st.radio("Endometrial collection",
                             ["none", "mild", "moderate"],
                             horizontal=True, key="ut_endo_coll",
                             format_func=lambda x: x.title())
                    if ss("ut_endo_coll", "none") in ("mild", "moderate"):
                        st.checkbox("Heterogeneous collection", key="ut_endo_het")

                    st.checkbox("RPOC structure", key="ut_rpoc")
                    if ss("ut_rpoc", False):
                        st.text_input("Size (mm, e.g. 15x12)", key="ut_rpoc_size")
                        st.radio("Echo", ["hypo", "hyper"], horizontal=True,
                                 format_func=lambda x: {
                                     "hypo": "Hypo-echoic",
                                     "hyper": "Hyper-echoic"}[x],
                                 key="ut_rpoc_echo")

                    st.checkbox("Elongated cervix", key="ut_cx_elong")
                    st.checkbox("Bulky cervix", key="ut_cx_bulky")
                    if ss("ut_cx_bulky", False):
                        st.text_input("Bulky cervix size (mm)",
                                      key="ut_cx_bulky_size")
                    st.radio("Nabothian cysts",
                             ["none", "one", "few", "multiple"],
                             horizontal=True, key="ut_nabothian",
                             format_func=lambda x: x.title())
                    st.radio("Cervix collection",
                             ["none", "mild", "moderate"],
                             horizontal=True, key="ut_cx_coll",
                             format_func=lambda x: x.title())

                    st.checkbox("Add: no RPOC/pregnancy line", key="ut_neg_rpoc")
                    st.checkbox("Add: no focal SOL/RPOC line", key="ut_neg_sol")

        # ------- OVARIES -------
        ov_open = organ_button("OVARIES", "OVARIES")
        if ov_open:
            with st.container(border=True):
                c1, c2 = st.columns(2)
                for col, side in ((c1, "right"), (c2, "left")):
                    with col:
                        st.markdown(f"**{side.title()} Ovary**")
                        _sd = side[0]
                        st.radio("Status",
                                 ["normal", "not_visualized"],
                                 horizontal=True,
                                 format_func=lambda x: {
                                     "normal": "Present",
                                     "not_visualized": "Not visualized"}[x],
                                 key=f"ov_{_sd}_status")
                        if ss(f"ov_{_sd}_status", "normal") == "normal":
                            st.text_input("Size (mm) — e.g. 42x26",
                                          key=f"ov_{_sd}_size")
                            st.text_input("Volume (cc)",
                                          key=f"ov_{_sd}_vol")
                            st.caption("Findings")
                            st.checkbox("Simple cyst",
                                        key=f"ov_{_sd}_f_simple")
                            if ss(f"ov_{_sd}_f_simple", False):
                                st.text_input("Simple cyst size",
                                              key=f"ov_{_sd}_f_simple_size")
                                st.radio("Count", ["single", "few", "multiple"],
                                         horizontal=True,
                                         key=f"ov_{_sd}_f_simple_count",
                                         format_func=lambda x: x.title())
                            st.checkbox("Hemorrhagic cyst",
                                        key=f"ov_{_sd}_f_hem")
                            if ss(f"ov_{_sd}_f_hem", False):
                                st.text_input("Hemorrhagic cyst size",
                                              key=f"ov_{_sd}_f_hem_size")
                                st.radio("Count", ["single", "few", "multiple"],
                                         horizontal=True,
                                         key=f"ov_{_sd}_f_hem_count",
                                         format_func=lambda x: x.title())
                            st.checkbox("Hemorrhagic follicle",
                                        key=f"ov_{_sd}_f_hemf")
                            if ss(f"ov_{_sd}_f_hemf", False):
                                st.text_input("Hemorrhagic follicle size",
                                              key=f"ov_{_sd}_f_hemf_size")
                                st.radio("Count", ["single", "few", "multiple"],
                                         horizontal=True,
                                         key=f"ov_{_sd}_f_hemf_count",
                                         format_func=lambda x: x.title())
                            st.checkbox("Follicular cyst",
                                        key=f"ov_{_sd}_f_foll")
                            if ss(f"ov_{_sd}_f_foll", False):
                                st.text_input("Follicular cyst size",
                                              key=f"ov_{_sd}_f_foll_size")
                                st.radio("Count", ["single", "few", "multiple"],
                                         horizontal=True,
                                         key=f"ov_{_sd}_f_foll_count",
                                         format_func=lambda x: x.title())

                st.markdown("---")
                st.markdown("**PCOS spectrum**")
                st.checkbox("PCOS features present", key="ov_pcos")
                if ss("ov_pcos", False):
                    st.caption("If no sub-ticks are chosen, all three features "
                               "are assumed present (conclusive).")
                    st.checkbox("Feature 2: multiple small follicles arranged peripherally",
                                key="ov_pcos_f2")
                    st.checkbox("Feature 3: centrally echogenic stroma",
                                key="ov_pcos_f3")
                    st.checkbox("Feature 4: variable-sized follicles (variant)",
                                key="ov_pcos_f4")
                    if ss("ov_pcos_f4", False):
                        st.caption("Variable-sized may combine with either "
                                   "peripheral or random; peripheral and random "
                                   "are mutually exclusive.")
                        st.checkbox("Variable size follicles",
                                    key="ov_pcos_f4_var")
                        st.radio("Distribution",
                                 ["peripheral", "random"],
                                 horizontal=True,
                                 format_func=lambda x: {
                                     "peripheral": "Predominantly peripheral",
                                     "random": "Random"}[x],
                                 key="ov_pcos_f4_dist")

                st.markdown("---")
                st.markdown("**Manual advices (ovary-related)**")
                st.checkbox("Adv- Follicular Monitoring for fertility work-up.",
                            key="ov_adv_follicular_monitoring")
                st.checkbox("Adv- LH/FSH & AMH Correlation. (PCOS)",
                            key="ov_adv_lh_fsh")
                st.checkbox("Adv- Clinico-Lab Correlation. (PCOS partial)",
                            key="ov_adv_clinico_lab")
                st.checkbox("Adv- Follow up.", key="ov_adv_followup")
    else:
        pr_open = organ_button("PROSTATE", "PROSTATE")
        if pr_open:
            with st.container(border=True):
                st.checkbox("Not visualized", key="pr_not_visualized")
                if not ss("pr_not_visualized", False):
                    st.checkbox("Partially visualized",
                                key="pr_partial_visualization")
                    if not ss("pr_partial_visualization", False):
                        st.text_input("Prostate volume (cc)  —  auto-graded",
                                      key="pr_cc")
                        _cc_val = ss("pr_cc", "")
                        if _cc_val:
                            _band = prostate_band_from_volume(_cc_val)
                            if _band:
                                st.info(f"→ Auto-grade: **{_band[1]}**")
                        st.checkbox("Median lobe hypertrophy",
                                    key="pr_median_lobe")
                        if ss("pr_median_lobe", False):
                            st.text_input("Median lobe size (mm)",
                                          key="pr_median_lobe_size")
                            st.checkbox("Impinging upon bladder outlet",
                                        key="pr_median_lobe_boo")
                        st.checkbox("Prostatic cyst", key="pr_cyst")
                        if ss("pr_cyst", False):
                            c_a, c_b, c_c = st.columns(3)
                            with c_a:
                                st.text_input("Cyst size (mm, e.g. 12x10)",
                                              key="pr_cyst_size")
                            with c_b:
                                st.checkbox("Left hemi-prostate",
                                            key="pr_cyst_left_hemi")
                            with c_c:
                                st.checkbox("Right hemi-prostate",
                                            key="pr_cyst_right_hemi")

    # ------- BOWEL -------
    bowel_open = organ_button("BOWEL / FREE FLUID", "BOWEL / FREE FLUID")
    if bowel_open:
        with st.container(border=True):
            st.radio("Free fluid", ["none", "minimal", "mild", "moderate"],
                     horizontal=True, key="bw_ff")
            st.radio("Mesenteric LN", ["none", "present"],
                     horizontal=True, key="bw_ln")

    # ------- APPENDIX -------
    ap_open = organ_button("APPENDIX", "APPENDIX")
    if ap_open:
        with st.container(border=True):
            st.radio("Status",
                     ["not_assessed", "not_visualized", "normal", "dilated"],
                     horizontal=True,
                     format_func=lambda x: x.replace("_", " ").title(),
                     key="ap_status")
            if ss("ap_status", "not_assessed") in ("normal", "dilated"):
                st.text_input("Diameter (mm)", key="ap_d")


# ============================================================
# ASSEMBLE DATA
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
    "portal_vein_mm": ss("liver_portal_mm", "")
        if ss("liver_portal", "normal") == "dilated" else "",
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
                uc_sizes = [ss(f"kd_{sd}_uc_s1", ""), ss(f"kd_{sd}_uc_s2", "")]
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
            cyst_type = ss(f"kd_{sd}_cyst_type", "simple")
            cyst_count = ss(f"kd_{sd}_cyst_count", "single")
            cyst_size = ss(f"kd_{sd}_cyst_size", "")
            cyst_loc = ss(f"kd_{sd}_cyst_loc", "")
            if cyst_type == "complex":
                cyst_septa = ss(f"kd_{sd}_cyst_septa", "none")
                cyst_calc = ss(f"kd_{sd}_cyst_calc", "none")
            else:
                cyst_septa = "none"
                cyst_calc = "none"
        else:
            cyst_type = "none"
            cyst_count = "single"
            cyst_size = ""
            cyst_loc = ""
            cyst_septa = "none"
            cyst_calc = "none"

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
            "cyst_type": cyst_type,
            "cyst_count": cyst_count,
            "cyst_size_mm": cyst_size,
            "cyst_location": cyst_loc,
            "cyst_septa": cyst_septa,
            "cyst_calc": cyst_calc,
            "cyst_bosniak": _cyst_bosniak({
                "cyst_type": cyst_type,
                "cyst_septa": cyst_septa,
                "cyst_calc": cyst_calc,
            }) if cyst_type != "none" else "",
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
            "cyst_type": "none",
            "cyst_count": "single",
            "cyst_size_mm": "",
            "cyst_location": "",
            "cyst_septa": "none",
            "cyst_calc": "none",
            "cyst_bosniak": "",
            "hydronephrosis": "none",
            "hydronephrosis_no_obstructive_calculus": False,
            "recently_passed_calculus_suspected": False,
            "contralateral_normal_clause": False,
        })

data["kidneys"]["cortical_echogenicity"] = ss("kd_cort_echo", "normal")
data["kidneys"]["cortical_echogenicity_laterality"] = ss("kd_cort_lat",
                                                         "bilateral")
_echo_val = ss("kd_cort_echo", "normal")
if _echo_val == "mildly_raised":
    data["kidneys"]["cortical_cmd"] = "preserved"
elif _echo_val == "moderately_raised":
    data["kidneys"]["cortical_cmd"] = ss("kd_cort_cmd", "preserved")
elif _echo_val == "significantly_raised":
    data["kidneys"]["cortical_cmd"] = ss("kd_cort_cmd", "hazy")
else:
    data["kidneys"]["cortical_cmd"] = "preserved"
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

# ---- URINARY BLADDER ----
ub_status_val = None
if ss("ub_partially_empty", False):
    ub_status_val = "partially_empty"
elif ss("ub_empty", False):
    ub_status_val = "empty"
elif ss("ub_over", False):
    ub_status_val = "over"

data["urinary_bladder"].update({
    "status": ub_status_val,
    "catheterized": bool(ss("ub_catheterized", False)),
    "pre_void_cc": ss("ub_pre_void", ""),
    "post_void_cc": ss("ub_post_void", ""),
    "mass_calculus": not ss("ub_no_mass", True),
    "sedimentation_free_floating": bool(ss("ub_sed_free_floating", False)),
    "sedimentation_settled_debris": bool(ss("ub_sed_settled_debris", False)),
    "sedimentation_trace": bool(ss("ub_sed_trace", False)),
    "sedimentation_extensive": bool(ss("ub_sed_extensive", False)),
    "uti_suspected": bool(ss("ub_uti", False)),
    "wall_thickening": bool(ss("ub_wall_check", False)),
    "wall_irregular": bool(ss("ub_wall_irregular", False)),
    "wall_mm": ss("ub_wall_mm", ""),
    "force_empty_suboptimal": (data["kidneys"]["right"]["ureter_not_traced"]
                                or data["kidneys"]["left"]["ureter_not_traced"]),
})

# ---- UTERUS (F) or PROSTATE (M) ----
if p_sex == "F":
    ut_status = ss("ut_status", "anteverted")
    adeno_feats = []
    if ss("ut_aden_globular", False):
        adeno_feats.append("globular_shape")
    if ss("ut_aden_asymm", False):
        adeno_feats.append("asymmetric_posterior")
    if ss("ut_aden_venetian", False):
        adeno_feats.append("venetian_blind_sign")
    if ss("ut_aden_iface_indistinct", False):
        adeno_feats.append("interface_indistinct")
    if ss("ut_aden_iface_fundus", False):
        adeno_feats.append("interface_barely_perceptible_fundus")
    if ss("ut_aden_iface_lost", False):
        adeno_feats.append("interface_lost")
    if ss("ut_aden_cysts", False):
        adeno_feats.append("subendometrial_cysts")

    fib_types = []
    fib_lesions = []
    _fib_type_keys = [
        ("ut_fib_type_intramural", "intramural", "ut_fib_type_intramural"),
        ("ut_fib_type_intra_subser", "intramural_subserosal",
         "ut_fib_type_intra_subser"),
        ("ut_fib_type_intra_submuc", "intramural_submucosal",
         "ut_fib_type_intra_submuc"),
        ("ut_fib_type_submuc_intra", "submucosal_intramural",
         "ut_fib_type_submuc_intra"),
        ("ut_fib_type_subser_intra", "subserosal_intramural",
         "ut_fib_type_subser_intra"),
    ]
    for k, type_key, prefix in _fib_type_keys:
        if ss(k, False):
            fib_types.append(type_key)
            fib_lesions.append({
                "location": ss(f"{prefix}_loc", "fundal"),
                "size": ss(f"{prefix}_size", ""),
            })

    data["uterus"].update({
        "status": ut_status,
        "size": ss("ut_size", ""),
        "retroflexed": bool(ss("ut_retroflexed", False)),
        "gravid": bool(ss("ut_gravid", False)),
        "gravid_type": ss("ut_gravid_type", "CRL"),
        "gravid_length": ss("ut_gravid_length", ""),
        "gravid_weeks": ss("ut_gravid_weeks", ""),
        "gravid_days": ss("ut_gravid_days", ""),
        "low_lying": bool(ss("ut_low_lying", False)),
        "cervix_visualized": ss("ut_cervix_vis", "partially"),
        "myometrium": ss("ut_myometrium", "homogenous"),
        "prominent_vascular_channels": bool(ss("ut_prom_vasc", False)),
        "endometrial_thickness_mm": ss("ut_endo_mm", ""),
        "endometrial_collection": ss("ut_endo_coll", "none"),
        "endometrial_collection_het": bool(ss("ut_endo_het", False)),
        "rpoc": bool(ss("ut_rpoc", False)),
        "rpoc_size": ss("ut_rpoc_size", ""),
        "rpoc_echo": ss("ut_rpoc_echo", "hypo"),
        "cervix_elongated": bool(ss("ut_cx_elong", False)),
        "cervix_bulky": bool(ss("ut_cx_bulky", False)),
        "cervix_bulky_size_mm": ss("ut_cx_bulky_size", ""),
        "nabothian": ss("ut_nabothian", "none"),
        "cervix_collection": ss("ut_cx_coll", "none"),
        "fibroid": bool(ss("ut_fibroid", False)),
        "fibroid_count": ss("ut_fibroid_count", "single"),
        "fibroid_types": fib_types,
        "fibroid_lesions": fib_lesions,
        "fibroid_figo_manual": ss("ut_fib_figo_manual", ""),
        "adenomyosis": bool(ss("ut_adenomyosis", False)),
        "adenomyosis_features": adeno_feats,
        "adenomyosis_confidence": ss("ut_aden_conf", "auto"),
        "adenomyomas": bool(ss("ut_adenomyomas", False)),
        "adenomyomas_count": ss("ut_adenoma_count", "couple"),
        "adenomyomas_echo": ss("ut_adenoma_echo", "hyperechoic"),
        "adenomyomas_location": ss("ut_adenoma_loc", "posterior"),
        "adenomyomas_size": ss("ut_adenoma_size", ""),
        "adenomyomas_avascular": bool(ss("ut_adenoma_avascular", True)),
        "neg_rpoc": bool(ss("ut_neg_rpoc", False)),
        "neg_sol": bool(ss("ut_neg_sol", False)),
    })

    for side, sd in (("right", "r"), ("left", "l")):
        o = data["ovaries"][side]
        o["status"] = ss(f"ov_{sd}_status", "normal")
        o["size_text"] = ss(f"ov_{sd}_size", "")
        o["volume_cc"] = ss(f"ov_{sd}_vol", "")
        o["findings"] = []
        if o["status"] == "normal":
            if ss(f"ov_{sd}_f_simple", False):
                o["findings"].append({
                    "type": "simple_cyst",
                    "size_mm": ss(f"ov_{sd}_f_simple_size", ""),
                    "count": ss(f"ov_{sd}_f_simple_count", "single"),
                })
            if ss(f"ov_{sd}_f_hem", False):
                o["findings"].append({
                    "type": "hemorrhagic_cyst",
                    "size_mm": ss(f"ov_{sd}_f_hem_size", ""),
                    "count": ss(f"ov_{sd}_f_hem_count", "single"),
                })
            if ss(f"ov_{sd}_f_hemf", False):
                o["findings"].append({
                    "type": "hemorrhagic_follicle",
                    "size_mm": ss(f"ov_{sd}_f_hemf_size", ""),
                    "count": ss(f"ov_{sd}_f_hemf_count", "single"),
                })
            if ss(f"ov_{sd}_f_foll", False):
                o["findings"].append({
                    "type": "follicular_cyst",
                    "size_mm": ss(f"ov_{sd}_f_foll_size", ""),
                    "count": ss(f"ov_{sd}_f_foll_count", "single"),
                })

    data["ovaries"].update({
        "pcos": bool(ss("ov_pcos", False)),
        "pcos_feature2": bool(ss("ov_pcos_f2", False)),
        "pcos_feature3": bool(ss("ov_pcos_f3", False)),
        "pcos_feature4": bool(ss("ov_pcos_f4", False)),
        "pcos_feature4_variable": bool(ss("ov_pcos_f4_var", False)),
        "pcos_feature4_peripheral": (ss("ov_pcos_f4_dist", "peripheral") == "peripheral"),
        "pcos_feature4_random": (ss("ov_pcos_f4_dist", "peripheral") == "random"),
    })
else:
    ub_empty = (data["urinary_bladder"]["status"] in ("empty", "partially_empty"))
    pr_not_vis = bool(ss("pr_not_visualized", False))
    pr_partial = bool(ss("pr_partial_visualization", False))
    if ub_empty and not pr_not_vis and not pr_partial:
        if data["urinary_bladder"]["status"] == "empty":
            pr_not_vis = True
        else:
            pr_partial = True

    data["prostate"].update({
        "size_cc": ss("pr_cc", ""),
        "median_lobe": bool(ss("pr_median_lobe", False)),
        "median_lobe_size_mm": ss("pr_median_lobe_size", ""),
        "median_lobe_boo": bool(ss("pr_median_lobe_boo", False)),
        "cyst": bool(ss("pr_cyst", False)),
        "cyst_size": ss("pr_cyst_size", ""),
        "cyst_left_hemi": bool(ss("pr_cyst_left_hemi", False)),
        "cyst_right_hemi": bool(ss("pr_cyst_right_hemi", False)),
        "not_visualized": pr_not_vis,
        "partial_visualization": pr_partial,
    })

data["bowel"].update({
    "free_fluid": ss("bw_ff", "none"),
    "mesenteric_ln": ss("bw_ln", "none"),
})
data["appendix"].update({
    "status": ss("ap_status", "not_assessed"),
    "diameter_mm": ss("ap_d", ""),
})

data["ovary_manual_advices"] = {
    "follicular_monitoring": bool(ss("ov_adv_follicular_monitoring", False)),
    "lh_fsh": bool(ss("ov_adv_lh_fsh", False)),
    "clinico_lab": bool(ss("ov_adv_clinico_lab", False)),
    "followup": bool(ss("ov_adv_followup", False)),
}


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
        height=120, label_visibility="collapsed",
        key="additional_body_findings_box",
        placeholder="e.g. A small umbilical hernia is noted in the anterior "
                    "abdominal wall.",
        on_change=_cap_addendum_cb)

data["additional_body_findings"] = addendum_text or ""


def render_preview_findings(data):
    p = data["patient"]
    out = ["         ULTRASOUND WHOLE ABDOMEN", ""]
    sex, age = p["sex"], p["age"]
    age_years = parse_age(age)
    is_ped = age_years is not None and age_years < 18
    is_ped_female = (sex == "F" and age_years is not None and age_years < 14)
    p_status = data["pancreas"].get("status", "normal")
    spleen_enlarged = data["spleen"]["size_descriptor"] != "normal"
    ub_status = data["urinary_bladder"].get("status") or "adequately_distended"
    secs = [liver_sentence(data["liver"], sex, age,
                            spleen_enlarged=spleen_enlarged,
                            cbd_ihbr=data["cbd"].get("ihbr", "normal")),
            gall_bladder_sentence(data["gall_bladder"]),
            cbd_sentence(data["cbd"]),
            pancreas_sentence(data["pancreas"]),
            spleen_sentence(data["spleen"]),
            kidneys_sentence(data["kidneys"], ub_status=ub_status, age_text=age),
            urinary_bladder_sentence(data["urinary_bladder"])]
    if sex == "F":
        secs.append(uterus_sentence(data["uterus"], pediatric=is_ped_female,
                                     age_years=age_years))
        uterus_operated = data["uterus"].get("status") == "operated"
        secs.append(ovaries_sentence(data["ovaries"], age_years=age_years,
                                      uterus_operated=uterus_operated))
    else:
        secs.append(prostate_sentence(data["prostate"], pediatric=is_ped,
                                       ub_status=ub_status))
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
    st.markdown(f'<div class="preview-box">{_safe}</div>', unsafe_allow_html=True)


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
