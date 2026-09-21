# Radiology Report Generator — Project State

Last updated: 2026-09-21
Repo: github.com/Kukkuthakur/radiology-reports
Files: `app.py`, `requirements.txt`, `PROJECT_STATE.md`
Deployed: Streamlit Cloud (working)

---

## Tech Stack
- Streamlit web app (Phase 1) — will migrate to .exe later
- python-docx for output
- SQLite for local storage (ephemeral on Streamlit Cloud)
- Single-file app: `app.py`

## UI Layout (current)
Top → Bottom:
1. Patient demographics row (Name / Age / Sex / Date / Ref By)
2. Findings (expanders, one per organ)
3. Live Preview (text area)
4. Editable Impression (text area)
5. Actions (Download .docx / Save to DB)

## Global Rules
- Default = Normal report (nothing selected → all normal)
- Male → Prostate section; Female → Uterus + Ovaries
- Age < 18 → pediatric wording ("for age", "age-appropriate")
- Age ≥ 18 → adult rules
- All fields and impression lines editable
- Full auto-impression with manual override

## Formatting (Locked)
- Page: A4 (21 × 29.7 cm)
- Margins: Top 4.5 cm, Bottom 3.5 cm, Left 2 cm, Right 2 cm
  - ⚠️ Earlier confirmed 4 cm top; code has 4.5 cm. Reconcile when convenient.
- Fonts: Calibri 11pt body, 13pt title, 10pt disclaimer
- Title colour: RGB(0x1F, 0x4E, 0x79) — dark blue
- **BLUE + BOLD + UNDERLINED**: "ULTRASOUND WHOLE ABDOMEN" (centered) and "IMPRESSION:"
- Header table (NAME/DATE/AGE-SEX/REF.BY): bordered, 2×2, bold text
- Disclaimer: single-cell bordered table, bold, 10pt
- Impression lines: bold, with "Adv-…" part italicised

---

## LOCKED ORGANS

### LIVER ✅
Size (adult):
- < 150 → Normal | = 150 → Borderline | 151–170 → Mild | 171–190 → Moderate | > 190 → Gross
Size (pediatric, age < 18):
- Compare to age table (mm). Above upper limit → "Enlarged for age"
- Impression: `HEPATOMEGALY FOR AGE`
- Auto-classification from size + age + sex; user can override

Echotexture: Normal / Increased (steatosis) / Coarse / Low
Outline: Normal / Crenated-nodular
Steatosis grades: `Mild+` / `Mild to Moderate++` / `Moderate++` / `Moderate to Severe+++` / `Severe+++`
- Severe+++ → `SIGNIFICANT FATTY INFILTRATION(SEVERE+++) - SUSPICIOUS FOR NON-ALCOHOLIC STEAT-HEPATITIS`
Focal lesions (with impression rules):
- Simple cyst — single/few, lobe, size → `A/FEW SIMPLE HEPATIC CYST(S) IN {LOBE} LOBE.`
- Hemangioma — single/few, lobe, size → `A/FEW HEPATIC HEMANGIOMA(S) IN {LOBE} LOBE.`
- Abscess — single/few/multiple, segments I–VIII, dims, vol
  - Wording: `An irregular marginated ill-defined avascular SOL(...)`
  - Impression: `... - likely Liver Abscess(es). Adv- Lab & Clinical Correlation.`
- Calcified focus (right lobe)
- Free text
IHBR: Normal / Dilated
Portal vein: Normal / Dilated (mm)
Impression format: `{SIZE_DESCRIPTOR} WITH HEPATIC STEATOSIS({GRADE}). Adv- LFT Correlation.`

### SPLEEN ✅
Size (adult):
- < 120 → Normal | 120–130 → Borderline | 130–145 → Mild | > 145 → Moderate
Size (pediatric): age × sex table; above upper limit → `Enlarged for age`
Impression: `SPLENOMEGALY FOR AGE` / `MILD/MODERATE SPLENOMEGALY (mm)`
Portal vein mention with size (when relevant)

### Pediatric Reference Tables ✅ (in code)
- `PEDIATRIC_SPLEEN_MAX_MM` — (age_min, age_max, sex) → upper limit mm
- `PEDIATRIC_LIVER_MAX_MM` — (age_min, age_max) → upper limit mm

---

## IN PROGRESS (Code written, needs final validation)

### GALL BLADDER 🟡 (code written, awaiting review)
Distension status: Adequately distended / Over-distended / Partially contracted / Contracted / Empty / Operated
Wall thickness: Normal / Thickened (mm); auto "MILDLY" ≤ 8mm, "SIGNIFICANTLY" > 8mm
Calculi:
- Count: single / few / multiple / innumerable
- Size category: small / large; size (mm)
- Neck location (with separate neck size)
- Wording: `A/Few/Multiple {size_cat} calculi seen in the GB lumen, largest … mm`
Sludge: None / Trace / Significant / Echogenic / Organized
Sludge ball/polyp: checkbox + count + size + wall (anterior/posterior)
Comet tail: checkbox + count + wall (anterior/posterior)
Pericholecystic fluid: checkbox
Impression rules:
- Over + calculi + wall/peri → ACUTE CHOLECYSTITIS
- Over + calculi → cholelithiasis, no acute features
- Contracted + calculi → ?chronic cholecystitis
- Calculi + wall/peri → ?acute cholecystitis
- Calculi only → `CHOLELITHIASIS WITH NO APPRECIABLE PERICHOLECYSTIC FLUID OR GB WALL THICKENING.`
- Wall thick + peri fluid (no calculi) → ?acalculus cholecystitis
- Wall thick alone → ?acalculus cholecystitis
- Peri fluid alone → ?significance
- Sludge → `{GRADE} SLUDGE SEEN IN THE GALLBLADDER LUMEN. Adv- Review scan after a month.`
- Sludge ball → similar + review
- Comet tail → `GALL BLADDER ADENOMYOMATOSIS/CHOLESTEROLOSIS.`

---

## PENDING (Not yet started)

### CBD
- Caliber (mm)
- Status: Normal / Dilated
- Calculi in CBD (size, proximal/mid/distal)
- IHBR dilation

### PANCREAS
- Normal / Fat stranding / Necrotic LN
- MPD dilation

### KIDNEYS (structured calculi)
- Per-kidney: status, calculi (multi, size + location), cyst (type/size/location/Bosniak), hydronephrosis, nephrocalcinosis
- Cortical echogenicity

### URINARY BLADDER
- Status + sedimentation

### UTERUS
- Statuses: anteverted, retroverted, retroflexed, bulky, operated, not_visualized, partially
- Size, ET, myometrium, fibroid
- **GRAVID UTERUS** (from SANIYA sample — see below)

### OVARIES
- Per-ovary: normal / cyst (type + size) / not visualized / atrophic
- AFC + dominant follicle

### PROSTATE
- Normal / Borderline / Bulky / Grade-I BPH
- Size in cc

### BOWEL / MESENTERY / FREE FLUID
- Wall thickening, free fluid levels, mesenteric LN (size cat, location, character), pleural effusion

### APPENDIX
- Not assessed / Not visualized / Normal (mm) / Dilated (mm)

### Later Modules (Phase 2+)
- Early OBS scan
- OBS (16–28 weeks)
- OBS with Doppler
- Anomaly scan (Level II)
- IUD / fetal demise
- Scrotal colour Doppler
- Pediatric WA variants

---

## SPECIAL CASES NOTED (to handle later)
- **Gravid uterus** in WA report (SANIYA sample): 
  `UTERUS is GRAVID(ALIVE; GA= Approx. 30 Weeks, AFI is adequate. Placenta is fundo-posterior & away from the internal Os).`
  Impression is multi-line parenthetical.

---

## Open Questions / To Verify
- Top margin: 4 cm (earlier) vs 4.5 cm (code) — reconcile
- "marinated" → fixed to "marginated" ✅
- Pediatric liver/spleen: upper limit logic validated with sample reports ✅

---

## Session Log
- 2026-09-21: Liver + Spleen locked. Gall Bladder code written (needs validation). Formatting locked. Pediatric tables added. Layout restructured to Demographics → Findings → Preview → Actions.

## Next Step
1. Validate Gall Bladder output against real reports
2. Move to CBD, then Pancreas
3. Continue organ by organ in order
