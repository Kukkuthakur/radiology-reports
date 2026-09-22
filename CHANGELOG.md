# Changelog — Radiology Report Generator

## v1.0-stable — Frozen
Baseline working version. All features below are verified working.

### Layout
- Two-column layout: Findings (left) | Live Preview + Impression + Actions (right)
- Every organ takes one line when collapsed; expander + size box aligned on same row
- Light theme forced via CSS regardless of browser dark mode
- Mobile-safe: columns do not stack on narrow viewports
- Header: "PG Imaging & Diagnostics" + tagline

### Findings panel
- Demographics row: Name | Age | Sex | Date | Referred by (single line)
- LIVER — size box + auto status badge (normal / mild / moderate / gross / enlarged-for-age)
- GALL BLADDER — distension, wall, calculi (single/few/multiple/innumerable), sludge, sludge ball, comet tail, pericholecystic fluid
- COMMON BILE DUCT — decimal size (step 0.1), status, calculus (single/few) with location
- PANCREAS — normal / fat stranding / necrotic LN
- SPLEEN — size box + auto status badge
- KIDNEYS — right/left status, calculi as free text
- URINARY BLADDER — status + sedimentation
- UTERUS (with size + endometrium) + OVARIES (Rt + LT) — on same line as expander
- PROSTATE (male) — size in cc
- BOWEL / FREE FLUID — free fluid + mesenteric LN
- APPENDIX (optional)

### Live Preview
- Black background, white monospace text (rendered as HTML `<pre>`-style block)
- Word-wraps; no horizontal scroll
- Updates on every rerun (no widget caching)
- Demographics + disclaimer stripped out
- Shows: title → all organ findings → IMPRESSION

### Impression editor
- Black background, white monospace text
- Auto-syncs to auto-generated impression whenever findings change
- Preserves manual edits (only auto-refreshes if user hasn't typed)
- "↺ Reset to auto-generated impression" button
- Feeds directly into downloaded .docx

### DOCX export
- Page: A4, top margin 4.0 cm, bottom 3.5 cm, left/right 2.0 cm
- Demographics table (2×2, bordered)
- Centered title "ULTRASOUND WHOLE ABDOMEN" (bold, underlined, #1F4E79)
- Findings paragraphs: organ name bold+underline, abnormal findings bold+italic+mixed case
- Impression: bulleted, ALL CAPS, "Adv-" part bold+italic
- Disclaimer: bordered box, bold, 8pt, no extra trailing space
- Filename: `{name}_{YYYYMMDD_HHMMSS}.docx`, sanitized

### Database
- SQLite at `~/RadiologyReports/reports.db`
- `referrers` table (seeded with 17 names)
- `reports` table with full JSON snapshot per save
- "💾 Save to Database" button

### Known working behaviors (do not break)
- CBD accepts decimals (5.4 stays 5.4 in report)
- Liver/Spleen auto-classify from size + age + sex (pediatric tables for <18)
- Preview reflects radio changes within one rerun
- DOCX always matches the impression currently shown in the editable box
