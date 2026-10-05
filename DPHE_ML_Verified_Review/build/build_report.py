# -*- coding: utf-8 -*-
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from analysis import (P, N, count, crosstab, FAMILY_ORDER, INPUT_ORDER, OUTPUT_ORDER, VAL_ORDER, BINS)
from content import MODEL_TEXT, GAPS, BENCHMARKS, RECOMMENDATION
from paragraphs import PARA
from removed_data import REMOVED

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, "figs")
os.makedirs(FIG, exist_ok=True)
OUT = os.path.join(HERE, "..", "DPHE_ML_Literature_Review_v2_Verified.docx")
STATS = json.load(open(os.path.join(HERE, "stats.json")))
BY = {d["id"]: d for d in P}
INK, MUTED, BAR = "#1F2937", "#6B7280", "#2F5597"


def hbar(name, od, title, sort=True, xlabel="Number of papers"):
    items = list(od.items())
    if sort:
        items.sort(key=lambda kv: -kv[1])
    labels = [str(k) for k, _ in items][::-1]
    vals = [v for _, v in items][::-1]
    fig, ax = plt.subplots(figsize=(6.4, 0.32 * len(items) + 1.0), dpi=200)
    ax.barh(labels, vals, color=BAR, height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + 0.15, i, f"{v} ({v / N:.0%})", va="center", fontsize=7.5, color=INK)
    ax.set_title(title, fontsize=9.5, color=INK, loc="left", fontweight="bold")
    ax.set_xlabel(xlabel, fontsize=8, color=MUTED)
    ax.tick_params(labelsize=7.5, colors=INK)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color("#D1D5DB")
    ax.set_xlim(0, max(vals + [1]) * 1.25 + 0.5)
    ax.xaxis.grid(True, color="#EEEEEE", linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    p = os.path.join(FIG, name + ".png")
    fig.savefig(p)
    plt.close(fig)
    return p


def heat(name, t, title):
    rows = list(t.keys())
    cols = list(next(iter(t.values())).keys())
    import numpy as np
    m = np.array([[t[r][c] for c in cols] for r in rows])
    fig, ax = plt.subplots(figsize=(6.4, 0.32 * len(rows) + 1.6), dpi=200)
    ax.imshow(m, cmap="Blues", aspect="auto", vmin=0, vmax=max(1, m.max()))
    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(cols, fontsize=7, rotation=25, ha="right")
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels(rows, fontsize=7)
    for i in range(len(rows)):
        for j in range(len(cols)):
            v = m[i, j]
            ax.text(j, i, str(v), ha="center", va="center", fontsize=7, color="white" if v > m.max() * 0.6 else INK)
    ax.set_title(title, fontsize=9.5, loc="left", fontweight="bold", color=INK)
    fig.tight_layout()
    p = os.path.join(FIG, name + ".png")
    fig.savefig(p)
    plt.close(fig)
    return p


doc = Document()
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(10.5)
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2.0)


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def H(text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return h


def para(text, bold_prefix=None, size=None, italic=False):
    p = doc.add_paragraph()
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(6)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
    p.add_run(text)
    return p


def table(header, rows, widths=None, size=8.5, head_fill="1F3864"):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(header):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold, r.font.size = True, Pt(size)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, head_fill)
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            r = cells[j].paragraphs[0].add_run(str(v))
            r.font.size = Pt(size)
    if widths:
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
    doc.add_paragraph()
    return t


def fig(path, caption, width=15.5):
    doc.add_picture(path, width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = doc.add_paragraph(caption)
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in c.runs:
        r.italic, r.font.size = True, Pt(8.5)


def k(cond):
    return sum(1 for d in P if cond(d))


# ------------------------------------------------------------------ Title
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("Machine-Learning Prediction of Double-Pipe Heat Exchanger Thermal Performance: A Focused Structured Literature Review")
r.bold, r.font.size = True, Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.CENTER
s.add_run("Second-stage, full-text-verified review - rewritten from scratch on the verified database\n").italic = True
s.add_run("Prepared for: Ali Abdelwahab (Kuwait University) - MgO-TiO2 hybrid-nanofluid double-pipe heat exchanger project\n5 October 2026")
para("Companion files: DPHE_ML_Literature_Log_v2_Verified.xlsx (database, statistics dashboard, removed papers, PDF status, rename map); rename_pdfs.ps1 / rename_pdfs.sh. Scope: thermal-performance prediction only; hydraulic quantities are recorded for completeness but never steer conclusions.", size=9, italic=True)

# ------------------------------------------------------------------ 1 Executive summary
H("1. Executive Summary")
para(f"All 26 PDFs in the project 'Papers' folder were opened and matched to the first-pass database by title, authors and DOI read inside each PDF; none was mismatched and none was unmatched. Combined with the 37 first-pass records, the verified active database contains {N} papers (Tier 1: 8, Tier 2: 8, Tier 3: 8), of which 22 rest on full PDFs (Evidence A: 12, B: 10) and two (P01, P09) are retained abstract-only exceptions. Thirteen records were moved to a 'Removed Papers' sheet (nine without accessible evidence, three reviews and one meta-science reference).")
para("The verified literature is weaker than the first-pass map suggested. The key findings are:")
for b in [
    "Every double-pipe study with experimental data validated its model with a random row split or did not report a split. None withheld whole operating conditions, concentrations or flow levels, and none reported prediction intervals.",
    "High headline accuracies (R or R² ≥ 0.99 in 12 papers) mostly come from random splits of small factorial grids, often with networks whose parameters outnumber the data points (P10: about 9 runs and a 20-neuron network; P05: 111 weights for 72 points). In two papers the reported error is lower than the measurement uncertainty (P04, P16).",
    "The two strongest methodological papers come from outside heat exchangers. Withholding entire experimental campaigns raised a gradient-boosting model's error 2.5-fold, whereas a model anchored to a physical correlation barely degraded (P27). On unseen tube geometries a physics correlation beat the ML model (P29). A dimensionless, physics-structured network extrapolated far better than a dimensional one (P28).",
    "Studies that predict effectiveness, NTU and heat duty learn them as independent targets (P02, P07), so their predictions can violate the energy balance.",
    "Only one double-pipe hybrid-nanofluid ML study was identified (P01), and its full text could not be accessed.",
]:
    bullet(b)
para("Benchmarks after full-text review:", bold_prefix=None)
bullet(" P02 (Esfandyari et al., 2023), primary. It is the closest experimental analogue in fluid, inputs and thermal outputs.", "Primary:")
bullet(" P13 (Alsayah et al., 2025), secondary. It has the best-documented pipeline: cross-validation for model selection, an untouched test set, physical-unit errors and correlation baselines.", "Secondary:")
bullet(" P37 (Hadj Salah et al., 2026), supporting. It has the best leakage controls and the only calibrated uncertainty, plus mono-to-hybrid transfer learning.", "Supporting:")
para("Recommendation in brief:", bold_prefix=None)
for b in [
    "Learn one thermal transport quantity: the overall coefficient U, or the tube-side Nu if wall temperatures are reliable. Preferably learn it as a log-ratio to a correlation baseline. Derive NTU, effectiveness, heat duty and outlet temperatures through ε-NTU and energy-balance relations.",
    "Use dimensionless inputs: Re and Pr of each stream, concentration and the heat-capacity-rate ratio. Average repeated readings so that each operating condition is one row.",
    "Compare the following models under nested grouped validation (leave-one-condition-out, leave-one-concentration-out, leave-one-flow-level-out): a refitted power-law baseline, regularised log-linear regression, GPR, SVR, random forest, shallow gradient boosting, and a small MLP ensemble.",
    "Report physical-unit errors against the measurement uncertainty of U. Give calibrated prediction intervals, and refuse queries outside the experimental envelope.",
    "Deep learning, ANFIS and metaheuristic-tuned networks are not justified by the evidence for a dataset of tens of operating conditions.",
]:
    bullet(b)

# ------------------------------------------------------------------ 2 Introduction
H("2. Introduction")
para("Double-pipe (concentric-tube) heat exchangers are the standard laboratory platform for nanofluid heat-transfer studies because both streams are measurable and the geometry is simple. Each experimental operating point is costly - nanofluid preparation, stabilisation, steady-state attainment and cleaning - so a surrogate model that predicts the thermal performance at untested combinations of flow rates, inlet temperatures and concentrations is attractive. The project aims to train such a surrogate on the authors' own experimental MgO-TiO2 hybrid-nanofluid DPHE data and to package it as a user-facing application.")
para("A surrogate is only useful if its accuracy on genuinely new operating conditions is known. This review therefore focuses on how published ML models for double-pipe and closely related exchangers were built and validated - data structure, inputs, thermal targets, model capacity, validation design, leakage, uncertainty and physics integration - rather than on the heat-transfer enhancements reported.")

# ------------------------------------------------------------------ 3 Methodology
H("3. Review Methodology")
para("The first-pass review (4 October 2026) screened peer-reviewed literature from 2022 onwards and logged 37 papers plus nine screening candidates, many assessed only from abstracts or index records. This second stage re-built the evidence base from full texts:")
for b in [
    "Every PDF in the OneDrive folder 'Thesis/Machine learning project/ML_DPHE_Papers' was read through the Microsoft 365 connector; identity (title, authors, year, journal, DOI) was taken from the PDF itself, not from the filename.",
    "For each retained paper the full extraction template (bibliography, study design, ML methodology, performance, critical assessment) was filled from the PDF; values contradicting the first-pass log were overwritten; items absent from the paper are recorded as 'Not reported'; reviewer deductions are labelled as inferences.",
    "Evidence levels: A = full paper read directly; B = full PDF available, with data, ML set-up, validation, results tables and conclusions verified through targeted extraction; C = abstract only.",
    "Tiers were reassessed against the stated definitions (Tier 1: double-pipe/concentric + experimental + ML thermal prediction; Tier 2: strong transferable methodology; Tier 3: supporting). Relevance (1-5) and methodological quality (1-5) were scored independently.",
    "Validation was audited against the leakage taxonomy of Kapoor and Narayanan (2023): no test set, preprocessing or feature selection on the full data, duplicates, illegitimate (target-derived) features, and non-independence between training and test samples.",
    "Statistics and charts were generated programmatically from the database so that the report, the dashboard and the log cannot disagree.",
]:
    bullet(b)
para("Limitations of the process: journal quartiles could not be verified (SCImago and OpenAlex were blocked by the session's network policy) and are recorded as 'Not verified'; the OneDrive connector was read-only, so PDF renaming is delivered as a mapping and scripts rather than performed in place.", italic=True)

# ------------------------------------------------------------------ 4 Screening & verification
H("4. Screening and Verification Process")
table(["Step", "Count", "Notes"], [
    ["First-pass records (P01-P37)", 37, "Plus 9 screening candidates (X01-X09) never logged"],
    ["PDFs in Papers folder", 26, "Plus one non-paper file (software installer)"],
    ["PDFs matched to a Paper ID from PDF content", 26, "0 mismatched, 0 unmatched"],
    ["Records without PDF", 11, "P01, P08, P09, P11, P20, P23, P24, P25, P34, P35, P36"],
    ["Retained abstract-only (exception)", 2, "P01 (only hybrid-NF DPHE ML study); P09 (operating-variable inputs + regression baseline)"],
    ["Removed - no accessible evidence", 8, "P08, P11, P20, P23, P24, P34, P35, P36"],
    ["Removed - insufficient methodological relevance", 1, "P25"],
    ["Removed - review / supporting only", 4, "P30 (leakage reference - still cited), P31, P32, P33 (reviews)"],
    ["Active (retained) database", N, "Tier 1: 8 | Tier 2: 8 | Tier 3: 8; Evidence A 12, B 10, C 2"],
], widths=[6, 2, 9])
para("Corrections to the first-pass log included: P29's DOI (previously 'not retrieved') and P33's and P37's author lists were recovered from the PDFs; P03 was found to be an air/steam constant-wall-temperature exchanger rather than a liquid-liquid DPHE; P06's 'experimental' heat-transfer coefficient is partly computed from correlations; P18's network uses linear transfer functions; P12's data are analytical; and P05 reports no train/test split at all. Previous 'Q1 (recalled)' quartile labels were removed.")

# ------------------------------------------------------------------ 5 Landscape
H("5. Literature Landscape")
para(f"The 24 retained papers span 2022-2026. Fourteen concern double-pipe or concentric-tube exchangers (eight experimental Tier 1 papers, two air/steam DPHE papers, two CFD design studies, one CFD surrogate and one analytical design study); seven use single tubes or coils; three come from other thermal problems (critical heat flux, supercritical convection, condensation) and are retained for their validation methodology. Eighteen use experimental data, four CFD, one analytical design calculations and one a mixture. Twelve involve mono nanofluids and three hybrid nanofluids (only one of them in a DPHE, abstract-only).")
fig(hbar("hx", count("hx_class"), "Retained papers by heat-exchanger class"), "Figure 1. Heat-exchanger class (n = 24).")
fig(hbar("fluid", count("fluid_class", ["Conventional fluid", "Nanofluid", "Hybrid nanofluid"]), "Retained papers by fluid type", sort=False), "Figure 2. Fluid type.")

# ------------------------------------------------------------------ 6 Individual reviews
H("6. Individual Paper Reviews")
para("Papers are grouped by tier; benchmark papers are listed first. Each paragraph concentrates on ML methodology.", italic=True)
groups = [("6.1 Benchmark papers", lambda d: d["benchmark_status"] == "Yes")] + [
    (f"6.{i + 2} {t} papers (non-benchmark)", (lambda t: lambda d: d["tier"] == t and d["benchmark_status"] != "Yes")(t)) for i, t in enumerate(["Tier 1", "Tier 2", "Tier 3"])]
for title, cond in groups:
    H(title, 2)
    for d in P:
        if not cond(d):
            continue
        tag = f" [{d['benchmark_role']}]" if d["benchmark_status"] == "Yes" else ""
        p = doc.add_paragraph()
        r = p.add_run(f"{d['id']} - {d['first_author']} ({d['year']}), {d['journal'].split(',')[0]}. Evidence {d['evidence']}; {d['tier']}{tag}. ")
        r.bold = True
        r.font.size = Pt(10)
        p.add_run(PARA[d["id"]])
        p.paragraph_format.space_after = Pt(8)

# ------------------------------------------------------------------ 7 Comparative ML methods
H("7. Comparative Analysis of ML Methods")
fams = count("families", FAMILY_ORDER, multi=True)
fig(hbar("families", fams, "Frequency of ML model families (multi-label)"), "Figure 3. Model families used across retained papers.")
para(f"Neural networks dominate ({fams['ANN/MLP']} of {N} papers), usually Levenberg-Marquardt-trained MLPs sized by trial and error. Tree ensembles follow (random forest 6, XGBoost 6, other gradient boosting 4), while kernel methods are rare (SVR 3, GPR 1). Seventeen papers compare several models, but almost always under random splits, so the rankings mainly reflect interpolation smoothness. Where a properly structured simple model was included it was competitive: a six-parameter quadratic was within 1 Nu unit of a 111-weight MLP (P05); an RSM quadratic matched the MLP (P06); a power-law correlation matched XGBoost (P14). Conversely, untransformed linear regression used as a baseline (P16) is a straw man; a log-linear (power-law) form is the appropriate reference for convective heat transfer.")
rows = []
for name, t in MODEL_TEXT.items():
    rows.append([name, t["small"], t["extrap"], t["uq"], t["rel"]])
table(["Model family", "Small-data suitability", "Extrapolation", "Uncertainty", "Relevance to our project"], rows, widths=[3.8, 3, 3.2, 3, 4.5], size=8)

# ------------------------------------------------------------------ 8 Dataset size
H("8. Dataset Size and Data Structure")
fig(hbar("size", count("size_bin", BINS), "Reported number of ML samples", sort=False), "Figure 4. Dataset-size distribution (9 papers do not report n).")
para("Experimental thermal datasets are small: 72 (P05), 84 (P02), 70 (P15), 120 (P17) and about 9-20 runs (P10, P06, inferred). Large datasets come only from CFD (P13: 2,700; P37: 200) or compiled public databases (P27: 24,579; P29: 4,333). More important than the row count is the number of independent operating conditions: P04's 540 rows originate from a few conditions, and no Tier 1 paper states whether repeated steady-state readings were averaged - only P19 explicitly averaged three repeats per configuration before modelling. Factorial or Taguchi designs (P02, P03, P05, P06, P10) mean that every condition has close neighbours differing in one factor level, which makes random splits easy and uninformative.")
heat_t = crosstab("size_bin", "complexity", BINS, ["Low", "Medium", "High", "Unknown"])
fig(heat("size_complex", heat_t, "Dataset size x model complexity"), "Figure 5. Model complexity is not matched to dataset size.", width=13)

# ------------------------------------------------------------------ 9 Inputs
H("9. Input-Variable Analysis")
fig(hbar("inputs", count("input_cats", INPUT_ORDER, multi=True), "Input-variable categories (multi-label)"), "Figure 6. Input variables.")
para("Reynolds number (17 papers), geometric ratios (13) and concentration (11) dominate; geometry is irrelevant to our fixed plain exchanger. Operating-temperature inputs are rare (hot inlet temperature in 3 papers) and no paper used the cold inlet temperature. No double-pipe paper varied both flow rates and both inlet temperatures together with concentration: P02 fixed the cold stream, P06 fixed the hot flow and P07 fixed the hot stream. Dimensional inputs (P02, P06, P10) are common in Tier 1, whereas the strongest extrapolation evidence favours dimensionless groups (P28) and physics-informed features (P29). Feeding correlation intermediates as inputs (friction factors and pumping powers in P12) or any quantity computed from the measured outlet temperatures constitutes target leakage.")

# ------------------------------------------------------------------ 10 Outputs
H("10. Thermal Output Analysis")
fig(hbar("outputs", count("output_cats", OUTPUT_ORDER, multi=True), "Predicted thermal outputs (multi-label)"), "Figure 7. Thermal outputs (hydraulic outputs excluded by scope).")
para("Nusselt number is the dominant target (16 papers), followed by heat-transfer coefficient (6), effectiveness and NTU (4 each), thermal performance factor (3), overall coefficient U (2) and heat duty (1); no paper learns outlet temperatures. Two observations decide our target choice:")
bullet(" P02 and P07 train separate models for Q, Nu, NTU and effectiveness. Because effectiveness is a deterministic function of NTU and the capacity-rate ratio, independent models can return mutually inconsistent values; P02's NTU and effectiveness models also generalised worst (train/test MSE ratios of 25-65).", "Linked outputs learned independently.")
bullet(" P28 learned Nu from dimensionless groups and recovered the dimensional wall temperature through Newton's law; the directly learned dimensional network produced non-physical predictions out of range. P13 shows that U is learnable as a single transport quantity.", "Learning one transport quantity and deriving the rest.")
para("Conclusion: learn U (or tube-side Nu when wall temperatures are reliable), preferably as a ratio to a correlation, and derive NTU, effectiveness, Q and outlet temperatures physically. Learning outlet temperatures directly is not recommended: temperature differences are small relative to absolute temperatures, so small temperature errors become large errors in derived U or Q, and there is no literature support for it.")

# ------------------------------------------------------------------ 11 Validation
H("11. Validation Methodology")
fig(hbar("validation", count("val_multi", VAL_ORDER, multi=True), "Validation methods (multi-label)", sort=False), "Figure 8. Validation methods.")
para("Seventeen papers used a random train/test split and five random k-fold CV; five reported no clear validation (P01, P05, P09, P10, P19). Grouped validation appears only in P27 (campaign-disjoint, leave-campaign-out) and P29 (unseen tube geometries); separate external datasets appear in P28 (new experiments outside the training range) and P26 (ten experiments for a CFD-trained model); confirmation simulations in P21 and P22. No paper used leave-one-concentration-out or leave-one-Re-level-out validation, and no double-pipe paper performed independent confirmation experiments. Model selection is frequently contaminated by the test set: architecture chosen on pooled performance (P02), on minimum MSE including test points (P10), or with a goodness function that includes test metrics (P21, P22). P13 and P37 are the exceptions, selecting models by CV on training data and evaluating the test set once.")
para("Three regimes must be distinguished in our project:")
bullet(" predicting between tested points of the grid - what random splits measure;", "Interpolation:")
bullet(" predicting beyond the tested range of any input - not to be offered by the application;", "Extrapolation:")
bullet(" predicting an operating condition, concentration or flow level withheld entirely (all repeats) from training - the quantity that justifies replacing experiments, and the only one that should be reported as the model's accuracy.", "True held-out-condition prediction:")

# ------------------------------------------------------------------ 12 Leakage
H("12. Data Leakage and Generalisation Concerns")
rows = [
    ["Rows from the same condition in train and test", "P04 (540 rows from few conditions); possible in P02, P05, P07, P15, P16 (repeats not reported)", "Average repeats; group by condition"],
    ["Repeated time samples treated as independent", "Not reported by any Tier 1 paper; averaged in P19 only", "One row per steady state"],
    ["Random split of factorial grid", "P02, P03, P05?, P06, P07, P10, P17", "Leave-condition / level-out CV"],
    ["Scaling before splitting", "P06, P17 (min-max on full data); P13 scope not stated", "Fit scaler inside each training fold"],
    ["Feature selection / target-derived features", "P12 (correlation intermediates as inputs); P06 target from correlations", "Inputs only from controlled variables"],
    ["Hyperparameters tuned on test data", "P02, P10, P21, P22 (likely or explicit)", "Nested CV"],
    ["Only training/fit metrics reported", "P05, P19 (no hold-out); P10 (R only)", "Report held-out metrics"],
    ["R/R2 without physical-unit error", "P10, P02 (MSE units unclear), P17 (normalised only), P04", "RMSE/MAE/MAPE in physical units vs measurement uncertainty"],
    ["Error below measurement uncertainty", "P04 (0.33% vs 1.9%), P16 (0.6% vs 2.4-4.1%)", "Sanity check against uncertainty"],
]
table(["Leakage / generalisation issue", "Where observed", "Remedy for our project"], rows, widths=[4.6, 7, 5], size=8)
para("P27 quantifies the consequence: the same learner's error more than doubled (4.69% to 11.91%) when campaigns were withheld instead of random records, while correlation-anchored learners changed little. P29 shows the same pattern for unseen geometries. The reported accuracies in Tier 1 should therefore be read as interpolation fits, not as evidence that experiments can be replaced.")

# ------------------------------------------------------------------ 13 UQ
H("13. Uncertainty Quantification")
para("Experimental measurement uncertainty is reported in most experimental papers (e.g., Nu +/-1.6-4.3%), but predictive uncertainty is almost absent: only P37 provides prediction intervals (quantile XGBoost with 91% coverage of a nominal 90% interval; a bootstrap ensemble that under-covered at 72.5%), and P27 provides one-sided tolerance divisors. P03 used GPR without reporting its predictive variance. No double-pipe study reports intervals or compares model error with measurement uncertainty. For our project, GPR offers native variance; for other learners, grouped conformal intervals give distribution-free coverage; in all cases empirical coverage must be checked on held-out conditions, because P37 shows that uncalibrated ensembles are overconfident.")

# ------------------------------------------------------------------ 14 Physics
H("14. Physics-Informed and Physics-Anchored Approaches")
para("Only two retained papers embed physics in the learning problem, both outside heat exchangers. P27 learns the multiplicative log-residual of the Bowring correlation and achieves lower grouped error (3.94-4.96%) and far lower leakage amplification (1.06-1.36) than direct learners (2.54); it also tolerated a 6% bias in the anchor. P28 shows that a network with dimensionless inputs and a Nu output extrapolates better and remains physically consistent with only 200 samples. P29 adds that expert (physics-based) features improve ML accuracy and that a plain correlation beats ML on unseen geometries. In heat-exchanger ML, correlations appear only as competitors (P05, P13, P14, P15) or for rig validation (P02, P07, P17). Anchoring a DPHE model to a Gnielinski/Dittus-Boelter prediction (or to a power law refitted on the base fluid) is therefore well supported by cross-domain evidence but untested for nanofluid double-pipe exchangers.")

# ------------------------------------------------------------------ 15 Benchmarks
H("15. Benchmark Papers")
para("Three benchmarks were selected from Evidence A papers. The first-pass benchmarks P03 and P04 were dropped: P03 is an air/steam constant-wall-temperature configuration validated by random k-fold CV, and P04 reports accuracy below its measurement uncertainty with a high leakage risk.")
for b in BENCHMARKS:
    d = BY[b["id"]]
    H(f"{b['id']} - {d['first_author']} ({d['year']}) - {b['role']}", 2)
    bullet(" " + b["why"], "Why selected:")
    bullet(" " + b["adopt"], "Adopt:")
    bullet(" " + b["not_copy"], "Do not copy:")
    bullet(" " + b["improve"], "What we can improve:")

# ------------------------------------------------------------------ 16 Statistical findings
H("16. Statistical Literature Findings")
para("Evidence base: the 'Statistical Analysis' sheet of the companion workbook (KPI blocks, 18 charts A-R, 14 cross-tabs). The key findings generated from the database are:")
for i, f in enumerate(STATS["findings"]):
    bullet(f" {f}", f"F{i + 1}.")
fig(heat("fam_val", crosstab("families", "val_quality", ["ANN/MLP", "ANFIS", "SVR", "GPR", "Random Forest", "Gradient Boosting", "XGBoost", "Linear regression", "Polynomial regression", "Physics-informed / hybrid"], ["None / unclear", "Random only", "Grouped / external"], rmulti=True), "Model family x validation quality"), "Figure 9. Model family against validation quality.", width=13)
fig(heat("tier_src", crosstab("tier", "src_class", ["Tier 1", "Tier 2", "Tier 3"], ["Experimental", "CFD", "Analytical", "Mixed"]), "Tier x data source"), "Figure 10. Tier against data source.", width=11)
REP = [("Dataset size", "dataset"), ("Exact split", "split"), ("Hyperparameters", "hyper"), ("Scaling", "scaling"), ("Exact inputs", "inputs"),
       ("Exact outputs", "outputs"), ("Predictive uncertainty", "uncertainty"), ("Independent validation", "indep"), ("Extrapolation test", "extrap"), ("Applicability domain", "ad")]
from collections import OrderedDict
fig(hbar("reporting", OrderedDict((a, k(lambda d, kk=kk: d["reporting"][kk])) for a, kk in REP), "Papers reporting each item", sort=False), "Figure 11. Reporting quality.")

# ------------------------------------------------------------------ 17 Gaps
H("17. Research Gaps")
para("Each gap is stated relative to the reviewed literature; absence from this database is not proof of absence elsewhere. Two first-pass gaps were withdrawn: the 'hydraulic penalty left out of models' gap (outside the thermal-only scope) and the claim that 'no ML study exists for MgO-TiO2 hybrid nanofluid in a DPHE' (not verifiable from this database; subsumed in Gap 1).")
table(["#", "Gap (cautious wording)", "Supporting papers", "Priority", "Strength"],
      [[i + 1, g["gap"], g["papers"], g["priority"], g["strength"]] for i, g in enumerate(GAPS)], widths=[0.7, 6.3, 6, 1.6, 2.4], size=8)
para("Final research-gap statement: ", bold_prefix=None)
para("No study was identified in the reviewed literature that (i) predicts the thermal performance of a double-pipe heat exchanger operating with a hybrid nanofluid using verified, fully reported methods; (ii) validates such a model on operating conditions, concentrations or flow levels withheld entirely from training; (iii) reports calibrated predictive uncertainty and applicability-domain warnings; or (iv) anchors the learner to a heat-transfer correlation and derives effectiveness, NTU, heat duty and outlet temperatures from a single learned transport quantity. A study combining these elements on MgO-TiO2 hybrid-nanofluid DPHE data would address several independently documented weaknesses at once.", italic=True)

# ------------------------------------------------------------------ 18 Direction
H("18. Recommended Direction for Our Study")
for a, b in RECOMMENDATION:
    p = doc.add_paragraph()
    p.add_run(a + ". ").bold = True
    p.add_run(b)
    p.paragraph_format.space_after = Pt(4)
para("Thermal input/output architecture:", bold_prefix=None)
table(["Layer", "Content"], [
    ["User inputs", "Th,in; Tc,in; m_h; m_c; nanoparticle concentration (MgO-TiO2, tube side)"],
    ["Property layer", "Nanofluid rho, cp, k, mu at bulk temperature from measured property correlations (base fluid from IAPWS/standard tables)"],
    ["Model features", "Re_t, Pr_t (nanofluid, tube); Re_a, Pr_a (annulus); phi; Cr; optionally ln-transformed"],
    ["Baseline", "U_base from Gnielinski (tube) + annulus correlation + wall resistance; or power law refitted on training folds"],
    ["Learned quantity", "y = ln(U_meas / U_base) (or ln Nu_t ratio if wall temperatures are reliable)"],
    ["Physics layer", "U = U_base * exp(y); NTU = UA/Cmin; eps = f(NTU, Cr) counter-flow; Q = eps*Cmin*(Th,in - Tc,in); Th,out, Tc,out from energy balances"],
    ["Trust layer", "Prediction interval (GPR variance or grouped conformal) propagated to all outputs; applicability-domain flag (range + Mahalanobis)"],
], widths=[3.5, 13], size=8.5)

# ------------------------------------------------------------------ 19 Candidates
H("19. Recommended Model Candidates")
table(["Candidate", "Role", "Settings to start from", "Evidence"], [
    ["Classical correlations (no fit)", "Reference", "Gnielinski / Dittus-Boelter with nanofluid properties", "P13, P29 (correlations), P27 (anchor)"],
    ["Refitted power law (log-linear, ridge)", "Baseline to beat", "ln Nu or ln U on ln Re, ln Pr, phi, ln Cr", "P14 (power law ~ XGBoost), P19, P16 (counter-example)"],
    ["GPR (ARD Matern 5/2 + white noise)", "Strong candidate", "Direct and residual; standardised dimensionless inputs", "P03 (2nd best, best for f), small-data theory"],
    ["SVR (RBF)", "Candidate", "C, epsilon, gamma by inner CV", "P03, P26, P37"],
    ["Random forest", "Reference ensemble", "500 trees, min_samples_leaf >= 3", "P04, P16, P37"],
    ["Shallow gradient boosting", "Candidate", "depth 2-3, lr <= 0.1, strong L2, early stopping", "P37 (good), P14/P15 (deep = overfit)"],
    ["Small MLP ensemble", "Candidate", "1 hidden layer, 4-8 tanh units, weight decay, 10-member ensemble", "P02, P05, P13 (capacity lessons)"],
    ["Not recommended", "-", "Deep nets, ANFIS, PSO/GA-trained networks, CatBoost (no categorical inputs)", "P13 justification only with 2,700 CFD samples"],
], widths=[3.8, 2.6, 5.6, 4.5], size=8)
para("No model is pre-selected: the final choice should be the simplest model whose grouped-CV error is not significantly worse than the best and whose intervals achieve nominal coverage. If GPR and the anchored power law perform equally, the simpler and more extrapolation-robust option should be preferred.")

# ------------------------------------------------------------------ 20 Validation strategy
H("20. Recommended Validation Strategy")
for b in [
    "Unit of independence = one steady-state operating condition (all repeats averaged; pure error estimated from replicates).",
    "Outer loop (reported accuracy): leave-one-condition-out (or grouped K-fold by condition); plus leave-one-concentration-out and leave-one-flow-level-out (hot and cold separately) to measure transfer to unseen levels.",
    "Inner loop: hyperparameters and scaler fitted on outer-training folds only (nested CV); no test information in model selection.",
    "Comparison run: random split reported only to compute a leakage-amplification factor (grouped error / random error) as in P27.",
    "Metrics: RMSE, MAE, MAPE and max error of U (and of derived Q and outlet temperatures) in physical units; % within +/-5% and +/-10%; bias; PI coverage and mean width; all compared with the propagated measurement uncertainty of U.",
    "Final test: 4-8 confirmation experiments chosen after the model is frozen - interior untested combinations and near-boundary points - never used for tuning.",
    "Documentation: model information sheet (data, splits, features, leakage checks), random seeds and code archived.",
]:
    bullet(b)

# ------------------------------------------------------------------ 21 Software
H("21. Recommended Software / Application Concept")
for b in [
    "Inputs: hot and cold inlet temperatures, hot and cold flow rates, concentration; unit selection; validation of ranges.",
    "Engine: frozen scikit-learn pipeline (property layer -> features -> model -> physics layer) serialised with version metadata.",
    "Outputs: U, tube-side h/Nu (if modelled), NTU, effectiveness, heat duty and both outlet temperatures, each with a prediction interval.",
    "Trust features: applicability-domain indicator (green / amber / red), nearest tested condition, model card with validation results, refusal outside the validated envelope.",
    "Form: a small Streamlit (web) or desktop application; no literature evidence of an existing comparable tool was found (P21/P22 only publish network weights).",
]:
    bullet(b)

# ------------------------------------------------------------------ 22 Action plan
H("22. Immediate Action Plan (baby steps)")
steps = [
    ("Data audit", ["List every run: date, fluid batch, concentration, four temperatures, two flows, wall temperatures (if any), pressure data, duration.", "Flag runs with |Qh - Qc|/Qmean above a threshold (e.g., 10%) and non-steady segments.", "Record sensor uncertainties and compute propagated uncertainty of U."]),
    ("Define independent conditions", ["Group readings by set-point (Th,in, Tc,in, m_h, m_c, phi); one condition = one row.", "Count conditions per concentration and per flow level; identify gaps in the factorial coverage."]),
    ("Data-cleaning rules", ["Steady-state criterion (e.g., temperature drift < 0.1 C/min over 5 min).", "Average repeats; keep replicate spread as pure-error estimate.", "Write rules down before modelling; never remove points because the model fits them badly."]),
    ("Thermal target definition", ["Compute U from Qmean and LMTD with the same reference area for all runs.", "Decide whether wall temperatures are reliable enough for tube-side Nu.", "Compute baseline U_base (and Nu_base) with nanofluid properties."]),
    ("Feature definition", ["Implement nanofluid property models (from your measured k, mu if available).", "Compute Re, Pr of each stream, Cr, phi; log-transform where appropriate.", "Exclude any feature derived from outlet temperatures."]),
    ("Baseline model", ["Evaluate un-fitted correlations.", "Fit ridge-regularised power law in log space inside CV."]),
    ("Validation code", ["Implement grouped splitters (condition, concentration, hot-flow level, cold-flow level).", "Nested CV wrapper; scaler inside pipeline; fixed seeds; unit tests for leakage (no group overlap)."]),
    ("Candidate model training", ["Train GPR, SVR, RF, shallow GBR, small MLP ensemble - each direct and residual-on-baseline.", "Keep hyperparameter grids small."]),
    ("Model comparison", ["Tabulate grouped metrics, leakage-amplification factor and paired differences across folds.", "Select the simplest adequate model."]),
    ("Uncertainty analysis", ["GPR intervals; grouped conformal intervals for others.", "Check coverage at 90%/95% on outer folds; recalibrate if needed."]),
    ("Applicability domain", ["Fit range + Mahalanobis gate on training features; set threshold (e.g., 95th percentile of training distances)."]),
    ("Confirmation experiments", ["Freeze model; choose 4-8 new conditions (interior and edge); run them blind; compare with intervals."]),
    ("Model freeze", ["Version the pipeline, data snapshot and metrics; write the model card."]),
    ("Application development", ["Build the input form, prediction + intervals, derived outputs and AD indicator; add export of results."]),
    ("Documentation & reproducibility", ["Repository with code, environment file, data dictionary, model information sheet, and a script that reproduces every reported number."]),
]
for i, (h, items) in enumerate(steps):
    p = doc.add_paragraph()
    p.add_run(f"Step {i + 1} - {h}").bold = True
    for it in items:
        bullet(it)
para("Literature housekeeping (parallel): obtain full texts of P01 (highest priority), P35, P36 and P08 and re-assess them with the same template; verify quartiles on SCImago; rename the PDFs with the supplied scripts; screen X01 and X09 (hybrid-nanofluid ML).")

# ------------------------------------------------------------------ 23 Conclusion
H("23. Final Conclusion")
para("After full-text verification, the published evidence on machine-learning prediction of double-pipe heat-exchanger thermal performance is narrower and less rigorous than its headline accuracies suggest. Experimental DPHE studies are small, mostly use oversized neural networks, validate with random splits of factorial grids, omit uncertainty, and learn physically linked outputs independently. The strongest methodological guidance comes from outside heat exchangers: group-disjoint validation exposes leakage, correlation-anchored and dimensionless models transfer better, and calibrated intervals make predictions usable. Our project should therefore be framed not as 'another ANN with R2 > 0.99' but as a small-data, physics-anchored, uncertainty-aware surrogate for an MgO-TiO2 hybrid-nanofluid DPHE, validated on withheld operating conditions and deployed with explicit domain limits. That framing is supported by the verified literature, addresses documented gaps, and is realistic for a dataset of tens of operating conditions.")

H("Appendix A. Retained papers (24)", 1)
table(["ID", "Tier", "Ev.", "First author (year)", "Journal", "DOI"],
      [[d["id"], d["tier"], d["evidence"], f"{d['first_author']} ({d['year']})", d["journal"], d["doi"]] for d in P], widths=[1, 1.4, 0.8, 3.2, 6, 4.6], size=7.5)
H("Appendix B. Supporting references cited but removed from the active database", 1)
table(["ID", "Reference", "Reason"], [[x["id"], f"{x['authors']} ({x['year']}). {x['title']}. {x['journal']}. DOI {x['doi']}", x["reason"]] for x in REMOVED],
      widths=[1, 12, 4], size=7.5)

doc.save(OUT)
print("saved", OUT)
