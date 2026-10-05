# -*- coding: utf-8 -*-
"""Derived categories and statistics shared by the workbook and the report."""
from collections import Counter, OrderedDict
from papers_data import P

FAMILY_ORDER = ["ANN/MLP", "ANFIS", "RBF network", "SVR", "GPR", "Random Forest", "Decision tree",
                "Gradient Boosting", "XGBoost", "CatBoost", "AdaBoost", "Linear regression",
                "Regularized linear (Ridge/Lasso/ElasticNet)", "Polynomial regression", "Deep learning",
                "Physics-informed / hybrid", "Hybrid metaheuristic (PSO/GA)"]
INPUT_ORDER = ["Reynolds number", "Prandtl number", "Concentration", "Hot flow rate", "Cold flow rate",
               "Flow rate (side unspecified)", "Hot inlet temperature", "Cold inlet temperature",
               "Inlet temperature (side unspecified)", "Thermophysical properties", "Geometry",
               "Dimensionless groups", "Other operating variable"]
OUTPUT_ORDER = ["Nusselt number", "Heat-transfer coefficient", "Overall heat-transfer coefficient",
                "Heat-transfer rate", "Effectiveness", "NTU", "Outlet temperatures", "Entropy/exergy",
                "Thermal performance factor", "Other thermal output"]
VAL_ORDER = ["Random train/test split", "Random k-fold CV", "Grouped validation",
             "Leave-one-condition/group-out", "Leave-one-level-out (concentration/Re)",
             "External validation (separate dataset)", "Independent confirmation run",
             "Unclear / no validation"]

VAL_MULTI = {
    "P02": ["Random train/test split"], "P04": ["Random train/test split", "Random k-fold CV"],
    "P05": ["Unclear / no validation"], "P06": ["Random train/test split"], "P07": ["Random train/test split"],
    "P10": ["Unclear / no validation"], "P01": ["Unclear / no validation"], "P09": ["Unclear / no validation"],
    "P03": ["Random k-fold CV"], "P13": ["Random train/test split", "Random k-fold CV"],
    "P37": ["Random train/test split", "Random k-fold CV"], "P14": ["Random train/test split"],
    "P17": ["Random train/test split"], "P19": ["Unclear / no validation"],
    "P27": ["Random train/test split", "Grouped validation", "Leave-one-condition/group-out"],
    "P28": ["Random train/test split", "External validation (separate dataset)"],
    "P12": ["Random train/test split"], "P15": ["Random train/test split"], "P16": ["Random train/test split"],
    "P18": ["Random train/test split"], "P21": ["Random k-fold CV", "Independent confirmation run"],
    "P22": ["Random train/test split", "Independent confirmation run"],
    "P26": ["Random train/test split", "External validation (separate dataset)"],
    "P29": ["Random train/test split", "Grouped validation"],
}
COMPLEXITY = {"P02": "High", "P04": "High", "P05": "Medium", "P06": "Medium", "P07": "Medium", "P10": "Medium",
              "P01": "Unknown", "P09": "Medium", "P03": "Medium", "P13": "High", "P37": "Medium", "P14": "High",
              "P17": "Medium", "P19": "Low", "P27": "Medium", "P28": "High", "P12": "Medium", "P15": "Medium",
              "P16": "Medium", "P18": "Low", "P21": "Medium", "P22": "Medium", "P26": "Medium", "P29": "High"}

BINS = ["<30", "30-60", "61-120", "121-250", ">250", "Not reported"]


def size_bin(d):
    n = d.get("n_numeric")
    if n is None:
        return "Not reported"
    if n < 30:
        return "<30"
    if n <= 60:
        return "30-60"
    if n <= 120:
        return "61-120"
    if n <= 250:
        return "121-250"
    return ">250"


def acc_level(d):
    r = d.get("headline_r2")
    if r is None:
        return "R/R2 not reported"
    if r >= 0.99:
        return ">=0.99"
    if r >= 0.95:
        return "0.95-0.99"
    return "<0.95"


def val_quality(d):
    v = VAL_MULTI[d["id"]]
    if any(x in v for x in ["Grouped validation", "Leave-one-condition/group-out", "External validation (separate dataset)"]):
        return "Grouped / external"
    if v == ["Unclear / no validation"]:
        return "None / unclear"
    return "Random only"


def src_class(d):
    s = d["data_source"]
    if s.startswith("Mixed"):
        return "Mixed"
    if s.startswith("CFD"):
        return "CFD"
    if s.startswith("Analytical"):
        return "Analytical"
    return "Experimental"


def fam_group(f):
    return f


NP_GROUP = {"P02": "Al2O3", "P05": "Al2O3", "P18": "Al2O3", "P06": "Graphene family (GO/rGO/GNP)",
            "P17": "Graphene family (GO/rGO/GNP)", "P15": "Graphene family (GO/rGO/GNP)",
            "P26": "Graphene family (GO/rGO/GNP)", "P07": "MWCNT", "P09": "MWCNT", "P10": "Fe3O4", "P16": "SiO2",
            "P01": "Hybrid Al2O3-CuO", "P14": "Hybrid Al2O3-CuO", "P37": "Hybrid Al2O3-TiO2",
            "P12": "Several (property-based, analytical)"}


def enrich():
    for d in P:
        d["val_multi"] = VAL_MULTI[d["id"]]
        d["size_bin"] = size_bin(d)
        d["acc_level"] = acc_level(d)
        d["val_quality"] = val_quality(d)
        d["src_class"] = src_class(d)
        d["complexity"] = COMPLEXITY[d["id"]]
        d["uses_ann"] = "Yes" if "ANN/MLP" in d["families"] or "Deep learning" in d["families"] else "No"
        d["uq_flag"] = "Yes" if d["uq"].startswith("Yes") or d["uq"].startswith("Partial") else "No"
        d["phys_flag"] = "Yes" if (d["physics_features"].startswith("Yes") or d["corr_anchored"].startswith("Yes")) else "No"
        d["dimless_flag"] = {"Yes": "Yes", "Partial": "Partial"}.get(d["dimless"], "No / not reported")
        d["indep_flag"] = "Yes" if d["reporting"]["indep"] else "No"
        d["np_group"] = NP_GROUP.get(d["id"], "None (conventional fluid)")
    return P


def count(key, order=None, multi=False, subset=None):
    rows = subset if subset is not None else P
    c = Counter()
    for d in rows:
        v = d[key]
        if multi:
            for x in set(v):
                c[x] += 1
        else:
            c[v] += 1
    if order:
        return OrderedDict((k, c.get(k, 0)) for k in order)
    return OrderedDict(sorted(c.items(), key=lambda kv: (-kv[1], str(kv[0]))))


def crosstab(rkey, ckey, rorder, corder, rmulti=False, cmulti=False):
    t = OrderedDict((r, OrderedDict((c, 0) for c in corder)) for r in rorder)
    for d in P:
        rv = d[rkey] if rmulti else [d[rkey]]
        cv = d[ckey] if cmulti else [d[ckey]]
        for r in set(rv):
            for c in set(cv):
                if r in t and c in t[r]:
                    t[r][c] += 1
    return t


enrich()
N = len(P)
