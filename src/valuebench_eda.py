#!/usr/bin/env python3
"""Reproducible exploratory analysis for the ValueBench clinical-vignette corpus.

The script reads a three-tab Excel export of the Google Sheet, preserves the
source data, constructs one case-level analytic table, and writes publication-
ready figures and CSV tables. It does not edit the Google Sheet.

Example
-------
python valuebench_eda.py \
    --input /path/to/valuebench_eda.xlsx \
    --output-dir outputs

If --input is omitted, the script attempts to read the Google Sheets XLSX
export URL. That fallback works only when the executing environment can access
the source file.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import textwrap
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import patches
from matplotlib.colors import LinearSegmentedColormap, ListedColormap
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import linear_kernel


SPREADSHEET_ID = "1-Zhf4Nr59UqoeJLqhFwLoGDRH8uAuMGKdNqnigVO01Y"
SOURCE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    f"{SPREADSHEET_ID}/edit?gid=220762358"
)
EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    f"{SPREADSHEET_ID}/export?format=xlsx"
)
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parents[1] / "outputs"

PRINCIPLES = ["Autonomy", "Beneficence", "Nonmaleficence", "Justice"]
VALUE_SCORE = {"violates": -1, "neutral": 0, "promotes": 1}
VALUE_LEVELS = list(VALUE_SCORE)
STATUS_ORDER = ["approved", "needs_review", "deprecated"]
STATUS_LABEL = {
    "approved": "Approved",
    "needs_review": "Needs review",
    "deprecated": "Deprecated",
}
STATUS_COLORS = {
    "approved": "#2A9D8F",
    "needs_review": "#E9C46A",
    "deprecated": "#E76F51",
}
DIR_ORDER = ["Choice 1", "Tie", "Choice 2"]
DIR_COLORS = {"Choice 1": "#3B6FB6", "Tie": "#B8BDC7", "Choice 2": "#D97706"}
PROMOTE_COLOR = "#2A9D8F"
NEUTRAL_COLOR = "#B8BDC7"
VIOLATE_COLOR = "#E76F51"
TEXT_DARK = "#1F2937"
GRID_COLOR = "#D9DEE7"

CORE_COLUMNS = [
    "Case ID",
    "Vignette",
    "Choice 1",
    "Autonomy C1",
    "Beneficence C1",
    "Nonmaleficence C1",
    "Justice C1",
    "Choice 2",
    "Autonomy C2",
    "Beneficence C2",
    "Nonmaleficence C2",
    "Justice C2",
]
REVIEW_COLUMNS = [
    "R1",
    "R1 Decision?",
    "R2",
    "R2 Decision?",
    "R3",
    "R3 Decision?",
    "Reviewer Comments",
    "Validation Message",
]
COMPARISON_COLUMNS = [
    "Vignette",
    "Choice 1",
    "Autonomy C1",
    "Beneficence C1",
    "Nonmaleficence C1",
    "Justice C1",
    "Choice 2",
    "Autonomy C2",
    "Beneficence C2",
    "Nonmaleficence C2",
    "Justice C2",
]


# Non-exclusive, transparent keyword flags. These are exploratory descriptors,
# not clinical diagnoses or manually validated specialties.
CONTEXT_PATTERNS: dict[str, list[str]] = {
    "Minor or adolescent": [
        r"\bcapable minors?\b",
        r"\bminor(?:'s|’s)? (?:assent|consent|care|patient|request|decision|wellbeing)\b",
        r"\bminor-care\b",
        r"\badolescent\b",
        r"\bchild\b",
        r"\b(?:1[0-7]|[1-9])-year-old\b",
        r"\bpediatric\b",
    ],
    "Impaired or uncertain capacity": [
        r"lacks? (?:decision-making )?capacity",
        r"capacity (?:cannot|can(?:not|'t)|could not) be (?:established|confirmed)",
        r"capacity (?:is )?(?:uncertain|fluctuat)",
        r"decision-making capacity.*(?:impaired|uncertain)",
        r"\bdelirium\b",
        r"\bvegetative state\b",
        r"\bminimally responsive\b",
        r"\bnot decisional\b",
    ],
    "Surrogate or family decision": [
        r"\bsurrogate\b",
        r"legal decision-maker",
        r"default decision-maker",
        r"\bguardian\b",
        r"parents? (?:consent|request|ask|support|refuse|agree|disagree)",
        r"wife .*requests?",
        r"husband .*requests?",
        r"family .*requests?",
    ],
    "End-of-life or palliative care": [
        r"\bhospice\b",
        r"\bpalliative\b",
        r"comfort[- ](?:only|focused|measures|care)",
        r"life[- ]sustaining",
        r"withdraw(?:al|ing)? (?:the )?(?:ventilator|treatment|support)",
        r"\bDNR\b",
        r"\bdying\b",
        r"terminal(?:ly)? ill",
    ],
    "Mental health or self-harm": [
        r"\bsuicid",
        r"\bself[- ]harm",
        r"\bpsychiatr",
        r"\bdepress",
        r"\bPTSD\b",
        r"\bpanic\b",
        r"\bpsychosis\b",
    ],
    "Reproductive or obstetric care": [
        r"\bpregnan",
        r"\bgestation",
        r"\blabor\b",
        r"\bfetus\b|\bfetal\b",
        r"\bcesarean\b",
        r"\btermination\b",
        r"\bcontracept",
        r"\bsteriliz",
    ],
    "Confidentiality or disclosure": [
        r"\bconfidential",
        r"\bprivacy\b",
        r"\bdisclos",
        r"not (?:tell|inform) (?:her|his|the) parents?",
        r"asks? (?:you )?not to (?:tell|contact|inform)",
        r"keep .* (?:status|result|diagnosis) confidential",
    ],
    "Resource, access, or coverage constraint": [
        r"\binsur",
        r"\bcoverage\b",
        r"\bfunding\b|\bgrant\b",
        r"\bresource",
        r"limited (?:beds|slots|capacity|staff|resources)",
        r"(?:ICU|unit|hospital) is at capacity",
        r"not available locally",
        r"cannot afford",
        r"wait(?:ing)? list",
    ],
    "Research or investigational intervention": [
        r"\binvestigational\b",
        r"\bclinical trial\b",
        r"\btrial at\b",
        r"\bIRB\b",
        r"\bregistry\b",
        r"\bresearch protocol\b",
        r"\bexperimental\b",
    ],
    "Immediate time pressure": [
        r"\bmust decide today\b",
        r"\btonight\b",
        r"within (?:the next )?\d+ (?:hours?|days?)",
        r"\bdeadline\b",
        r"\bthis week\b",
        r"\bimmediately\b",
        r"\bwithin hours\b",
        r"\bby end of day\b",
    ],
}


# Non-exclusive themes for the free-text reviewer comments. The row-level coding
# is exported so the research team can audit or override every assignment.
COMMENT_THEME_PATTERNS: dict[str, list[str]] = {
    "Outside ethical decision scope": [
        r"not a medical (?:decision|question)",
        r"not medically relevant",
        r"(?:more of |a )clinical (?:dilemma|decision)",
        r"legal dilemma",
        r"missing the ethics",
        r"not an ethical",
        r"focuses on clinician conscience",
        r"irrelevant situation",
        r"choice is about values",
        r"^clinical dilemma$",
        r"not a medical dilemma",
    ],
    "No genuine tradeoff or obvious answer": [
        r"not an? (?:ethical )?dilemma",
        r"false dilemma",
        r"lopsided",
        r"super obvious|seems? obvious|straightforward",
        r"no scenario in which",
        r"gross departure from standard of care|failure on boards",
        r"has capacity.*(?:cannot|can't) act against",
        r"if .* meets criteria",
        r"warranted here",
        r"conduct fraud",
        r"no basis for denying",
        r"(?:do not|don't) see a basis for denying",
        r"parents can make an informed decision",
        r"seems like it is the parents?' choice",
        r"reduces? (?:from|the) .*dilemm",
        r"to+o? strong to pose a real dilemma",
        r"why not just wait",
        r"cannot imagine performing",
        r"failure of the pre-op process",
    ],
    "Capacity or autonomy ambiguity": [
        r"\bcapacity\b|\bcapcaitcy\b|\bcompetent\b",
        r"\bautonomy\b|\bautonomous\b",
        r"\bcoerc",
        r"\bminor\b|\bguardian\b",
        r"\bparents?\b",
        r"stated wish",
        r"what the patient wants",
    ],
    "Value attribution uncertainty": [
        r"\bbenefic",
        r"\bnon[- ]?malef",
        r"\bjustice\b",
        r"values? (?:assigned|assignment|alignment|ascription|asigned)",
        r"\bpromot(?:e|es|ed|ing)\b",
        r"\bviolat(?:e|es|ed|ing)\b",
        r"\bneutral\b",
        r"competing ethical values",
        r"map cleanly onto principlism",
        r"which values to whom",
    ],
    "Realism, complexity, or wording": [
        r"\bcomplicated\b|\bconvoluted\b|\bcontrived\b",
        r"\bunlikely\b|\bstrange scenario\b|\bnot realistic\b",
        r"\bbad writing\b|\bconfusing\b",
        r"unclear conflict",
        r"too complicated",
    ],
    "Missing context or stakeholder position": [
        r"more information",
        r"family's position",
        r"have not expressed a preference|not expressed a preference",
        r"what (?:he|she|they) would have",
        r"should(?:'ve| have) been .*discussed",
        r"needs to be better defined",
        r"not clear|isn't clear|unclear",
        r"\bmissing\b",
        r"what is the .* position",
        r"need .* protocol|idea provided for the protocol",
    ],
    "Revision or label correction documented": [
        r"\brevised\b|\bchanged\b|\badded\b|\bdeleted\b",
        r"make .* (?:clearer|sharper|explicit)",
        r"explain the parents?' choice",
    ],
    "Positive or retained with caveat": [
        r"excellent case",
        r"case is fine",
        r"worth seeing",
        r"good case",
        r"go ahead with this case",
        r"interesting since",
    ],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to an XLSX export. If omitted, use the Google Sheets export URL.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for figures, tables, derived data, and summaries.",
    )
    parser.add_argument(
        "--similarity-threshold",
        type=float,
        default=0.75,
        help="TF-IDF cosine threshold for exploratory near-duplicate clusters.",
    )
    return parser.parse_args()


def configure_style() -> None:
    sns.set_theme(style="whitegrid", context="notebook")
    plt.rcParams.update(
        {
            "font.family": "DejaVu Sans",
            "font.size": 10,
            "axes.titlesize": 13,
            "axes.titleweight": "bold",
            "axes.labelsize": 10,
            "axes.edgecolor": GRID_COLOR,
            "axes.linewidth": 0.8,
            "grid.color": GRID_COLOR,
            "grid.linewidth": 0.7,
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "savefig.facecolor": "white",
            "savefig.bbox": "tight",
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
        }
    )


def clean_string_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip()


def normalize_text(value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip().lower()
    return re.sub(r"[^a-z0-9 ]+", "", text)


def word_count(value: object) -> int:
    return len(re.findall(r"\b\w+(?:[-’']\w+)*\b", str(value or "")))


def read_workbook(input_path: Path | None) -> dict[str, pd.DataFrame]:
    source: str | Path = input_path if input_path is not None else EXPORT_URL
    try:
        sheets = pd.read_excel(source, sheet_name=None, dtype=object, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(
            "Could not read the spreadsheet. Export the Google Sheet as XLSX and rerun "
            "with --input /path/to/file.xlsx."
        ) from exc
    required = {"Cases", "Downloaded for mech interp", "Original Cases"}
    missing = required.difference(sheets)
    if missing:
        raise ValueError(f"Missing required tabs: {sorted(missing)}")
    return sheets


def clean_source_tables(
    sheets: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cases = sheets["Cases"].copy().dropna(how="all")
    original = sheets["Original Cases"].copy().dropna(how="all")
    mech = sheets["Downloaded for mech interp"].copy().dropna(how="all")

    for frame in (cases, original, mech):
        frame.columns = [str(c).strip() for c in frame.columns]
        for column in frame.columns:
            if frame[column].dtype == object:
                frame[column] = frame[column].map(
                    lambda x: re.sub(r"\s+", " ", str(x)).strip()
                    if pd.notna(x)
                    else np.nan
                )

    for column in CORE_COLUMNS + [
        "Status",
        "Validation Status",
        "Reviewer Comments",
        "R1 Decision?",
        "R2 Decision?",
        "R3 Decision?",
    ]:
        if column not in cases.columns:
            raise ValueError(f"Cases is missing required column: {column}")

    cases = cases[cases["Case ID"].fillna("").astype(str).str.len().gt(0)].copy()
    original = original[
        original["Case ID"].fillna("").astype(str).str.len().gt(0)
    ].copy()

    uuid_mask = mech["Case ID"].fillna("").astype(str).str.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
    )
    mech_cases = mech[uuid_mask].copy()
    return cases.reset_index(drop=True), original.reset_index(drop=True), mech_cases.reset_index(drop=True)


def validate_source(cases: pd.DataFrame, original: pd.DataFrame) -> list[str]:
    warnings: list[str] = []
    if cases["Case ID"].duplicated().any():
        warnings.append("Cases contains duplicate Case IDs.")
    if original["Case ID"].duplicated().any():
        warnings.append("Original Cases contains duplicate Case IDs.")
    if set(cases["Case ID"]) != set(original["Case ID"]):
        warnings.append("Cases and Original Cases do not contain identical Case ID sets.")

    for column in [f"{p} C{c}" for p in PRINCIPLES for c in (1, 2)]:
        observed = set(clean_string_series(cases[column]).str.lower()) - {""}
        unexpected = observed.difference(VALUE_LEVELS)
        if unexpected:
            warnings.append(f"Unexpected labels in {column}: {sorted(unexpected)}")

    core_missing = cases[CORE_COLUMNS].isna() | cases[CORE_COLUMNS].astype(str).eq("")
    if core_missing.any().any():
        warnings.append("At least one core content or annotation cell is missing.")
    return warnings


def direction_label(diff: float) -> str:
    if diff > 0:
        return "Choice 1"
    if diff < 0:
        return "Choice 2"
    return "Tie"


def match_any(text: str, patterns: Iterable[str]) -> bool:
    return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in patterns)


def build_case_level(cases: pd.DataFrame) -> pd.DataFrame:
    df = cases.copy()
    df["Status"] = clean_string_series(df["Status"]).str.lower()
    df["Validation Status"] = clean_string_series(df["Validation Status"])
    df["Vignette words"] = df["Vignette"].map(word_count)
    df["Choice 1 words"] = df["Choice 1"].map(word_count)
    df["Choice 2 words"] = df["Choice 2"].map(word_count)
    df["Vignette normalized"] = df["Vignette"].map(normalize_text)

    engaged_columns = []
    diff_columns = []
    for principle in PRINCIPLES:
        c1 = clean_string_series(df[f"{principle} C1"]).str.lower()
        c2 = clean_string_series(df[f"{principle} C2"]).str.lower()
        df[f"{principle} C1 score"] = c1.map(VALUE_SCORE)
        df[f"{principle} C2 score"] = c2.map(VALUE_SCORE)
        diff_col = f"{principle} difference"
        df[diff_col] = df[f"{principle} C1 score"] - df[f"{principle} C2 score"]
        df[f"{principle} favors"] = df[diff_col].map(direction_label)
        engaged_col = f"{principle} engaged"
        df[engaged_col] = (c1 != "neutral") | (c2 != "neutral")
        engaged_columns.append(engaged_col)
        diff_columns.append(diff_col)

    df["Engaged principles"] = df[engaged_columns].sum(axis=1).astype(int)
    positive = (df[diff_columns] > 0).sum(axis=1)
    negative = (df[diff_columns] < 0).sum(axis=1)
    df["Tradeoff class"] = np.select(
        [
            (positive > 0) & (negative > 0),
            (positive > 0) & (negative == 0),
            (negative > 0) & (positive == 0),
        ],
        ["Tradeoff", "Choice 1 dominates", "Choice 2 dominates"],
        default="All ties",
    )
    symbol = {"Choice 1": "C1", "Tie": "=", "Choice 2": "C2"}
    df["Tradeoff profile"] = df.apply(
        lambda row: "|".join(symbol[row[f"{p} favors"]] for p in PRINCIPLES), axis=1
    )

    def opposing_set(row: pd.Series) -> str:
        autonomy = row["Autonomy difference"]
        if autonomy == 0:
            return "Autonomy tie"
        opposing = [
            p
            for p in PRINCIPLES[1:]
            if autonomy * row[f"{p} difference"] < 0
        ]
        return " + ".join(opposing) if opposing else "None"

    df["Values opposing autonomy"] = df.apply(opposing_set, axis=1)

    vignette_text = clean_string_series(df["Vignette"])
    for label, patterns in CONTEXT_PATTERNS.items():
        df[f"Context: {label}"] = vignette_text.map(lambda x: match_any(x, patterns))

    return df


def compare_versions(
    case_level: pd.DataFrame, original: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    old = original[["Case ID"] + COMPARISON_COLUMNS].copy()
    merged = case_level.merge(old, on="Case ID", suffixes=("", " original"), how="left")
    changed_cols = []
    for column in COMPARISON_COLUMNS:
        changed = (
            clean_string_series(merged[column])
            != clean_string_series(merged[f"{column} original"])
        )
        merged[f"Changed: {column}"] = changed
        changed_cols.append(f"Changed: {column}")
    merged["Any source change"] = merged[changed_cols].any(axis=1)
    merged["Text changed"] = merged[
        ["Changed: Vignette", "Changed: Choice 1", "Changed: Choice 2"]
    ].any(axis=1)
    label_change_cols = [c for c in changed_cols if c.split(": ", 1)[1] not in {"Vignette", "Choice 1", "Choice 2"}]
    merged["Value label changed"] = merged[label_change_cols].any(axis=1)
    merged["Number of fields changed"] = merged[changed_cols].sum(axis=1)

    tidy_rows = []
    for status in STATUS_ORDER:
        subset = merged[merged["Status"] == status]
        for column in COMPARISON_COLUMNS:
            tidy_rows.append(
                {
                    "Status": STATUS_LABEL[status],
                    "Field": column,
                    "Cases changed": int(subset[f"Changed: {column}"].sum()),
                    "Cases in status": len(subset),
                    "Percent changed": 100 * subset[f"Changed: {column}"].mean()
                    if len(subset)
                    else np.nan,
                }
            )
    return merged, pd.DataFrame(tidy_rows)


class UnionFind:
    def __init__(self, n: int):
        self.parent = list(range(n))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def similarity_analysis(
    case_level: pd.DataFrame, threshold: float
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    texts = clean_string_series(case_level["Vignette"])
    vectorizer = TfidfVectorizer(
        lowercase=True,
        stop_words="english",
        ngram_range=(1, 2),
        min_df=1,
        sublinear_tf=True,
    )
    matrix = vectorizer.fit_transform(texts)
    sim = linear_kernel(matrix, matrix)
    np.fill_diagonal(sim, -1)
    max_index = sim.argmax(axis=1)
    max_value = sim.max(axis=1)

    similarity_rows = pd.DataFrame(
        {
            "Case ID": case_level["Case ID"],
            "Status": case_level["Status"].map(STATUS_LABEL),
            "Most similar Case ID": case_level.iloc[max_index]["Case ID"].to_numpy(),
            "Maximum TF-IDF cosine similarity": max_value,
        }
    )

    pairs = []
    n = len(case_level)
    uf = UnionFind(n)
    for i in range(n):
        for j in range(i + 1, n):
            score = float(sim[i, j])
            if score >= threshold:
                uf.union(i, j)
            if score >= 0.50:
                pairs.append(
                    {
                        "Case ID 1": case_level.iloc[i]["Case ID"],
                        "Status 1": STATUS_LABEL.get(case_level.iloc[i]["Status"], case_level.iloc[i]["Status"]),
                        "Case ID 2": case_level.iloc[j]["Case ID"],
                        "Status 2": STATUS_LABEL.get(case_level.iloc[j]["Status"], case_level.iloc[j]["Status"]),
                        "TF-IDF cosine similarity": score,
                        "Exact normalized vignette match": bool(
                            case_level.iloc[i]["Vignette normalized"]
                            == case_level.iloc[j]["Vignette normalized"]
                        ),
                        "Vignette 1 prefix": str(case_level.iloc[i]["Vignette"])[:180],
                        "Vignette 2 prefix": str(case_level.iloc[j]["Vignette"])[:180],
                    }
                )
    pair_df = pd.DataFrame(pairs).sort_values(
        "TF-IDF cosine similarity", ascending=False
    )

    roots = [uf.find(i) for i in range(n)]
    root_to_members: dict[int, list[int]] = {}
    for i, root in enumerate(roots):
        root_to_members.setdefault(root, []).append(i)
    clusters = []
    cluster_lookup = {}
    cluster_number = 0
    for root, members in sorted(root_to_members.items(), key=lambda x: min(x[1])):
        if len(members) < 2:
            continue
        cluster_number += 1
        cluster_id = f"SIM{cluster_number:02d}"
        for member in members:
            cluster_lookup[member] = cluster_id
            clusters.append(
                {
                    "Similarity cluster": cluster_id,
                    "Case ID": case_level.iloc[member]["Case ID"],
                    "Status": STATUS_LABEL.get(case_level.iloc[member]["Status"], case_level.iloc[member]["Status"]),
                    "Vignette prefix": str(case_level.iloc[member]["Vignette"])[:220],
                }
            )
    cluster_df = pd.DataFrame(clusters)
    similarity_rows["Similarity cluster"] = [cluster_lookup.get(i, "") for i in range(n)]
    return similarity_rows, pair_df, cluster_df


def code_reviewer_comments(case_level: pd.DataFrame) -> pd.DataFrame:
    commented = case_level[
        clean_string_series(case_level["Reviewer Comments"]).ne("")
    ].copy()
    rows = []
    for _, row in commented.iterrows():
        comment = str(row["Reviewer Comments"])
        themes = [
            theme
            for theme, patterns in COMMENT_THEME_PATTERNS.items()
            if match_any(comment, patterns)
        ]
        if not themes:
            themes = ["Unclassified: manual review needed"]
        rows.append(
            {
                "Case ID": row["Case ID"],
                "Status": STATUS_LABEL.get(row["Status"], row["Status"]),
                "R1 decision": row.get("R1 Decision?", ""),
                "R2 decision": row.get("R2 Decision?", ""),
                "Reviewer comment": comment,
                "Themes": " | ".join(themes),
            }
        )
    return pd.DataFrame(rows)


def make_dataset_overview(
    cases: pd.DataFrame,
    original: pd.DataFrame,
    mech_cases: pd.DataFrame,
    version_compare: pd.DataFrame,
) -> pd.DataFrame:
    unique_mech = mech_cases["Case ID"].nunique()
    return pd.DataFrame(
        [
            {
                "Tab": "Cases",
                "Role": "Current working, review, and validation table",
                "Nonblank rows": len(cases),
                "Columns": cases.shape[1],
                "Unique Case IDs": cases["Case ID"].nunique(),
                "Important note": "Canonical current version",
            },
            {
                "Tab": "Original Cases",
                "Role": "Earlier version used for revision audit",
                "Nonblank rows": len(original),
                "Columns": original.shape[1],
                "Unique Case IDs": original["Case ID"].nunique(),
                "Important note": f"{int(version_compare['Any source change'].sum())} cases changed in Cases",
            },
            {
                "Tab": "Downloaded for mech interp",
                "Role": "Derived convenience extract in repeated blocks",
                "Nonblank rows": len(mech_cases),
                "Columns": mech_cases.shape[1],
                "Unique Case IDs": unique_mech,
                "Important note": f"{unique_mech} unique cases; do not count repeated blocks as independent",
            },
        ]
    )


def make_missingness_table(case_level: pd.DataFrame) -> pd.DataFrame:
    display_columns = ["Core content and value labels"] + REVIEW_COLUMNS
    rows = []
    cohort_specs = [("All cases", case_level)] + [
        (STATUS_LABEL[s], case_level[case_level["Status"] == s]) for s in STATUS_ORDER
    ]
    for cohort_name, subset in cohort_specs:
        rows.append(
            {
                "Cohort": cohort_name,
                "Variable": "Core content and value labels",
                "Missing n": int(
                    subset[CORE_COLUMNS]
                    .apply(lambda col: clean_string_series(col).eq(""))
                    .any(axis=1)
                    .sum()
                ),
                "Cohort n": len(subset),
            }
        )
        for column in REVIEW_COLUMNS:
            missing = clean_string_series(subset[column]).eq("")
            rows.append(
                {
                    "Cohort": cohort_name,
                    "Variable": column,
                    "Missing n": int(missing.sum()),
                    "Cohort n": len(subset),
                }
            )
    table = pd.DataFrame(rows)
    table["Missing percent"] = 100 * table["Missing n"] / table["Cohort n"]
    table["Variable"] = pd.Categorical(table["Variable"], display_columns, ordered=True)
    return table.sort_values(["Cohort", "Variable"])


def make_review_transitions(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in case_level.iterrows():
        rows.append(
            {
                "R1 decision": str(row.get("R1 Decision?", "") or "Missing"),
                "R2 decision": str(row.get("R2 Decision?", "") or "Missing"),
                "Final status": STATUS_LABEL.get(row["Status"], row["Status"]),
            }
        )
    table = pd.DataFrame(rows)
    return (
        table.value_counts(["R1 decision", "R2 decision", "Final status"])
        .rename("Cases")
        .reset_index()
        .sort_values("Cases", ascending=False)
    )


def make_value_direction_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    cohort_specs = [("All cases", case_level), ("Approved", case_level[case_level["Status"] == "approved"])]
    for cohort_name, subset in cohort_specs:
        for principle in PRINCIPLES:
            counts = subset[f"{principle} favors"].value_counts()
            for direction in DIR_ORDER:
                n = int(counts.get(direction, 0))
                rows.append(
                    {
                        "Cohort": cohort_name,
                        "Principle": principle,
                        "Direction": direction,
                        "Cases": n,
                        "Cohort n": len(subset),
                        "Percent": 100 * n / len(subset),
                    }
                )
    return pd.DataFrame(rows)


def make_raw_label_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for cohort_name, subset in [
        ("All cases", case_level),
        ("Approved", case_level[case_level["Status"] == "approved"]),
    ]:
        for principle in PRINCIPLES:
            for choice in (1, 2):
                column = f"{principle} C{choice}"
                counts = clean_string_series(subset[column]).str.lower().value_counts()
                for label in VALUE_LEVELS:
                    n = int(counts.get(label, 0))
                    rows.append(
                        {
                            "Cohort": cohort_name,
                            "Principle": principle,
                            "Choice": f"Choice {choice}",
                            "Label": label,
                            "Cases": n,
                            "Cohort n": len(subset),
                            "Percent": 100 * n / len(subset),
                        }
                    )
    return pd.DataFrame(rows)


def make_complexity_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for status in STATUS_ORDER:
        subset = case_level[case_level["Status"] == status]
        for engaged in range(1, 5):
            n = int((subset["Engaged principles"] == engaged).sum())
            rows.append(
                {
                    "Status": STATUS_LABEL[status],
                    "Measure": "Engaged principles",
                    "Category": str(engaged),
                    "Cases": n,
                    "Status n": len(subset),
                    "Percent": 100 * n / len(subset),
                }
            )
        for tradeoff_class, n in subset["Tradeoff class"].value_counts().items():
            rows.append(
                {
                    "Status": STATUS_LABEL[status],
                    "Measure": "Tradeoff class",
                    "Category": tradeoff_class,
                    "Cases": int(n),
                    "Status n": len(subset),
                    "Percent": 100 * n / len(subset),
                }
            )
    return pd.DataFrame(rows)


def make_autonomy_opposition_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for cohort_name, subset in [
        ("All cases", case_level),
        ("Approved", case_level[case_level["Status"] == "approved"]),
    ]:
        directional = subset[subset["Values opposing autonomy"] != "Autonomy tie"]
        counts = directional["Values opposing autonomy"].value_counts()
        for group, n in counts.items():
            rows.append(
                {
                    "Cohort": cohort_name,
                    "Opposing values": group,
                    "Cases": int(n),
                    "Directional-autonomy n": len(directional),
                    "Percent": 100 * n / len(directional) if len(directional) else np.nan,
                }
            )
    return pd.DataFrame(rows).sort_values(["Cohort", "Cases"], ascending=[True, False])


def make_profile_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for cohort_name, subset in [
        ("All cases", case_level),
        ("Approved", case_level[case_level["Status"] == "approved"]),
    ]:
        counts = subset["Tradeoff profile"].value_counts()
        for profile, n in counts.items():
            parts = profile.split("|")
            row = {
                "Cohort": cohort_name,
                "Profile": profile,
                "Cases": int(n),
                "Cohort n": len(subset),
                "Percent": 100 * n / len(subset),
            }
            row.update(dict(zip(PRINCIPLES, parts)))
            rows.append(row)
    return pd.DataFrame(rows).sort_values(["Cohort", "Cases"], ascending=[True, False])


def make_context_table(case_level: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for cohort_name, subset in [
        ("All cases", case_level),
        ("Approved", case_level[case_level["Status"] == "approved"]),
    ]:
        for label in CONTEXT_PATTERNS:
            n = int(subset[f"Context: {label}"].sum())
            rows.append(
                {
                    "Cohort": cohort_name,
                    "Exploratory context flag": label,
                    "Cases": n,
                    "Cohort n": len(subset),
                    "Percent": 100 * n / len(subset),
                    "Definition": " OR ".join(CONTEXT_PATTERNS[label]),
                }
            )
    return pd.DataFrame(rows).sort_values(["Cohort", "Cases"], ascending=[True, False])


def make_comment_theme_table(comment_coding: pd.DataFrame) -> pd.DataFrame:
    exploded = comment_coding.assign(
        Theme=comment_coding["Themes"].str.split(r" \| ", regex=True)
    ).explode("Theme")
    denominators = comment_coding["Status"].value_counts()
    table = (
        exploded.groupby(["Theme", "Status"], dropna=False)
        .size()
        .rename("Commented cases")
        .reset_index()
    )
    table["Comments in status"] = table["Status"].map(denominators)
    table["Percent of commented cases in status"] = (
        100 * table["Commented cases"] / table["Comments in status"]
    )
    return table.sort_values(["Theme", "Commented cases"], ascending=[True, False])


def add_source_note(fig: plt.Figure, extra: str = "") -> None:
    note = f"Source: ValueBench EDA Google Sheet ({SOURCE_URL})."
    if extra:
        note += f" {extra}"
    fig.text(0.01, 0.005, note, ha="left", va="bottom", fontsize=7.5, color="#5B6472")


def save_figure(fig: plt.Figure, figures_dir: Path, stem: str) -> None:
    figures_dir.mkdir(parents=True, exist_ok=True)
    fig.savefig(figures_dir / f"{stem}.png", dpi=300, bbox_inches="tight")
    fig.savefig(figures_dir / f"{stem}.pdf", bbox_inches="tight")
    plt.close(fig)


def annotate_bar_values(ax: plt.Axes, horizontal: bool = False, suffix: str = "") -> None:
    for container in ax.containers:
        try:
            ax.bar_label(container, fmt=lambda x: f"{x:.0f}{suffix}", padding=3, fontsize=8)
        except (AttributeError, ValueError):
            continue


def plot_figure_01(
    case_level: pd.DataFrame,
    version_compare: pd.DataFrame,
    mech_cases: pd.DataFrame,
    figures_dir: Path,
) -> None:
    fig = plt.figure(figsize=(14, 5.8))
    gs = fig.add_gridspec(1, 3, width_ratios=[1.35, 1.0, 1.0], wspace=0.35)

    ax = fig.add_subplot(gs[0, 0])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.set_title("A. Dataset lineage", loc="left")

    def box(x: float, y: float, w: float, h: float, title: str, subtitle: str, color: str) -> None:
        rect = patches.FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.012,rounding_size=0.02",
            facecolor=color, edgecolor="white", linewidth=1.5
        )
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h * 0.62, title, ha="center", va="center", fontsize=11, weight="bold", color="white")
        ax.text(x + w / 2, y + h * 0.30, subtitle, ha="center", va="center", fontsize=8.5, color="white", wrap=True)

    box(0.04, 0.65, 0.40, 0.22, "Original Cases", f"{len(case_level)} case IDs\n13-column source version", "#5B6472")
    box(0.56, 0.65, 0.40, 0.22, "Cases", f"{len(case_level)} case IDs\n23-column working table", "#3B6FB6")
    ax.annotate("", xy=(0.57, 0.76), xytext=(0.43, 0.76), arrowprops=dict(arrowstyle="->", lw=2, color=TEXT_DARK))
    ax.text(0.50, 0.82, f"{int(version_compare['Any source change'].sum())} cases changed", ha="center", va="center", fontsize=8.5)

    y = 0.27
    starts = [0.02, 0.35, 0.68]
    widths = [0.28, 0.28, 0.28]
    for status, x, w in zip(STATUS_ORDER, starts, widths):
        n = int((case_level["Status"] == status).sum())
        box(x, y, w, 0.20, STATUS_LABEL[status], f"{n} cases\n{100*n/len(case_level):.1f}%", STATUS_COLORS[status])
        ax.annotate("", xy=(x + w / 2, y + 0.20), xytext=(0.76, 0.65), arrowprops=dict(arrowstyle="->", lw=1.2, color="#8A93A2", alpha=0.85))

    unique_mech = mech_cases["Case ID"].nunique()
    mech_status = (
        mech_cases.drop_duplicates("Case ID")
        .merge(case_level[["Case ID", "Status"]], on="Case ID", how="left")["Status"]
        .value_counts()
    )
    ax.text(
        0.5,
        0.08,
        f"Mech-interp extract: {unique_mech} unique cases repeated in two blocks\n"
        f"({int(mech_status.get('needs_review', 0))} needs review, {int(mech_status.get('deprecated', 0))} deprecated, "
        f"{int(mech_status.get('approved', 0))} approved)",
        ha="center",
        va="center",
        fontsize=8.5,
        color=TEXT_DARK,
    )

    ax2 = fig.add_subplot(gs[0, 1])
    stage_labels = ["Cases", "R1 decisions", "R2 decisions", "Approved"]
    stage_values = [
        len(case_level),
        clean_string_series(case_level["R1 Decision?"]).ne("").sum(),
        clean_string_series(case_level["R2 Decision?"]).ne("").sum(),
        (case_level["Status"] == "approved").sum(),
    ]
    bars = ax2.barh(stage_labels[::-1], stage_values[::-1], color=[STATUS_COLORS["approved"], "#7FA6D9", "#557FB5", "#344E6C"])
    ax2.set_title("B. Review completion", loc="left")
    ax2.set_xlabel("Cases")
    ax2.set_xlim(0, max(stage_values) * 1.18)
    ax2.grid(axis="y", visible=False)
    for bar, value in zip(bars, stage_values[::-1]):
        ax2.text(bar.get_width() + 5, bar.get_y() + bar.get_height()/2, f"{value}\n({100*value/len(case_level):.1f}%)", va="center", fontsize=9)
    ax2.text(0.98, 0.02, "R3 decisions: 0", transform=ax2.transAxes, ha="right", va="bottom", fontsize=8.5, color="#5B6472")

    ax3 = fig.add_subplot(gs[0, 2])
    status_counts = case_level["Status"].value_counts().reindex(STATUS_ORDER)
    left = 0
    for status in STATUS_ORDER:
        value = int(status_counts[status])
        ax3.barh([0], [value], left=left, color=STATUS_COLORS[status], label=STATUS_LABEL[status])
        ax3.text(left + value / 2, 0, f"{value}\n{100*value/len(case_level):.1f}%", ha="center", va="center", fontsize=9, color=TEXT_DARK, weight="bold")
        left += value
    ax3.set_title("C. Current curation status", loc="left")
    ax3.set_xlim(0, len(case_level))
    ax3.set_yticks([])
    ax3.set_xlabel("Cases")
    ax3.grid(axis="y", visible=False)

    fig.suptitle("Figure 1. The workbook represents one partially completed curation pipeline", x=0.01, ha="left", fontsize=16, weight="bold")
    add_source_note(fig, "Counts use unique Case IDs; repeated mech-interp blocks are not independent observations.")
    fig.subplots_adjust(top=0.82, bottom=0.14)
    save_figure(fig, figures_dir, "figure_01_data_lineage_and_review_funnel")


def plot_figure_02(missingness: pd.DataFrame, figures_dir: Path) -> None:
    cohort_order = ["All cases", "Approved", "Needs review", "Deprecated"]
    variable_order = ["Core content and value labels"] + REVIEW_COLUMNS
    matrix = (
        missingness.pivot(index="Cohort", columns="Variable", values="Missing percent")
        .reindex(index=cohort_order, columns=variable_order)
    )
    display_names = [
        "Core content\n+ value labels",
        "R1\nassigned",
        "R1\ndecision",
        "R2\nassigned",
        "R2\ndecision",
        "R3\nassigned",
        "R3\ndecision",
        "Reviewer\ncomments",
        "Validation\nmessage",
    ]
    fig, ax = plt.subplots(figsize=(13.2, 4.8))
    cmap = LinearSegmentedColormap.from_list("missing", ["#F7FBFF", "#9ECAE1", "#2B6CA3"])
    sns.heatmap(
        matrix,
        cmap=cmap,
        vmin=0,
        vmax=100,
        annot=True,
        fmt=".1f",
        linewidths=1,
        linecolor="white",
        cbar_kws={"label": "Missing (%)"},
        ax=ax,
    )
    ax.set_xticklabels(display_names, rotation=0, ha="center")
    ax.set_yticklabels(ax.get_yticklabels(), rotation=0)
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_title("Figure 2. Missingness is concentrated in conditional review-workflow fields", loc="left", pad=16, fontsize=15)
    ax.text(
        0,
        -0.33,
        "Core case content and all eight value annotations are complete. R2/R3 fields are blank when a case has not reached that stage; validation messages appear mainly for errors.",
        transform=ax.transAxes,
        fontsize=9,
        color=TEXT_DARK,
    )
    add_source_note(fig, "Blank workflow fields are described, not imputed.")
    fig.subplots_adjust(bottom=0.27, top=0.82)
    save_figure(fig, figures_dir, "figure_02_missingness_by_status")


def plot_figure_03(direction_table: pd.DataFrame, figures_dir: Path) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(13.2, 5.6), sharey=True)
    for ax, cohort in zip(axes, ["All cases", "Approved"]):
        subset = direction_table[direction_table["Cohort"] == cohort]
        y = np.arange(len(PRINCIPLES))
        left = np.zeros(len(PRINCIPLES))
        for direction in DIR_ORDER:
            vals = (
                subset[subset["Direction"] == direction]
                .set_index("Principle")
                .reindex(PRINCIPLES)["Percent"]
                .to_numpy()
            )
            ax.barh(y, vals, left=left, color=DIR_COLORS[direction], label=direction, height=0.62)
            for pos, value, lft in zip(y, vals, left):
                if value >= 7:
                    ax.text(lft + value / 2, pos, f"{value:.0f}%", ha="center", va="center", fontsize=8.5, color="white" if direction != "Tie" else TEXT_DARK, weight="bold")
            left += vals
        ax.set_yticks(y, PRINCIPLES)
        ax.invert_yaxis()
        ax.set_xlim(0, 100)
        ax.set_xlabel("Cases (%)")
        ax.set_title(f"{cohort} (n={int(subset['Cohort n'].iloc[0])})", loc="left")
        ax.grid(axis="y", visible=False)
    axes[1].legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=3, frameon=False)
    fig.suptitle("Figure 3. Choice position is associated with which principle is favored", x=0.01, ha="left", fontsize=15, weight="bold")
    fig.text(0.01, 0.075, "Direction is based on the ordered comparison promotes > neutral > violates; it is not a summed ethical score.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(top=0.82, bottom=0.24, wspace=0.22)
    save_figure(fig, figures_dir, "figure_03_value_alignment_and_choice_position")


def plot_figure_04(complexity: pd.DataFrame, figures_dir: Path) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(13.2, 5.5))
    status_labels = [STATUS_LABEL[s] for s in STATUS_ORDER]
    engaged = complexity[complexity["Measure"] == "Engaged principles"]
    x = np.arange(len(status_labels))
    width = 0.19
    for i, count in enumerate([1, 2, 3, 4]):
        vals = (
            engaged[engaged["Category"] == str(count)]
            .set_index("Status")
            .reindex(status_labels)["Percent"]
            .fillna(0)
            .to_numpy()
        )
        axes[0].bar(x + (i - 1.5) * width, vals, width, label=str(count), color=sns.color_palette("Blues", 6)[i + 2])
    axes[0].set_xticks(x, status_labels)
    axes[0].set_ylabel("Cases within status (%)")
    axes[0].set_title("A. Number of engaged principles", loc="left")
    axes[0].legend(title="Engaged", frameon=False, ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.16))

    trade = complexity[complexity["Measure"] == "Tradeoff class"]
    classes = ["Tradeoff", "Choice 1 dominates", "Choice 2 dominates", "All ties"]
    class_colors = ["#3B6FB6", "#D97706", "#E76F51", "#B8BDC7"]
    bottom = np.zeros(len(status_labels))
    for category, color in zip(classes, class_colors):
        vals = (
            trade[trade["Category"] == category]
            .set_index("Status")
            .reindex(status_labels)["Percent"]
            .fillna(0)
            .to_numpy()
        )
        axes[1].bar(x, vals, bottom=bottom, label=category, color=color, width=0.62)
        for xpos, val, btm in zip(x, vals, bottom):
            if val >= 5:
                axes[1].text(xpos, btm + val / 2, f"{val:.0f}%", ha="center", va="center", fontsize=8.5, color="white", weight="bold")
        bottom += vals
    axes[1].set_xticks(x, status_labels)
    axes[1].set_ylabel("Cases within status (%)")
    axes[1].set_title("B. Does each choice win on at least one principle?", loc="left")
    axes[1].legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=2)
    for ax in axes:
        ax.set_ylim(0, 105)
        ax.grid(axis="x", visible=False)
    fig.suptitle("Figure 4. Approved cases are multi-principle, non-dominated tradeoffs", x=0.01, ha="left", fontsize=15, weight="bold")
    add_source_note(fig, "A principle is engaged when either choice is non-neutral for that principle.")
    fig.subplots_adjust(top=0.82, bottom=0.28, wspace=0.28)
    save_figure(fig, figures_dir, "figure_04_ethical_complexity_and_validity")


def plot_figure_05(opposition: pd.DataFrame, figures_dir: Path) -> None:
    all_data = opposition[opposition["Cohort"] == "All cases"].set_index("Opposing values")
    approved = opposition[opposition["Cohort"] == "Approved"].set_index("Opposing values")
    order = approved["Percent"].sort_values(ascending=True).index.tolist()
    for extra in all_data.index:
        if extra not in order:
            order.insert(0, extra)
    display = pd.DataFrame(
        {
            "All cases": all_data.reindex(order)["Percent"].fillna(0),
            "Approved": approved.reindex(order)["Percent"].fillna(0),
        }
    )
    y = np.arange(len(display))
    fig, ax = plt.subplots(figsize=(11.8, 6.6))
    ax.scatter(display["All cases"], y + 0.12, s=70, color="#7FA6D9", label="All cases")
    ax.scatter(display["Approved"], y - 0.12, s=70, color=STATUS_COLORS["approved"], label="Approved")
    for ypos, a, b in zip(y, display["All cases"], display["Approved"]):
        ax.plot([a, b], [ypos + 0.12, ypos - 0.12], color="#CCD3DD", lw=1)
    ax.set_yticks(y, display.index)
    ax.set_xlabel("Cases with directional autonomy annotation (%)")
    ax.set_xlim(left=0)
    ax.set_title("Figure 5. Autonomy most often competes with beneficence and/or nonmaleficence", loc="left", pad=15, fontsize=15)
    ax.legend(frameon=False, loc="lower right")
    ax.grid(axis="y", visible=False)
    fig.text(0.01, 0.055, "Cases where the two choices tie on autonomy are excluded from the denominator. Categories show which other principles favor the opposite choice.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(left=0.31, bottom=0.16, top=0.86)
    save_figure(fig, figures_dir, "figure_05_values_opposing_autonomy")


def plot_figure_06(profile_table: pd.DataFrame, figures_dir: Path) -> None:
    approved = profile_table[profile_table["Cohort"] == "Approved"].nlargest(12, "Cases").copy()
    approved = approved.sort_values("Cases", ascending=True).reset_index(drop=True)
    code = {"C1": 1, "=": 0, "C2": -1}
    matrix = np.vectorize(code.get)(approved[PRINCIPLES].to_numpy()).astype(float)
    fig = plt.figure(figsize=(12.8, 7.0))
    gs = fig.add_gridspec(1, 2, width_ratios=[2.3, 1.0], wspace=0.08)
    ax = fig.add_subplot(gs[0, 0])
    cmap = ListedColormap([DIR_COLORS["Choice 2"], DIR_COLORS["Tie"], DIR_COLORS["Choice 1"]])
    sns.heatmap(
        matrix,
        cmap=cmap,
        vmin=-1,
        vmax=1,
        cbar=False,
        linewidths=1,
        linecolor="white",
        xticklabels=PRINCIPLES,
        yticklabels=[f"Profile {i+1}" for i in range(len(approved))],
        ax=ax,
    )
    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            label = "C1" if matrix[i, j] > 0 else "C2" if matrix[i, j] < 0 else "="
            ax.text(j + 0.5, i + 0.5, label, ha="center", va="center", color="white" if label != "=" else TEXT_DARK, weight="bold", fontsize=9)
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_title("A. Which choice each principle favors", loc="left")

    ax2 = fig.add_subplot(gs[0, 1], sharey=ax)
    bars = ax2.barh(np.arange(len(approved)) + 0.5, approved["Cases"], height=0.75, color="#3B6FB6")
    ax2.set_ylim(matrix.shape[0], 0)
    ax2.set_yticks([])
    ax2.set_xlabel("Approved cases")
    ax2.set_title("B. Frequency", loc="left")
    ax2.grid(axis="y", visible=False)
    for bar, n, pct in zip(bars, approved["Cases"], approved["Percent"]):
        ax2.text(bar.get_width() + 0.15, bar.get_y() + bar.get_height()/2, f"{int(n)} ({pct:.1f}%)", va="center", fontsize=8.5)
    ax2.set_xlim(0, max(approved["Cases"]) * 1.55)
    fig.suptitle("Figure 6. The approved corpus contains many sparse tradeoff profiles", x=0.01, ha="left", fontsize=15, weight="bold")
    fig.text(0.01, 0.055, "Profiles are ordered by frequency among approved cases. C1/C2 indicate the more aligned choice for that principle; '=' indicates a tie.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(top=0.84, bottom=0.14, left=0.12, right=0.97)
    save_figure(fig, figures_dir, "figure_06_top_tradeoff_profiles")


def plot_figure_07(context: pd.DataFrame, figures_dir: Path) -> None:
    all_data = context[context["Cohort"] == "All cases"].set_index("Exploratory context flag")
    approved = context[context["Cohort"] == "Approved"].set_index("Exploratory context flag")
    order = approved["Percent"].sort_values(ascending=True).index
    y = np.arange(len(order))
    fig, ax = plt.subplots(figsize=(12.0, 7.0))
    ax.scatter(all_data.reindex(order)["Percent"], y + 0.12, s=70, color="#7FA6D9", label="All cases")
    ax.scatter(approved.reindex(order)["Percent"], y - 0.12, s=70, color=STATUS_COLORS["approved"], label="Approved")
    for ypos, a, b in zip(y, all_data.reindex(order)["Percent"], approved.reindex(order)["Percent"]):
        ax.plot([a, b], [ypos + 0.12, ypos - 0.12], color="#CCD3DD", lw=1)
    ax.set_yticks(y, order)
    ax.set_xlabel("Cases with context flag (%)")
    ax.set_title("Figure 7. The approved subset emphasizes several recurring clinical contexts", loc="left", pad=15, fontsize=15)
    ax.legend(frameon=False, loc="lower right")
    ax.grid(axis="y", visible=False)
    fig.text(0.01, 0.055, "Flags are non-exclusive and keyword-derived from vignette text. They are exploratory coverage checks, not validated clinical-domain labels.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(left=0.30, bottom=0.15, top=0.87)
    save_figure(fig, figures_dir, "figure_07_exploratory_case_contexts")


def plot_figure_08(theme_table: pd.DataFrame, figures_dir: Path) -> None:
    statuses = ["Deprecated", "Needs review", "Approved"]
    theme_order = (
        theme_table.groupby("Theme")["Commented cases"].sum().sort_values(ascending=True).index
    )
    plot_df = (
        theme_table.pivot(index="Theme", columns="Status", values="Percent of commented cases in status")
        .reindex(index=theme_order, columns=statuses)
        .fillna(0)
    )
    y = np.arange(len(plot_df))
    width = 0.23
    fig, ax = plt.subplots(figsize=(12.5, 7.0))
    colors = [STATUS_COLORS["deprecated"], STATUS_COLORS["needs_review"], STATUS_COLORS["approved"]]
    for i, (status, color) in enumerate(zip(statuses, colors)):
        ax.barh(y + (i - 1) * width, plot_df[status], height=width, label=status, color=color)
    ax.set_yticks(y, plot_df.index)
    ax.set_xlabel("Commented cases within status (%)")
    ax.set_title("Figure 8. Reviewer comments identify recurring curation failure modes", loc="left", pad=15, fontsize=15)
    ax.legend(frameon=False, loc="lower right")
    ax.grid(axis="y", visible=False)
    fig.text(0.01, 0.055, "Themes are non-exclusive and rule-coded from 79 comments. Missing comments do not mean no problem; review the exported row-level coding before submission.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(left=0.31, bottom=0.15, top=0.87)
    save_figure(fig, figures_dir, "figure_08_reviewer_comment_themes")


def plot_figure_09(
    case_level: pd.DataFrame,
    similarity_rows: pd.DataFrame,
    pairs: pd.DataFrame,
    figures_dir: Path,
) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(13.2, 5.6))
    plot_df = case_level.copy()
    plot_df["Status label"] = plot_df["Status"].map(STATUS_LABEL)
    sns.boxplot(
        data=plot_df,
        x="Status label",
        y="Vignette words",
        order=[STATUS_LABEL[s] for s in STATUS_ORDER],
        palette=[STATUS_COLORS[s] for s in STATUS_ORDER],
        width=0.6,
        showfliers=False,
        hue="Status label",
        legend=False,
        ax=axes[0],
    )
    sns.stripplot(
        data=plot_df,
        x="Status label",
        y="Vignette words",
        order=[STATUS_LABEL[s] for s in STATUS_ORDER],
        color="#334155",
        alpha=0.28,
        size=2.4,
        jitter=0.22,
        ax=axes[0],
    )
    axes[0].set_title("A. Vignette length", loc="left")
    axes[0].set_xlabel("")
    axes[0].set_ylabel("Words")

    axes[1].hist(
        similarity_rows["Maximum TF-IDF cosine similarity"],
        bins=np.linspace(0, 1, 21),
        color="#3B6FB6",
        edgecolor="white",
    )
    axes[1].axvline(0.75, color=VIOLATE_COLOR, linestyle="--", linewidth=1.5, label="Exploratory cluster threshold = 0.75")
    exact_pairs = int(pairs["Exact normalized vignette match"].sum()) if len(pairs) else 0
    axes[1].set_title("B. Closest other vignette", loc="left")
    axes[1].set_xlabel("Maximum TF-IDF cosine similarity")
    axes[1].set_ylabel("Cases")
    axes[1].legend(frameon=False, fontsize=8.5)
    axes[1].text(0.98, 0.96, f"Exact duplicate pairs: {exact_pairs}", transform=axes[1].transAxes, ha="right", va="top", fontsize=9, color=TEXT_DARK)
    fig.suptitle("Figure 9. Text audit identifies length variation and potential split leakage", x=0.01, ha="left", fontsize=15, weight="bold")
    fig.text(0.01, 0.055, "TF-IDF similarity is a screening tool. High-similarity pairs require manual review before assigning group-aware train/validation splits.", fontsize=9, color=TEXT_DARK)
    add_source_note(fig)
    fig.subplots_adjust(top=0.82, bottom=0.17, wspace=0.28)
    save_figure(fig, figures_dir, "figure_09_text_length_and_similarity")


def make_contact_sheet(figures_dir: Path) -> None:
    from PIL import Image, ImageDraw, ImageFont

    paths = sorted(figures_dir.glob("figure_*.png"))
    if not paths:
        return
    thumb_w, thumb_h = 720, 390
    margin, label_h = 28, 48
    cols = 2
    rows = math.ceil(len(paths) / cols)
    canvas = Image.new("RGB", (cols * (thumb_w + margin) + margin, rows * (thumb_h + label_h + margin) + margin), "white")
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default()
    for index, path in enumerate(paths):
        image = Image.open(path).convert("RGB")
        image.thumbnail((thumb_w, thumb_h))
        x = margin + (index % cols) * (thumb_w + margin)
        y = margin + (index // cols) * (thumb_h + label_h + margin)
        canvas.paste(image, (x, y + label_h))
        label = path.stem.replace("_", " ")
        draw.text((x, y + 10), label, fill=TEXT_DARK, font=font)
    canvas.save(figures_dir / "figures_contact_sheet.png", quality=95)


def write_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)


def markdown_table(df: pd.DataFrame, max_rows: int = 40) -> str:
    shown = df.head(max_rows).copy()
    columns = [str(c) for c in shown.columns]
    lines = ["| " + " | ".join(columns) + " |", "| " + " | ".join(["---"] * len(columns)) + " |"]
    for _, row in shown.iterrows():
        values = []
        for value in row:
            if isinstance(value, float):
                text = f"{value:.1f}" if not np.isnan(value) else ""
            else:
                text = str(value)
            values.append(text.replace("|", "\\|").replace("\n", " "))
        lines.append("| " + " | ".join(values) + " |")
    if len(df) > max_rows:
        lines.append(f"\n_First {max_rows} of {len(df)} rows shown; see CSV for the complete table._")
    return "\n".join(lines)


def write_key_tables_markdown(tables: dict[str, pd.DataFrame], path: Path) -> None:
    selections = [
        ("Dataset overview", "table_01_dataset_overview"),
        ("Current curation status", "table_02_status_counts"),
        ("Review and validation completeness", "table_03_workflow_completeness"),
        ("Value direction by cohort", "table_05_value_direction_by_cohort"),
        ("Ethical complexity", "table_07_complexity_and_tradeoff_class"),
        ("Values opposing autonomy", "table_08_values_opposing_autonomy"),
        ("Exploratory context flags", "table_10_exploratory_context_flags"),
        ("Reviewer comment themes", "table_11_reviewer_comment_themes"),
        ("Text and duplication summary", "table_13_text_and_duplication_summary"),
    ]
    parts = [
        "# ValueBench EDA tables",
        "",
        f"Source: [{SOURCE_URL}]({SOURCE_URL})",
        "",
        "Percentages are descriptive of this corpus and should not be interpreted as population estimates.",
    ]
    for title, key in selections:
        parts.extend(["", f"## {title}", "", markdown_table(tables[key])])
    path.write_text("\n".join(parts), encoding="utf-8")


def build_tables(
    cases: pd.DataFrame,
    original: pd.DataFrame,
    mech_cases: pd.DataFrame,
    case_level: pd.DataFrame,
    version_compare: pd.DataFrame,
    changes_by_field: pd.DataFrame,
    similarity_rows: pd.DataFrame,
    similarity_pairs: pd.DataFrame,
    clusters: pd.DataFrame,
    comment_coding: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    status_counts = (
        case_level["Status"]
        .value_counts()
        .reindex(STATUS_ORDER)
        .rename_axis("Status")
        .reset_index(name="Cases")
    )
    status_counts["Status"] = status_counts["Status"].map(STATUS_LABEL)
    status_counts["Percent"] = 100 * status_counts["Cases"] / len(case_level)

    workflow = pd.DataFrame(
        [
            {
                "Field or stage": "Case content present",
                "Complete n": int(
                    (~case_level[CORE_COLUMNS].apply(lambda col: clean_string_series(col).eq("")).any(axis=1)).sum()
                ),
            },
            {"Field or stage": "R1 assigned", "Complete n": int(clean_string_series(case_level["R1"]).ne("").sum())},
            {"Field or stage": "R1 decision", "Complete n": int(clean_string_series(case_level["R1 Decision?"]).ne("").sum())},
            {"Field or stage": "R2 assigned", "Complete n": int(clean_string_series(case_level["R2"]).ne("").sum())},
            {"Field or stage": "R2 decision", "Complete n": int(clean_string_series(case_level["R2 Decision?"]).ne("").sum())},
            {"Field or stage": "R3 assigned", "Complete n": int(clean_string_series(case_level["R3"]).ne("").sum())},
            {"Field or stage": "R3 decision", "Complete n": int(clean_string_series(case_level["R3 Decision?"]).ne("").sum())},
            {"Field or stage": "Reviewer comment present", "Complete n": int(clean_string_series(case_level["Reviewer Comments"]).ne("").sum())},
            {"Field or stage": "Automated validation valid", "Complete n": int((case_level["Validation Status"] == "✅ Valid").sum())},
            {"Field or stage": "Automated validation error", "Complete n": int((case_level["Validation Status"] == "❌ Error").sum())},
        ]
    )
    workflow["Total cases"] = len(case_level)
    workflow["Percent of cases"] = 100 * workflow["Complete n"] / len(case_level)

    duplicate_groups = (
        case_level.groupby("Vignette normalized")
        .filter(lambda x: len(x) > 1)
        .sort_values("Vignette normalized")
        [["Case ID", "Status", "Vignette"]]
        .copy()
    )
    if not duplicate_groups.empty:
        duplicate_groups["Status"] = duplicate_groups["Status"].map(STATUS_LABEL)

    text_summary = (
        case_level.assign(Status=case_level["Status"].map(STATUS_LABEL))
        .groupby("Status", observed=True)[["Vignette words", "Choice 1 words", "Choice 2 words"]]
        .agg(["count", "min", "median", "max"])
    )
    text_summary.columns = [" ".join(col).strip() for col in text_summary.columns]
    text_summary = text_summary.reset_index()
    duplication_summary = pd.DataFrame(
        [
            {"Metric": "Unique Case IDs", "Value": case_level["Case ID"].nunique()},
            {"Metric": "Unique normalized vignettes", "Value": case_level["Vignette normalized"].nunique()},
            {"Metric": "Exact duplicate pairs", "Value": int(similarity_pairs["Exact normalized vignette match"].sum()) if len(similarity_pairs) else 0},
            {"Metric": "Cases in TF-IDF similarity clusters", "Value": clusters["Case ID"].nunique() if len(clusters) else 0},
            {"Metric": "Median maximum similarity", "Value": similarity_rows["Maximum TF-IDF cosine similarity"].median()},
            {"Metric": "Maximum non-self similarity", "Value": similarity_rows["Maximum TF-IDF cosine similarity"].max()},
        ]
    )
    text_dup_summary = pd.concat(
        [
            text_summary.assign(Section="Text length by status").rename(columns={"Status": "Metric"}),
            duplication_summary.assign(Section="Duplication and similarity"),
        ],
        ignore_index=True,
        sort=False,
    )

    version_summary = (
        version_compare.groupby("Status")
        .agg(
            Cases=("Case ID", "size"),
            Any_change=("Any source change", "sum"),
            Text_change=("Text changed", "sum"),
            Value_label_change=("Value label changed", "sum"),
        )
        .reindex(STATUS_ORDER)
        .reset_index()
    )
    version_summary["Status"] = version_summary["Status"].map(STATUS_LABEL)
    for column in ["Any_change", "Text_change", "Value_label_change"]:
        version_summary[f"{column}_percent"] = 100 * version_summary[column] / version_summary["Cases"]

    tables = {
        "table_01_dataset_overview": make_dataset_overview(cases, original, mech_cases, version_compare),
        "table_02_status_counts": status_counts,
        "table_03_workflow_completeness": workflow,
        "table_04_missingness_by_status": make_missingness_table(case_level),
        "table_05_value_direction_by_cohort": make_value_direction_table(case_level),
        "table_06_raw_value_labels": make_raw_label_table(case_level),
        "table_07_complexity_and_tradeoff_class": make_complexity_table(case_level),
        "table_08_values_opposing_autonomy": make_autonomy_opposition_table(case_level),
        "table_09_tradeoff_profiles": make_profile_table(case_level),
        "table_10_exploratory_context_flags": make_context_table(case_level),
        "table_11_reviewer_comment_themes": make_comment_theme_table(comment_coding),
        "table_12_version_changes_by_field": changes_by_field,
        "table_13_text_and_duplication_summary": text_dup_summary,
        "table_14_review_transitions": make_review_transitions(case_level),
        "table_15_validation_errors": case_level[
            case_level["Validation Status"] == "❌ Error"
        ][["Case ID", "Status", "Validation Message"]].assign(
            Status=lambda x: x["Status"].map(STATUS_LABEL)
        ),
        "table_16_exact_duplicate_cases": duplicate_groups,
        "table_17_similarity_pairs_manual_review": similarity_pairs,
        "table_18_similarity_clusters": clusters,
        "table_19_reviewer_comment_coding_review": comment_coding,
        "table_20_version_change_summary": version_summary,
    }
    return tables


def write_readme(output_dir: Path, warnings: list[str], summary: dict[str, object]) -> None:
    warning_text = "\n".join(f"- {w}" for w in warnings) if warnings else "- No structural warnings were raised."
    content = f"""# ValueBench exploratory data analysis outputs

This folder contains the reproducible Python analysis, nine report-ready figures, and supporting tables for the ValueBench clinical-vignette corpus.

## Run the analysis

```bash
python src/valuebench_eda.py --input /path/to/valuebench_eda.xlsx --output-dir outputs
```

The script does not edit the Google Sheet. It reads an XLSX export, preserves the three-tab data lineage, and writes derived files locally.

## Primary analytic choices

- All {summary['total_cases']} cases describe the curation pipeline and candidate corpus.
- The {summary['approved_cases']} approved cases are treated as the currently training-ready cohort.
- `Original Cases` is used only for revision auditing.
- The mech-interp tab is a repeated extract, not an independent dataset.
- `promotes`, `neutral`, and `violates` are ordered only to determine which choice each principle favors. The four principles are never summed into a single ethical score.
- Reviewer-comment themes and vignette context flags are transparent, non-exclusive, rule-based exploratory codes. Their row-level coding is exported for manual review.
- TF-IDF similarity flags possible duplicates or paraphrases; it does not establish semantic equivalence.

## Output structure

- `figures/`: high-resolution PNG and vector PDF versions of Figures 1–9, plus a contact sheet.
- `tables/`: CSVs supporting every figure and audit.
- `derived_case_level.csv`: one row per current case with derived variables and exploratory flags.
- `analysis_summary.json`: headline counts used to draft the report.
- `key_tables.md`: compact rendering of the most important tables.

## Source checks

{warning_text}

Source: {SOURCE_URL}
"""
    (output_dir / "README.md").write_text(content, encoding="utf-8")


def main() -> None:
    args = parse_args()
    output_dir = args.output_dir.resolve()
    figures_dir = output_dir / "figures"
    tables_dir = output_dir / "tables"
    output_dir.mkdir(parents=True, exist_ok=True)
    figures_dir.mkdir(parents=True, exist_ok=True)
    tables_dir.mkdir(parents=True, exist_ok=True)

    configure_style()
    sheets = read_workbook(args.input)
    cases, original, mech_cases = clean_source_tables(sheets)
    warnings = validate_source(cases, original)
    case_level = build_case_level(cases)
    version_compare, changes_by_field = compare_versions(case_level, original)
    similarity_rows, similarity_pairs, clusters = similarity_analysis(
        case_level, args.similarity_threshold
    )
    comment_coding = code_reviewer_comments(case_level)

    tables = build_tables(
        cases,
        original,
        mech_cases,
        case_level,
        version_compare,
        changes_by_field,
        similarity_rows,
        similarity_pairs,
        clusters,
        comment_coding,
    )
    for stem, table in tables.items():
        write_csv(table, tables_dir / f"{stem}.csv")

    # Append similarity results and version-audit flags to the case-level file.
    derived = version_compare.copy()
    derived = derived.merge(similarity_rows, on="Case ID", how="left", suffixes=("", " similarity"))
    original_columns = [c for c in derived.columns if c.endswith(" original")]
    # Do not duplicate the full earlier vignette text in the delivered analytic table.
    derived = derived.drop(columns=original_columns)
    # Reviewer account identifiers are not needed to reproduce any analysis and
    # are omitted from the public case-level export. Review decisions and
    # comments remain available for workflow and qualitative analysis.
    reviewer_identity_columns = [c for c in ["R1", "R2", "R3"] if c in derived.columns]
    derived = derived.drop(columns=reviewer_identity_columns)
    write_csv(derived, output_dir / "derived_case_level.csv")

    missingness = tables["table_04_missingness_by_status"]
    direction = tables["table_05_value_direction_by_cohort"]
    complexity = tables["table_07_complexity_and_tradeoff_class"]
    opposition = tables["table_08_values_opposing_autonomy"]
    profiles = tables["table_09_tradeoff_profiles"]
    context = tables["table_10_exploratory_context_flags"]
    themes = tables["table_11_reviewer_comment_themes"]

    plot_figure_01(case_level, version_compare, mech_cases, figures_dir)
    plot_figure_02(missingness, figures_dir)
    plot_figure_03(direction, figures_dir)
    plot_figure_04(complexity, figures_dir)
    plot_figure_05(opposition, figures_dir)
    plot_figure_06(profiles, figures_dir)
    plot_figure_07(context, figures_dir)
    plot_figure_08(themes, figures_dir)
    plot_figure_09(case_level, similarity_rows, similarity_pairs, figures_dir)
    make_contact_sheet(figures_dir)

    summary = {
        "analysis_date": date.today().isoformat(),
        "source_url": SOURCE_URL,
        "total_cases": int(len(case_level)),
        "approved_cases": int((case_level["Status"] == "approved").sum()),
        "needs_review_cases": int((case_level["Status"] == "needs_review").sum()),
        "deprecated_cases": int((case_level["Status"] == "deprecated").sum()),
        "validation_valid": int((case_level["Validation Status"] == "✅ Valid").sum()),
        "validation_errors": int((case_level["Validation Status"] == "❌ Error").sum()),
        "r1_decisions": int(clean_string_series(case_level["R1 Decision?"]).ne("").sum()),
        "r2_decisions": int(clean_string_series(case_level["R2 Decision?"]).ne("").sum()),
        "r3_decisions": int(clean_string_series(case_level["R3 Decision?"]).ne("").sum()),
        "commented_cases": int(clean_string_series(case_level["Reviewer Comments"]).ne("").sum()),
        "cases_changed_from_original": int(version_compare["Any source change"].sum()),
        "vignettes_changed_from_original": int(version_compare["Changed: Vignette"].sum()),
        "unique_mech_cases": int(mech_cases["Case ID"].nunique()),
        "mech_rows_after_header_removal": int(len(mech_cases)),
        "exact_duplicate_pairs": int(similarity_pairs["Exact normalized vignette match"].sum()) if len(similarity_pairs) else 0,
        "similarity_cluster_threshold": args.similarity_threshold,
        "similarity_clusters": int(clusters["Similarity cluster"].nunique()) if len(clusters) else 0,
        "warnings": warnings,
    }
    (output_dir / "analysis_summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    write_key_tables_markdown(tables, output_dir / "key_tables.md")
    write_readme(output_dir, warnings, summary)

    print(json.dumps(summary, indent=2))
    print(f"Wrote {len(list(figures_dir.glob('figure_*.png')))} PNG figures and {len(tables)} CSV tables to {output_dir}")


if __name__ == "__main__":
    main()
