import io
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

import streamlit as st

# -------------------------------------------------
# KONFIG
# -------------------------------------------------

st.set_page_config(page_title="DGE-rapport", layout="wide")

REGION_POPULATION = {
    "Region Nordjylland": 590000,
    "Region Midtjylland": 1330000,
    "Region Syddanmark": 1220000,
    "Region Hovedstaden": 1870000,
    "Region Sjælland": 840000,
}

TEST_PATTERNS = ["test", "demo", "euv"]


# -------------------------------------------------
# HJÆLPEFUNKTIONER
# -------------------------------------------------

def parse_date_series(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def load_excel(uploaded_file):
    if uploaded_file is None:
        return None
    return pd.read_excel(uploaded_file)


def clean_groups_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    if "Dato for arkivering" in df.columns:
        df["Dato for arkivering"] = df["Dato for arkivering"].replace("-", np.nan)
        df["Dato for arkivering"] = parse_date_series(df["Dato for arkivering"])
    if "Antal medlemmer" in df.columns:
        df["Antal medlemmer"] = pd.to_numeric(df["Antal medlemmer"], errors="coerce")
    return df


def clean_meetings_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    if "Starttidspunkt" in df.columns:
        df["Starttidspunkt"] = parse_date_series(df["Starttidspunkt"])
    if "Sluttidspunkt" in df.columns:
        df["Sluttidspunkt"] = parse_date_series(df["Sluttidspunkt"])
    if "Antal deltagere" in df.columns:
        df["Antal deltagere"] = (
            pd.to_numeric(df["Antal deltagere"], errors="coerce")
            .fillna(0)
            .astype(int)
        )
    return df


def clean_seats_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    if "Antal møder" in df.columns:
        df["Antal møder"] = (
            pd.to_numeric(df["Antal møder"], errors="coerce")
            .fillna(0)
            .astype(int)
        )
    return df


def filter_by_period(meetings_df, start_date, end_date):
    if meetings_df is None:
        return None
    df = meetings_df.copy()
    mask = (df["Starttidspunkt"] >= start_date) & (df["Starttidspunkt"] <= end_date)
    return df.loc[mask].reset_index(drop=True)


def get_region_options(groups_df, meetings_df, seats_df):
    regions = set()
    for df in [groups_df, meetings_df, seats_df]:
        if df is not None and "Region" in df.columns:
            regions.update(df["Region"].dropna().unique().tolist())
    return sorted(list(regions))


# -------------------------------------------------
# BEREGNINGER
# -------------------------------------------------

def compute_basic_stats(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return 0, 0
    total_meetings = len(meetings_df)
    total_participant_days = meetings_df["Antal deltagere"].sum()
    return total_meetings, total_participant_days


def compute_meetings_by_type(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    return meetings_df.groupby("Mødetype").size().reset_index(name="Antal møder")


def compute_meetings_by_participant_bins(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    bins = [0, 4, 6, 8, 10, 12, 14, 99]
    labels = ["2-4", "5-6", "7-8", "9-10", "11-12", "13-14", "15-99"]
    df = meetings_df.copy()
    df["Deltagerkategori"] = pd.cut(df["Antal deltagere"], bins=bins, labels=labels, right=True)
    return df.groupby("Deltagerkategori").size().reset_index(name="Antal møder")


def compute_meetings_per_group(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    grp = meetings_df.groupby("Gruppenavn").size().reset_index(name="Antal møder")
    bins = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 1000]
    labels = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10+"]
    grp["Mødekategori"] = pd.cut(grp["Antal møder"], bins=bins, labels=labels, right=True)
    return grp.groupby("Mødekategori").size().reset_index(name="Antal grupper")


def compute_group_size_distribution(groups_df):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    df = groups_df.copy()
    bins = [0, 4, 6, 8, 10, 12, 14, 99]
    labels = ["1-4", "5-6", "7-8", "9-10", "11-12", "13-14", "15-99"]
    df["Størrelseskategori"] = pd.cut(df["Antal medlemmer"], bins=bins, labels=labels, right=True)
    return df.groupby("Størrelseskategori").size().reset_index(name="Antal grupper")


def compute_group_size_by_type(groups_df):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    df = groups_df.copy()
    bins = [0, 4, 6, 8, 10, 12, 14, 99]
    labels = ["1-4", "5-6", "7-8", "9-10", "11-12", "13-14", "15-99"]
    df["Størrelseskategori"] = pd.cut(df["Antal medlemmer"], bins=bins, labels=labels, right=True)
    if "Gruppetyper" not in df.columns:
        df["Gruppetyper"] = "Ukendt"
    return df.groupby(["Størrelseskategori", "Gruppetyper"]).size().reset_index(name="Antal grupper")


def compute_meetings_by_weekday(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()

    df = meetings_df.copy()
    df["Ugedag_eng"] = df["Starttidspunkt"].dt.day_name()

    mapping = {
        "Monday": "Mandag",
        "Tuesday": "Tirsdag",
        "Wednesday": "Onsdag",
        "Thursday": "Torsdag",
        "Friday": "Fredag",
        "Saturday": "Lørdag",
        "Sunday": "Søndag",
    }

    df["Ugedag"] = df["Ugedag_eng"].map(mapping)

    order = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]

    return df.groupby("Ugedag").size().reindex(order).reset_index(name="Antal møder")


def compute_meeting_status(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    df = meetings_df.copy()
    if "Status" not in df.columns:
        df["Status"] = "Ukendt"
    return df.groupby("Status").size().reset_index(name="Antal møder")


def compute_groups_without_meetings(groups_df, meetings_df, period_start, period_end):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    df_g = groups_df.copy()
    if meetings_df is None or meetings_df.empty:
        return df_g[["Gruppenavn"]].copy()

    df_m = meetings_df.copy()
    mask = (df_m["Starttidspunkt"] >= period_start) & (df_m["Starttidspunkt"] <= period_end)
    df_m = df_m.loc[mask]
    groups_with_meetings = set(df_m["Gruppenavn"].dropna().unique().tolist())
    df_g["Har møde"] = df_g["Gruppenavn"].isin(groups_with_meetings)
    return df_g.loc[~df_g["Har møde"], ["Gruppenavn"]]


def compute_closed_groups_this_year(groups_df, year):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    df = groups_df.copy()
    if "Dato for arkivering" not in df.columns:
        return pd.DataFrame()
    df["År for arkivering"] = df["Dato for arkivering"].dt.year
    return df.loc[df["År for arkivering"] == year, ["Gruppenavn"]]


def compute_member_types(seats_df):
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    df = seats_df.copy()
    if "Stillingsbetegnelse" not in df.columns:
        return pd.DataFrame()
    return df.groupby("Stillingsbetegnelse").size().reset_index(name="Antal")


def compute_groups_per_person(seats_df):
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    df = seats_df.copy()
    if "Medlemskaber" not in df.columns:
        return pd.DataFrame()
    df["Antal grupper"] = df["Medlemskaber"].fillna("").apply(
        lambda x: len([g for g in str(x).split(",") if g.strip() != ""])
    )
    return df.groupby("Antal grupper").size().reset_index(name="Antal personer")


def compute_groups_per_region_per_100k(groups_df):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    df = groups_df.copy()
    if "Region" not in df.columns:
        return pd.DataFrame()
    grp = df.groupby("Region").size().reset_index(name="Antal grupper")
    grp["Befolkning"] = grp["Region"].map(REGION_POPULATION).fillna(np.nan)
    grp["Grupper pr. 100.000 borgere"] = grp["Antal grupper"] / grp["Befolkning"] * 100000
    return grp


# -------------------------------------------------
# PLOTS
# -------------------------------------------------

def plot_bar(df, x_col, y_col, title, rotation=0):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(df[x_col].astype(str), df[y_col])
    ax.set_title(title)
    ax.set_ylabel(y_col)
    ax.set_xlabel(x_col)
    plt.xticks(rotation=rotation)
    plt.tight_layout()
    return fig


def plot_stacked_group_size_by_type(df):
    if df.empty:
        return None
    pivot = df.pivot(index="Størrelseskategori", columns="Gruppetyper", values="Antal grupper").fillna(0)
    fig, ax = plt.subplots(figsize=(7, 4))
    pivot.plot(kind="bar", stacked=True, ax=ax)
    ax.set_title("Gruppestørrelse fordelt på gruppetype")
    ax.set_xlabel("Gruppestørrelse")
    ax.set_ylabel("Antal grupper")
    plt.xticks(rotation=0)
    plt.tight_layout()
    return fig


# -------------------------------------------------
# PDF MED FRASORTEREDE DATA
# -------------------------------------------------

def build_removed_pdf(removed_groups, removed_meetings, removed_seats):
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:

        def write_list(title, items):
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title(title, loc="left")
            y = 0.95
            for item in items:
                ax.text(0.05, y, f"- {item}", fontsize=9, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

        if removed_groups is not None and not removed_groups.empty:
            write_list("Frasorterede grupper", removed_groups["Gruppenavn"].astype(str).tolist())

        if removed_meetings is not None and not removed_meetings.empty:
            write_list("Frasorterede møder (gruppenavne)", removed_meetings["Gruppenavn"].astype(str).tolist())

        if removed_seats is not None and not removed_seats.empty:
            if "Fornavn" in removed_seats.columns and "Efternavn" in removed_seats.columns:
                navne = (removed_seats["Fornavn"].fillna("") + " " +
                         removed_seats["Efternavn"].fillna("")).str.strip()
            else:
                navne = removed_seats.get("Medlemskaber", pd.Series()).astype(str)
            write_list("Frasorterede brugere", navne.tolist())

    buf.seek(0)
    return buf


# -------------------------------------------------
# NY DATARNSNING
# -------------------------------------------------

def is_test_group(name: str) -> bool:
    name = str(name).lower()
    return any(p in name for p in TEST_PATTERNS)


def clean_groups(groups_df):
    df = groups_df.copy()
    removed = df[df["Gruppenavn"].str.lower().apply(is_test_group)]
    df = df[~df["Gruppenavn"].str.lower().apply(is_test_group)]
    df = df[~df["Gruppenavn"].str.lower().str.contains("gruppeledere")]
    return df, removed


def clean_meetings(meetings_df):
    df = meetings_df.copy()
    removed = df[df["Gruppenavn"].str.lower().apply(is_test_group)]
    df = df[~df["Gruppenavn"].str.lower().apply(is_test_group)]
    df = df[~df["Gruppenavn"].str.lower().str.contains("gruppeledere")]
    return df, removed


def clean_seats(seats_df):
    df = seats_df.copy()

    df["Medlemsliste"] = df["Medlemskaber"].fillna("").apply(
        lambda x: [g.strip() for g in str(x).split(",") if g.strip() != ""]
    )

    df["gruppeledere_count"] = df["Medlemsliste"].apply(
        lambda lst: sum("gruppeledere" in g.lower() for g in lst)
    )

    removed_gl = df[df["gruppeledere_count"] > 1].copy()

    df = df[df["gruppeledere_count"] <= 1]

    def is_test_only(medlemsliste):
        if len(medlemsliste) == 0:
            return False
        tests = [g for g in medlemsliste if any(p in g.lower() for p in TEST_PATTERNS)]
        return len(tests) == len(medlemsliste) and len(medlemsliste) > 0

    removed_test = df[df["Medlemsliste"].apply(is_test_only)].copy()

    df = df[~df["Medlemsliste"].apply(is_test_only)]

    for frame in (df, removed_gl, removed_test):
        for col in ["Medlemsliste", "gruppeledere_count"]:
            if col in frame.columns:
                frame.drop(columns=[col], inplace=True)

    removed = pd.concat([removed_gl, removed_test], ignore_index=True)

    return df, removed


# -------------------------------------------------
# PDF-GENERERING (DIN ORIGINALE)
# -------------------------------------------------

def build_pdf_report(
    basic_stats,
    meetings_by_type,
    meetings_by_participant_bins,
    meetings_per_group,
    group_size_dist,
    group_size_by_type,
    meetings_by_weekday,
    meeting_status,
    groups_without_meetings,
    closed_groups,
    member_types,
    groups_per_person,
    groups_per_region_norm,
):
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        fig, ax = plt.subplots(figsize=(8.27, 11.69))
        ax.axis("off")
        total_meetings, total_participant_days = basic_stats
        text_lines = [
            "DGE-rapport (automatisk genereret)",
            "",
            f"Antal godkendte møder i perioden: {total_meetings}",
            f"Antal kursusdeltagerdage i perioden: {total_participant_days}",
        ]
        ax.text(0.05, 0.95, "\n".join(text_lines), va="top", fontsize=12)
        pdf.savefig(fig)
        plt.close(fig)

        if not meetings_by_type.empty:
            fig = plot_bar(meetings_by_type, "Mødetype", "Antal møder", "Møder fordelt på type", rotation=45)
            pdf.savefig(fig)
            plt.close(fig)

        if not meetings_by_participant_bins.empty:
            fig = plot_bar(meetings_by_participant_bins, "Deltagerkategori", "Antal møder", "Mødedeltagelse", rotation=0)
            pdf.savefig(fig)
