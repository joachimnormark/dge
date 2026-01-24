import io
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

import streamlit as st

# ---------------------------------------------------------
# KONFIGURATION
# ---------------------------------------------------------

st.set_page_config(page_title="DGE-rapport", layout="wide")

# Befolkningstal pr. region (kan tilpasses)
REGION_POPULATION = {
    "Region Nordjylland": 590000,
    "Region Midtjylland": 1330000,
    "Region Syddanmark": 1220000,
    "Region Hovedstaden": 1870000,
    "Region Sjælland": 840000,
}

# Navn på den særlige gruppe
GROUP_LEADERS_NAME = "Gruppeledere"


# ---------------------------------------------------------
# HJÆLPEFUNKTIONER TIL DATOER OG INDLÆSNING
# ---------------------------------------------------------

def parse_date_series(s):
    """Forsøger at parse datoer robust fra forskellige formater."""
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def load_excel(uploaded_file):
    """Indlæser en Excel-fil til DataFrame, hvis der er uploadet en fil."""
    if uploaded_file is None:
        return None
    return pd.read_excel(uploaded_file)


# ---------------------------------------------------------
# RENSNING AF RÅ DATAFRAMES (KUN TEKNISK, IKKE LOGISK FILTRERING)
# ---------------------------------------------------------

def clean_groups_df(df):
    """Standardiserer kolonnenavne og typer for grupper."""
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Dato for arkivering kan være NaN eller '-'
    if "Dato for arkivering" in df.columns:
        df["Dato for arkivering"] = df["Dato for arkivering"].replace("-", np.nan)
        df["Dato for arkivering"] = parse_date_series(df["Dato for arkivering"])

    # Antal medlemmer til numerisk
    if "Antal medlemmer" in df.columns:
        df["Antal medlemmer"] = pd.to_numeric(df["Antal medlemmer"], errors="coerce")

    return df


def clean_meetings_df(df):
    """Standardiserer kolonnenavne og typer for møder."""
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Datoer
    if "Starttidspunkt" in df.columns:
        df["Starttidspunkt"] = parse_date_series(df["Starttidspunkt"])
    if "Sluttidspunkt" in df.columns:
        df["Sluttidspunkt"] = parse_date_series(df["Sluttidspunkt"])

    # Antal deltagere
    if "Antal deltagere" in df.columns:
        df["Antal deltagere"] = pd.to_numeric(df["Antal deltagere"], errors="coerce").fillna(0).astype(int)

    return df


def clean_seats_df(df):
    """Standardiserer kolonnenavne og typer for medlems-/sædedata."""
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Antal møder
    if "Antal møder" in df.columns:
        df["Antal møder"] = pd.to_numeric(df["Antal møder"], errors="coerce").fillna(0).astype(int)

    return df


# ---------------------------------------------------------
# LOGISK FILTRERING IFT. DINE REGLER
# ---------------------------------------------------------

def mark_excluded_groups(groups_df):
    """
    Marker grupper, der skal frasorteres som grupper:
    1) Gruppenavn indeholder 'test' eller 'demo' (case-insensitive)
    2) Gruppenavn er præcist 'Gruppeledere'
    Returnerer:
        groups_df med kolonne 'Er_frasorteret_gruppe' (bool)
    """
    df = groups_df.copy()
    df["Gruppenavn"] = df["Gruppenavn"].astype(str)

    name_lower = df["Gruppenavn"].str.lower()

    is_test_or_demo = name_lower.str.contains("test") | name_lower.str.contains("demo")
    is_group_leaders = df["Gruppenavn"] == GROUP_LEADERS_NAME

    df["Er_frasorteret_gruppe"] = is_test_or_demo | is_group_leaders
    return df


def filter_meetings_by_period(meetings_df, start_dt, end_dt):
    """Filtrerer møder på valgt periode (inklusive begge ender)."""
    if meetings_df is None:
        return None
    df = meetings_df.copy()
    mask = (df["Starttidspunkt"] >= start_dt) & (df["Starttidspunkt"] <= end_dt)
    return df.loc[mask].reset_index(drop=True)


def filter_meetings_by_excluded_groups(meetings_df, excluded_group_names):
    """
    Fjerner møder, der ligger i grupper, som er frasorteret som grupper
    (test/demo eller præcis 'Gruppeledere').
    """
    if meetings_df is None or meetings_df.empty:
        return meetings_df
    df = meetings_df.copy()
    df = df[~df["Gruppenavn"].isin(excluded_group_names)].reset_index(drop=True)
    return df


def filter_meetings_by_group_archiving(meetings_df, groups_df, start_dt):
    """
    Håndterer inaktive grupper ift. arkiveringsdato:

    - Hvis Dato for arkivering er før periodens start -> alle møder i perioden fjernes.
    - Hvis Dato for arkivering er efter periodens start -> møder efter arkiveringsdato fjernes.
    - Hvis Dato for arkivering mangler -> gruppen antages aktiv (A-reglen).
    """
    if meetings_df is None or meetings_df.empty:
        return meetings_df

    df_m = meetings_df.copy()
    df_g = groups_df.copy()

    # Vi behøver kun Gruppenavn + Dato for arkivering
    merge_cols = ["Gruppenavn"]
    if "Dato for arkivering" in df_g.columns:
        df_g["Dato for arkivering"] = parse_date_series(df_g["Dato for arkivering"])
        merge_cols.append("Dato for arkivering")
    else:
        df_g["Dato for arkivering"] = pd.NaT
        merge_cols.append("Dato for arkivering")

    df = df_m.merge(df_g[merge_cols].drop_duplicates(), on="Gruppenavn", how="left")

    # Regel 1: grupper arkiveret før periodens start -> fjern alle møder
    mask_before_period = df["Dato for arkivering"].notna() & (df["Dato for arkivering"] < start_dt)

    # Regel 2: grupper arkiveret efter periodens start -> fjern møder efter arkiveringsdato
    mask_after_archiving = df["Dato for arkivering"].notna() & (df["Starttidspunkt"] > df["Dato for arkivering"])

    df_filtered = df[~mask_before_period & ~mask_after_archiving].copy()

    # Behold kun de oprindelige mødekolonner
    return df_filtered[df_m.columns].reset_index(drop=True)


def filter_approved_meetings(meetings_df):
    """
    Beholder kun møder med Status == 'Godkendt', hvis kolonnen findes.
    Dette bruges til de fleste statistikker, da du selv tæller godkendte møder.
    """
    if meetings_df is None or meetings_df.empty:
        return meetings_df
    if "Status" not in meetings_df.columns:
        return meetings_df
    return meetings_df[meetings_df["Status"] == "Godkendt"].reset_index(drop=True)


def clean_seats_logic(seats_df, excluded_group_names):
    """
    Implementerer dine regler for medlemsdata:

    4: Når der tælles hvor mange grupper et medlem er medlem i,
       skal medlemskab i 'Gruppeledere' ikke tælle med,
       og test/demo-grupper skal heller ikke tælle med.

    5: Medlemmer der er medlem i flere grupper der hedder 'Gruppeledere'
       skal slet ikke tælles med i medlemsstatistik.

    Returnerer:
        seats_for_stats: renset DataFrame til medlemsstatistik
    """
    if seats_df is None or seats_df.empty:
        return seats_df

    df = seats_df.copy()

    if "Medlemskaber" not in df.columns:
        # Hvis kolonnen ikke findes, kan vi ikke lave logikken – returner som er
        return df

    # Sørg for at arbejde på strenge
    df["Medlemskaber"] = df["Medlemskaber"].fillna("").astype(str)

    # Split medlemskaber til lister
    df["Medlemskaber_liste"] = df["Medlemskaber"].apply(
        lambda x: [g.strip() for g in str(x).split(",") if g.strip() != ""]
    )

    # Tæl hvor mange gange 'Gruppeledere' optræder
    df["Antal_gruppeledere_medlemskaber"] = df["Medlemskaber_liste"].apply(
        lambda lst: sum(1 for g in lst if g == GROUP_LEADERS_NAME)
    )

    # Regel 5: fjern medlemmer med flere 'Gruppeledere'-medlemskaber
    df = df[df["Antal_gruppeledere_medlemskaber"] <= 1].copy()

    # Til beregning af "Antal grupper man er medlem af":
    # - Ekskluder 'Gruppeledere'
    # - Ekskluder grupper med 'test' eller 'demo' i navnet
    excluded_lower = {g.lower() for g in excluded_group_names}

    def count_relevant_groups(lst):
        count = 0
        for g in lst:
            gl = g.lower()
            if g == GROUP_LEADERS_NAME:
                continue
            if "test" in gl or "demo" in gl:
                continue
            # Hvis gruppen er i den globale eksklusionsliste (fx 'Gruppeledere', test/demo)
            if gl in excluded_lower:
                continue
            count += 1
        return count

    df["Antal_relevante_grupper"] = df["Medlemskaber_liste"].apply(count_relevant_groups)

    return df


def get_region_options(groups_df, meetings_df, seats_df):
    """Finder alle regioner på tværs af de tre datasæt."""
    regions = set()
    for df in [groups_df, meetings_df, seats_df]:
        if df is not None and "Region" in df.columns:
            regions.update(df["Region"].dropna().unique().tolist())
    return sorted(list(regions))


# ---------------------------------------------------------
# BEREGNINGER / AGGREGATER
# ---------------------------------------------------------

def compute_basic_stats(meetings_df):
    """Returnerer (antal møder, antal kursusdeltagerdage) for de filtrerede møder."""
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

    result = df.groupby("Ugedag").size().reindex(order).reset_index(name="Antal møder")
    return result


def compute_meeting_status(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    df = meetings_df.copy()
    if "Status" not in df.columns:
        df["Status"] = "Ukendt"
    return df.groupby("Status").size().reset_index(name="Antal møder")


def compute_groups_without_meetings(groups_df, meetings_df, period_start, period_end):
    """
    Finder grupper uden godkendte møder i perioden.
    Bruger de allerede filtrerede (godkendte) møder.
    """
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


def compute_member_types(seats_df_for_stats):
    if seats_df_for_stats is None or seats_df_for_stats.empty:
        return pd.DataFrame()
    df = seats_df_for_stats.copy()
    if "Stillingsbetegnelse" not in df.columns:
        return pd.DataFrame()
    return df.groupby("Stillingsbetegnelse").size().reset_index(name="Antal")


def compute_groups_per_person(seats_df_for_stats):
    """
    Bruger kolonnen 'Antal_relevante_grupper', som er beregnet i clean_seats_logic.
    """
    if seats_df_for_stats is None or seats_df_for_stats.empty:
        return pd.DataFrame()
    df = seats_df_for_stats.copy()
    if "Antal_relevante_grupper" not in df.columns:
        return pd.DataFrame()
    return df.groupby("Antal_relevante_grupper").size().reset_index(name="Antal personer")


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


# ---------------------------------------------------------
# PLOTTES
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# PDF-GENERERING
# ---------------------------------------------------------

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
        # Side 1: Basis tal
        fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4
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

        # Side 2: Møder fordelt på type
        if not meetings_by_type.empty:
            fig = plot_bar(meetings_by_type, "Mødetype", "Antal møder", "Møder fordelt på type", rotation=45)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 3: Mødedeltagelse (bins)
        if not meetings_by_participant_bins.empty:
            fig = plot_bar(meetings_by_participant_bins, "Deltagerkategori", "Antal møder", "Mødedeltagelse", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 4: Gruppernes mødeaktivitet
        if not meetings_per_group.empty:
            fig = plot_bar(meetings_per_group, "Mødekategori", "Antal grupper", "Gruppernes mødeaktivitet", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 5: Gruppestørrelse
        if not group_size_dist.empty:
            fig = plot_bar(group_size_dist, "Størrelseskategori", "Antal grupper", "Gruppestørrelse", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 6: Gruppestørrelse fordelt på gruppetype
        if not group_size_by_type.empty:
            fig = plot_stacked_group_size_by_type(group_size_by_type)
            if fig is not None:
                pdf.savefig(fig)
                plt.close(fig)

        # Side 7: Mødedage
        if not meetings_by_weekday.empty:
            fig = plot_bar(meetings_by_weekday, "Ugedag", "Antal møder", "Mødedage", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 8: Mødestatus
        if not meeting_status.empty:
            fig = plot_bar(meeting_status, "Status", "Antal møder", "Mødestatus", rotation=45)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 9: Grupper uden møder
        if groups_without_meetings is not None and not groups_without_meetings.empty:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title("Grupper uden godkendte møder i perioden", loc="left")
            y = 0.95
            for name in groups_without_meetings["Gruppenavn"].tolist():
                ax.text(0.05, y, f"- {name}", fontsize=9, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

        # Side 10: Grupper lukket i år
        if closed_groups is not None and not closed_groups.empty:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title("Grupper lukket i året", loc="left")
            y = 0.95
            for name in closed_groups["Gruppenavn"].tolist():
                ax.text(0.05, y, f"- {name}", fontsize=9, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

        # Side 11: Medlemstyper
        if member_types is not None and not member_types.empty:
            fig = plot_bar(member_types, "Stillingsbetegnelse", "Antal", "Medlemstyper", rotation=90)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 12: Antal grupper man er medlem af
        if groups_per_person is not None and not groups_per_person.empty:
            fig = plot_bar(groups_per_person, "Antal_relevante_grupper", "Antal personer", "Antal grupper man er medlem af", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 13: Grupper pr. region pr. 100.000 borgere
        if groups_per_region_norm is not None and not groups_per_region_norm.empty:
            fig = plot_bar(groups_per_region_norm, "Region", "Grupper pr. 100.000 borgere", "Grupper pr. 100.000 borgere", rotation=45)
            pdf.savefig(fig)
            plt.close(fig)

    buf.seek(0)
    return buf


# ---------------------------------------------------------
# STREAMLIT-APP
# ---------------------------------------------------------

def main():
    st.title("DGE-rapportgenerator (Regioner, grupper og møder)")

    st.markdown(
        """
        Denne app bruger **kun** følgende rensningsregler:

        1. Grupper, hvor navnet indeholder *test* eller *demo*, frasorteres som grupper.
        2. Gruppen der hedder præcist **'Gruppeledere'** frasorteres som gruppe.
        3. Møder i de frasorterede grupper tæller ikke med nogen steder.
        4. Når antal grupper pr. medlem beregnes, tæller:
           - 'Gruppeledere' ikke med
           - test/demo-grupper ikke med
        5. Medlemmer der er medlem af 'Gruppeledere' **flere gange** tælles slet ikke i medlemsstatistikken.
        6. Inaktive grupper tæller med, hvis de har møder før deres arkiveringsdato.
        """
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        groups_file = st.file_uploader("Gruppe-data (fx workspace_groups_all-RN.xlsx)", type=["xlsx"])
    with col2:
        meetings_file = st.file_uploader("Møde-data (fx workspace_meetings_all-RN.xlsx)", type=["xlsx"])
    with col3:
        seats_file = st.file_uploader("Medlems-/sæde-data (fx workspace_seats_all-RN.xlsx)", type=["xlsx"])

    if not (groups_file and meetings_file and seats_file):
        st.info("Upload alle tre filer for at fortsætte.")
        return

    # Indlæs og rens data (teknisk)
    groups_df_raw = load_excel(groups_file)
    meetings_df_raw = load_excel(meetings_file)
    seats_df_raw = load_excel(seats_file)

    groups_df = clean_groups_df(groups_df_raw)
    meetings_df = clean_meetings_df(meetings_df_raw)
    seats_df = clean_seats_df(seats_df_raw)

    # Vælg periode
    st.subheader("Periodevalg")
    min_date = meetings_df["Starttidspunkt"].min()
    max_date = meetings_df["Starttidspunkt"].max()
    col_start, col_end = st.columns(2)
    with col_start:
        start_date = st.date_input(
            "Startdato",
            value=min_date.date() if pd.notnull(min_date) else datetime.today().date()
        )
    with col_end:
        end_date = st.date_input(
            "Slutdato",
            value=max_date.date() if pd.notnull(max_date) else datetime.today().date()
        )

    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())

    # Vælg region(er)
    st.subheader("Regioner")
    region_options = get_region_options(groups_df, meetings_df, seats_df)
    selected_regions = st.multiselect("Vælg region(er)", options=region_options, default=region_options)

    # Filtrér på regioner
    if selected_regions:
        groups_df = groups_df[groups_df["Region"].isin(selected_regions)].reset_index(drop=True)
        meetings_df = meetings_df[meetings_df["Region"].isin(selected_regions)].reset_index(drop=True)
        seats_df = seats_df[seats_df["Region"].isin(selected_regions)].reset_index(drop=True)

    # Marker frasorterede grupper (test/demo + Gruppeledere)
    groups_df = mark_excluded_groups(groups_df)
    excluded_groups = groups_df.loc[groups_df["Er_frasorteret_gruppe"], "Gruppenavn"].unique().tolist()

    # Lav en version af grupper til gruppestatistik (uden frasorterede grupper)
    groups_for_stats = groups_df[~groups_df["Er_frasorteret_gruppe"]].copy()

    # Filtrér møder på periode
    meetings_period_df = filter_meetings_by_period(meetings_df, start_dt, end_dt)

    # Fjern møder i frasorterede grupper
    meetings_period_df = filter_meetings_by_excluded_groups(meetings_period_df, excluded_groups)

    # Håndter inaktive grupper ift. arkiveringsdato
    meetings_period_df = filter_meetings_by_group_archiving(meetings_period_df, groups_df, start_dt)

    # Behold kun godkendte møder til de fleste statistikker
    meetings_approved_df = filter_approved_meetings(meetings_period_df)

    if meetings_approved_df is None or meetings_approved_df.empty:
        st.warning("Ingen godkendte møder i den valgte periode/region(er) efter filtrering. Prøv at ændre periode eller region.")
        return

    # Rens medlemsdata ift. dine regler
    seats_for_stats = clean_seats_logic(seats_df, excluded_groups)

    # -----------------------------------------------------
    # BEREGN ALLE AGGREGATER
    # -----------------------------------------------------

    basic_stats = compute_basic_stats(meetings_approved_df)
    meetings_by_type = compute_meetings_by_type(meetings_approved_df)
    meetings_by_participant_bins = compute_meetings_by_participant_bins(meetings_approved_df)
    meetings_per_group = compute_meetings_per_group(meetings_approved_df)
    group_size_dist = compute_group_size_distribution(groups_for_stats)
    group_size_by_type = compute_group_size_by_type(groups_for_stats)
    meetings_by_weekday = compute_meetings_by_weekday(meetings_approved_df)
    meeting_status = compute_meeting_status(meetings_period_df)  # her ser vi alle statusser
    groups_without_meetings = compute_groups_without_meetings(groups_for_stats, meetings_approved_df, start_dt, end_dt)
    closed_groups = compute_closed_groups_this_year(groups_for_stats, year=start_date.year)
    member_types = compute_member_types(seats_for_stats)
    groups_per_person = compute_groups_per_person(seats_for_stats)
    groups_per_region_norm = compute_groups_per_region_per_100k(groups_for_stats)

    # -----------------------------------------------------
    # VISUEL VISNING I APPEN
    # -----------------------------------------------------

    st.subheader("Overblik")
    total_meetings, total_participant_days = basic_stats
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Antal godkendte møder i perioden", total_meetings)
    with col_b:
        st.metric("Antal kursusdeltagerdage i perioden", total_participant_days)

    st.subheader("Møder fordelt på type (kun godkendte)")
    if not meetings_by_type.empty:
        st.bar_chart(meetings_by_type.set_index("Mødetype")["Antal møder"])

    st.subheader("Mødedeltagelse (antal deltagere pr. møde, kun godkendte)")
    if not meetings_by_participant_bins.empty:
        st.bar_chart(meetings_by_participant_bins.set_index("Deltagerkategori")["Antal møder"])

    st.subheader("Gruppernes mødeaktivitet (antal møder pr. gruppe, kun godkendte)")
    if not meetings_per_group.empty:
        st.bar_chart(meetings_per_group.set_index("Mødekategori")["Antal grupper"])

    st.subheader("Gruppestørrelse (frasorterede grupper er fjernet)")
    if not group_size_dist.empty:
        st.bar_chart(group_size_dist.set_index("Størrelseskategori")["Antal grupper"])

    st.subheader("Gruppestørrelse fordelt på gruppetype")
    if not group_size_by_type.empty:
        pivot = group_size_by_type.pivot(index="Størrelseskategori", columns="Gruppetyper", values="Antal grupper").fillna(0)
        st.bar_chart(pivot)

    st.subheader("Mødedage (kun godkendte møder)")
    if not meetings_by_weekday.empty:
        st.bar_chart(meetings_by_weekday.set_index("Ugedag")["Antal møder"])

    st.subheader("Mødestatus (alle møder i perioden efter filtrering af grupper)")
    if not meeting_status.empty:
        st.bar_chart(meeting_status.set_index("Status")["Antal møder"])

    st.subheader("Grupper uden godkendte møder i perioden (frasorterede grupper er fjernet)")
    if groups_without_meetings is not None and not groups_without_meetings.empty:
        st.dataframe(groups_without_meetings)

    st.subheader(f"Grupper lukket i {start_date.year} (frasorterede grupper er fjernet)")
    if closed_groups is not None and not closed_groups.empty:
        st.dataframe(closed_groups)

    st.subheader("Medlemstyper (stillingsbetegnelser, efter dine medlemsregler)")
    if member_types is not None and not member_types.empty:
        st.dataframe(member_types)

    st.subheader("Antal grupper man er medlem af (efter dine medlemsregler)")
    if groups_per_person is not None and not groups_per_person.empty:
        st.bar_chart(groups_per_person.set_index("Antal_relevante_grupper")["Antal personer"])

    st.subheader("Grupper pr. region pr. 100.000 borgere (frasorterede grupper er fjernet)")
    if groups_per_region_norm is not None and not groups_per_region_norm.empty:
        st.dataframe(groups_per_region_norm)

    # -----------------------------------------------------
    # PDF DOWNLOAD
    # -----------------------------------------------------

    st.subheader("Download rapport som PDF")
    pdf_buffer = build_pdf_report(
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
    )

    st.download_button(
        label="Download PDF-rapport",
        data=pdf_buffer,
        file_name="dge_rapport.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
