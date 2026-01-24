import io
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

import streamlit as st

# ---------- KONFIG ----------

st.set_page_config(page_title="DGE-rapport", layout="wide")

# Befolkningstal pr. region (ret disse til de rigtige tal)
REGION_POPULATION = {
    "Region Nordjylland": 590000,
    "Region Midtjylland": 1330000,
    "Region Syddanmark": 1220000,
    "Region Hovedstaden": 1870000,
    "Region Sjælland": 840000,
}

# ---------- HJÆLPEFUNKTIONER ----------

def parse_date_series(s):
    """
    Robust parsing af danske datoer som fx:
    '14. december 2025, kl. 20:57'
    '5. januar 2026 kl. 07:38'
    '24. maj 2024'
    '2025-12-14'
    '14/12/2025'
    """
    if s is None:
        return pd.NaT

    # Hvis det allerede er en datetime
    if isinstance(s, datetime):
        return s

    # Konverter til tekst
    s = str(s).strip()

    # Fjern "kl." og komma
    s = s.replace("kl.", "").replace("Kl.", "").replace("KL.", "")
    s = s.replace(",", "")

    # Dansk måned → tal
    months = {
        "januar": "01",
        "februar": "02",
        "marts": "03",
        "april": "04",
        "maj": "05",
        "juni": "06",
        "juli": "07",
        "august": "08",
        "september": "09",
        "oktober": "10",
        "november": "11",
        "december": "12",
    }

    for dk, num in months.items():
        if dk in s.lower():
            s = s.lower().replace(dk, num)

    # Nu ligner datoen noget pandas kan læse
    # Eksempel: "14. 12 2025 20:57"
    s = s.replace(".", " ").replace("/", " ")

    # Forsøg parsing
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except:
        return pd.NaT



def load_excel(uploaded_file):
    if uploaded_file is None:
        return None
    return pd.read_excel(uploaded_file)


def clean_groups_df(df):
    if df is None:
        return None

    df = df.copy()

    # DEBUG: vis kolonnenavne
    # st.write("DEBUG – kolonner i groups_df:", df.columns.tolist())

    # DEBUG: vis rå værdier i arkiveringskolonnen (hvis den findes)
    #if "Dato for arkivering" in df.columns:
     #   st.write("DEBUG – rå arkiveringsværdier:", df["Dato for arkivering"].head(20).tolist())
    #else:
     #   st.write("DEBUG – ingen kolonne med navnet 'Dato for arkivering' fundet")

    # Standardiser kolonnenavne
    df.columns = [c.strip() for c in df.columns]

    # ⭐ Dato for arkivering – robust parsing
    if "Dato for arkivering" in df.columns:

        # Erstat '-' og tomme felter med NaN
        df["Dato for arkivering"] = df["Dato for arkivering"].replace(["-", ""], np.nan)

        # Trim tekst
        df["Dato for arkivering"] = df["Dato for arkivering"].astype(str).str.strip()

        # ⭐ Fjern "kl. HH:MM" hvis det findes
        df["Dato for arkivering"] = df["Dato for arkivering"].str.replace(r",?\s*kl\.\s*\d{1,2}:\d{2}", "", regex=True)

        # ⭐ Dansk månedstabel
        month_map = {
            "januar": "01", "februar": "02", "marts": "03", "april": "04",
            "maj": "05", "juni": "06", "juli": "07", "august": "08",
            "september": "09", "oktober": "10", "november": "11", "december": "12"
        }

        # ⭐ Konverter tekst-måneder til tal
        for dk, num in month_map.items():
            df["Dato for arkivering"] = df["Dato for arkivering"].str.replace(dk, num, regex=False)

        # ⭐ Fjern punktum efter dag
        df["Dato for arkivering"] = df["Dato for arkivering"].str.replace(".", "", regex=False)

        # ⭐ Nu burde datoen være i formatet "24 05 2024"
        df["Dato for arkivering"] = pd.to_datetime(
            df["Dato for arkivering"],
            format="%d %m %Y",
            errors="coerce"
        )

    else:
        df["Dato for arkivering"] = pd.NaT

    # Antal medlemmer til numerisk
    if "Antal medlemmer" in df.columns:
        df["Antal medlemmer"] = pd.to_numeric(df["Antal medlemmer"], errors="coerce")

    return df






def clean_meetings_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    # Datoer
    if "Starttidspunkt" in df.columns:
        df["Starttidspunkt"] = df["Starttidspunkt"].apply(parse_date_series)
    if "Sluttidspunkt" in df.columns:
        df["Sluttidspunkt"] = df["Sluttidspunkt"].apply(parse_date_series)
    # Antal deltagere
    if "Antal deltagere" in df.columns:
        df["Antal deltagere"] = pd.to_numeric(df["Antal deltagere"], errors="coerce").fillna(0).astype(int)
    return df


def clean_seats_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    # Antal møder
    if "Antal møder" in df.columns:
        df["Antal møder"] = pd.to_numeric(df["Antal møder"], errors="coerce").fillna(0).astype(int)
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


# ---------- BEREGNINGER / AGGREGATER ----------

def compute_basic_stats(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return 0, 0
    total_meetings = len(meetings_df)
    # Kursusdeltagerdage ~ sum(antal deltagere) (her uden varighed)
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

    # Få engelske ugedage (virker altid)
    df["Ugedag_eng"] = df["Starttidspunkt"].dt.day_name()

    # Oversæt til dansk
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

    # Rækkefølge
    order = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]

    # Gør Ugedag til ordnet kategori
    cat = pd.CategoricalDtype(order, ordered=True)
    df["Ugedag"] = df["Ugedag"].astype(cat)

    # Gruppér og sortér
    result = df.groupby("Ugedag").size().reset_index(name="Antal møder")
    result = result.sort_values("Ugedag")

    return result




def compute_meeting_status(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()

    df = meetings_df.copy()

    # Normalisér status
    df["Status"] = (
        df["Status"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # Blank status = tom streng
    df.loc[df["Status"].isin(["nan", "none", ""]), "Status"] = ""

    # Gruppér
    result = df.groupby("Status").size().reset_index(name="Antal møder")

    # Korrekt rækkefølge
    order = ["godkendt", "afvist", "afsluttet", ""]

    # Gør status til ordnet kategori
    cat = pd.CategoricalDtype(order, ordered=True)
    result["Status"] = result["Status"].astype(cat)

    # Sortér
    result = result.sort_values("Status")

    # Sikr at alle statusser er med (også dem der ikke findes)
    result = result.set_index("Status").reindex(order, fill_value=0).reset_index()

    # Gør blank status pænere i tabellen
    result.loc[result["Status"] == "", "Status"] = "Ingen status"

    return result



def compute_groups_without_meetings(groups_df, meetings_df, period_start, period_end):
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()

    df_g = groups_df.copy()

    # Kun grupper der var aktive i perioden
    df_g["Status"] = "Aktiv"
    df_g.loc[
        df_g["Dato for arkivering"].notna() &
        (df_g["Dato for arkivering"] < period_start),
        "Status"
    ] = "Inaktiv"

    # Fjern inaktive grupper der blev arkiveret før perioden
    df_g = df_g[df_g["Status"] == "Aktiv"]

    # Fjern Gruppeledere
    df_g = df_g[
        df_g["Gruppenavn"].astype(str).str.strip().str.lower() != "gruppeledere"
    ]

    # Hvis der ingen møder er, returnér alle aktive grupper
    if meetings_df is None or meetings_df.empty:
        return df_g[["Gruppenavn", "Dato for arkivering", "Status"]].copy()

    df_m = meetings_df.copy()

    # Begræns møder til perioden
    mask = (df_m["Starttidspunkt"] >= period_start) & (df_m["Starttidspunkt"] <= period_end)
    df_m = df_m.loc[mask]

    groups_with_meetings = set(df_m["Gruppenavn"].dropna().unique().tolist())

    df_g["Har møde"] = df_g["Gruppenavn"].isin(groups_with_meetings)

    # Returnér grupper uden møder + arkiveringsdato + status
    return df_g.loc[~df_g["Har møde"], ["Gruppenavn", "Dato for arkivering", "Status"]]




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
    # Antal grupper = antal komma-separerede elementer
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

def compute_meetings_by_type_approved(meetings_df):
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()

    df = meetings_df.copy()

    # Normaliser status
    df["Status"] = df["Status"].astype(str).str.strip().str.lower()
    df = df[df["Status"] == "godkendt"]
    if df.empty:
        return pd.DataFrame()

    # Normaliser mødetype (trim + lowercase til mapping)
    df["Mødetype_norm"] = (
        df["Mødetype"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # Mapping af variationer → standardnavn
    type_map = {
        "dge-møde": "DGE-møde",
        "supervision": "Supervision",
        "sge-modul": "SGE-modul",
        "dge-modul": "DGE-modul",
        "ej refusion": "Ej refusion",
        "regional hyldevare": "Regional hyldevare",
    }

    # Map til standardnavne
    df["Mødetype_std"] = df["Mødetype_norm"].map(type_map)

    # Fjern typer vi ikke kender
    df = df[df["Mødetype_std"].notna()]

    # Gruppér
    result = df.groupby("Mødetype_std").size().reset_index(name="Antal møder")

    # Korrekt rækkefølge
    order = [
        "DGE-møde",
        "Supervision",
        "SGE-modul",
        "DGE-modul",
        "Regional hyldevare",
        "Ej refusion",
    ]

    # Ordnet kategori
    cat = pd.CategoricalDtype(order, ordered=True)
    result["Mødetype_std"] = result["Mødetype_std"].astype(cat)

    # Sortér
    result = result.sort_values("Mødetype_std")

    # Omdøb kolonnen tilbage til Mødetype
    result = result.rename(columns={"Mødetype_std": "Mødetype"})

    return result



# ---------- PLOTTES ----------

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


# ---------- PDF-GENERERING ----------

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
            fig = plot_bar(groups_per_person, "Antal grupper", "Antal personer", "Antal grupper man er medlem af", rotation=0)
            pdf.savefig(fig)
            plt.close(fig)

        # Side 13: Grupper pr. region pr. 100.000 borgere
        if groups_per_region_norm is not None and not groups_per_region_norm.empty:
            fig = plot_bar(groups_per_region_norm, "Region", "Grupper pr. 100.000 borgere", "Grupper pr. 100.000 borgere", rotation=45)
            pdf.savefig(fig)
            plt.close(fig)

    buf.seek(0)
    return buf


# ---------- STREAMLIT-APP ----------

def main():
    st.title("DGE-rapportgenerator (Regioner, grupper og møder)")

    st.markdown(
        "Upload de tre Excel-filer (grupper, møder, medlemmer) og vælg periode og regioner.\n"
        "Appen genererer både visuelle figurer og en samlet PDF-rapport, der ligner din eksisterende rapport."
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

    # Indlæs og rens data
    groups_df_raw = load_excel(groups_file)
    meetings_df_raw = load_excel(meetings_file)
    seats_df_raw = load_excel(seats_file)

    groups_df = clean_groups_df(groups_df_raw)
    meetings_df = clean_meetings_df(meetings_df_raw)
    seats_df = clean_seats_df(seats_df_raw)

    # ⭐ DEBUG: Se arkiveringsdatoer og gruppenavne
    #st.write("DEBUG – grupper og arkiveringsdatoer:", groups_df[["Gruppenavn", "Dato for arkivering"]])

    # Vælg periode
    st.subheader("Periodevalg")
    min_date = meetings_df["Starttidspunkt"].min()
    max_date = meetings_df["Starttidspunkt"].max()
    col_start, col_end = st.columns(2)
    with col_start:
        start_date = st.date_input("Startdato", value=min_date.date() if pd.notnull(min_date) else datetime.today().date())
    with col_end:
        end_date = st.date_input("Slutdato", value=max_date.date() if pd.notnull(max_date) else datetime.today().date())

    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())

    # Vælg region(er)
    st.subheader("Regioner")
    region_options = get_region_options(groups_df, meetings_df, seats_df)
    selected_regions = st.multiselect("Vælg region(er)", options=region_options, default=region_options)

    # Filtrér på regioner
    if selected_regions:
        groups_df = groups_df[groups_df["Region"].isin(selected_regions)]
        meetings_df = meetings_df[meetings_df["Region"].isin(selected_regions)]
        seats_df = seats_df[seats_df["Region"].isin(selected_regions)]

    # ⭐ Fjern gruppen "Gruppeledere" fra groups_df
    groups_df = groups_df[
        groups_df["Gruppenavn"].astype(str).str.strip().str.lower() != "gruppeledere"
]


# Filtrér møder på periode
    meetings_period_df = filter_by_period(meetings_df, start_dt, end_dt)

# ⭐ Gem en version med ALLE statusser til tabel 6
    meetings_period_all_status = meetings_period_df.copy()


# ⭐ Normalisér status
    meetings_period_df["Status"] = (
        meetings_period_df["Status"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

# ⭐ Filtrér på status
    valid_status = ["godkendt"]
    meetings_period_df = meetings_period_df[
    meetings_period_df["Status"].isin(valid_status)
    ]

# ⭐ Fjern møder i gruppen "Gruppeledere"
    meetings_period_df = meetings_period_df[
    meetings_period_df["Gruppenavn"].astype(str).str.strip().str.lower() != "gruppeledere"
    ]

    if meetings_period_df.empty:
        st.warning("Ingen møder i den valgte periode/region(er). Prøv at ændre periode eller region.")
        return



    if meetings_period_df is None or meetings_period_df.empty:
        st.warning("Ingen møder i den valgte periode/region(er). Prøv at ændre periode eller region.")
        return

    # ---------- BEREGN ALLE AGGREGATER ----------

    basic_stats = compute_basic_stats(meetings_period_df)
    meetings_by_type = compute_meetings_by_type_approved(meetings_period_df)
    meetings_by_participant_bins = compute_meetings_by_participant_bins(meetings_period_df)
    meetings_per_group = compute_meetings_per_group(meetings_period_df)
    group_size_dist = compute_group_size_distribution(groups_df)
    group_size_by_type = compute_group_size_by_type(groups_df)
    meetings_by_weekday = compute_meetings_by_weekday(meetings_period_df)
    meeting_status = compute_meeting_status(meetings_period_all_status)
    groups_without_meetings = compute_groups_without_meetings(groups_df, meetings_df, start_dt, end_dt)
    closed_groups = compute_closed_groups_this_year(groups_df, year=start_date.year)
    member_types = compute_member_types(seats_df)
    groups_per_person = compute_groups_per_person(seats_df)
    groups_per_region_norm = compute_groups_per_region_per_100k(groups_df)

    # ---------- VISUEL VISNING I APPEN ----------

    st.subheader("Overblik")
    total_meetings, total_participant_days = basic_stats
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Antal møder i perioden", total_meetings)
    with col_b:
        st.metric("Antal kursusdeltagerdage i perioden", total_participant_days)

    st.subheader("Tabel 1: Godkendte møder fordelt på type")
    if not meetings_by_type.empty:
        st.bar_chart(meetings_by_type.set_index("Mødetype")["Antal møder"])

    st.subheader("Tabel 2: Mødedeltagelse (antal deltagere pr. møde)")
    if not meetings_by_participant_bins.empty:
        st.bar_chart(meetings_by_participant_bins.set_index("Deltagerkategori")["Antal møder"])

    st.subheader("Tabel 3: Gruppernes mødeaktivitet (antal møder pr. gruppe)")
    if not meetings_per_group.empty:
        st.bar_chart(meetings_per_group.set_index("Mødekategori")["Antal grupper"])



    st.subheader("Tabel 4: Gruppestørrelse fordelt på gruppetype")
    if not group_size_by_type.empty:
        pivot = group_size_by_type.pivot(index="Størrelseskategori", columns="Gruppetyper", values="Antal grupper").fillna(0)
        st.bar_chart(pivot)

    st.subheader("Tabel 5: Mødedage")
    if not meetings_by_weekday.empty:
        st.bar_chart(meetings_by_weekday.set_index("Ugedag")["Antal møder"])

    st.subheader("Tabel 6: Mødestatus")
    if not meeting_status.empty:
        st.bar_chart(meeting_status.set_index("Status")["Antal møder"])

    st.subheader("Grupper uden godkendte møder i perioden")
    if groups_without_meetings is not None and not groups_without_meetings.empty:
        st.dataframe(groups_without_meetings)

    st.subheader(f"Grupper lukket i {start_date.year}")
    if closed_groups is not None and not closed_groups.empty:
        st.dataframe(closed_groups)

    st.subheader("Medlemstyper (stillingsbetegnelser)")
    if member_types is not None and not member_types.empty:
        st.dataframe(member_types)

    st.subheader("Antal grupper man er medlem af")
    if groups_per_person is not None and not groups_per_person.empty:
        st.bar_chart(groups_per_person.set_index("Antal grupper")["Antal personer"])

    st.subheader("Grupper pr. region pr. 100.000 borgere")
    if groups_per_region_norm is not None and not groups_per_region_norm.empty:
        st.dataframe(groups_per_region_norm)

    # ---------- PDF DOWNLOAD ----------

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
