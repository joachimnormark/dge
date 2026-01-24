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

TEST_DEMO_PATTERNS = ["test", "demo"]  # bruges til at filtrere grupper/møder


# ---------- HJÆLPEFUNKTIONER ----------

def parse_date_series(s):
    """Forsøger at parse datoer robust fra forskellige formater."""
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def load_excel(uploaded_file):
    if uploaded_file is None:
        return None
    # engine behøver ikke angives, men kan hvis du vil:
    # return pd.read_excel(uploaded_file, engine="openpyxl")
    return pd.read_excel(uploaded_file)


def name_matches_test_demo(name: str) -> bool:
    if not isinstance(name, str):
        return False
    lower = name.lower()
    return any(p in lower for p in TEST_DEMO_PATTERNS)


# ---------- RENSNING ----------

def clean_groups_df(df):
    """
    Rens grupper:
    - standardiser kolonnenavne
    - parse datoer
    - konverter 'Antal medlemmer'
    - fjern grupper der hedder noget med 'test' eller 'demo'
    - fjern gruppen 'Gruppeledere' fra al aktivitet
    Returnerer (renset_df, fjernet_df)
    """
    if df is None:
        return None, pd.DataFrame()

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Dato for arkivering
    if "Dato for arkivering" in df.columns:
        df["Dato for arkivering"] = df["Dato for arkivering"].replace("-", np.nan)
        df["Dato for arkivering"] = parse_date_series(df["Dato for arkivering"])

    # Antal medlemmer
    if "Antal medlemmer" in df.columns:
        df["Antal medlemmer"] = pd.to_numeric(df["Antal medlemmer"], errors="coerce")

    # Filtrering på navn
    removed_rows = []

    if "Gruppenavn" in df.columns:
        mask_test_demo = df["Gruppenavn"].apply(name_matches_test_demo)
        mask_gruppeledere = df["Gruppenavn"].astype(str).str.strip().eq("Gruppeledere")

        removed = df[mask_test_demo | mask_gruppeledere].copy()
        kept = df[~(mask_test_demo | mask_gruppeledere)].copy()

        removed["Fjernelsesårsag"] = np.where(
            mask_gruppeledere[mask_test_demo | mask_gruppeledere],
            "Gruppeledere-gruppe",
            "Test/demo-gruppe",
        )
        removed_rows.append(removed)
    else:
        kept = df
        removed = pd.DataFrame()

    removed_all = pd.concat(removed_rows, ignore_index=True) if removed_rows else pd.DataFrame()
    return kept.reset_index(drop=True), removed_all.reset_index(drop=True)


def clean_meetings_df(df, valid_groups_df=None):
    """
    Rens møder:
    - standardiser kolonnenavne
    - parse datoer
    - konverter 'Antal deltagere'
    - fjern møder i grupper der hedder noget med 'test' eller 'demo'
    - fjern møder i gruppen 'Gruppeledere'
    - fjern møder hvor mødetitel indeholder 'test'/'demo'
    - hvis valid_groups_df er givet, fjern møder for grupper der ikke findes dér
    Returnerer (renset_df, fjernet_df)
    """
    if df is None:
        return None, pd.DataFrame()

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Datoer
    if "Starttidspunkt" in df.columns:
        df["Starttidspunkt"] = parse_date_series(df["Starttidspunkt"])
    if "Sluttidspunkt" in df.columns:
        df["Sluttidspunkt"] = parse_date_series(df["Sluttidspunkt"])

    # Antal deltagere
    if "Antal deltagere" in df.columns:
        df["Antal deltagere"] = (
            pd.to_numeric(df["Antal deltagere"], errors="coerce").fillna(0).astype(int)
        )

    removed_parts = []

    # Filtrering på gruppenavn og mødetitel
    mask_remove = pd.Series(False, index=df.index)

    if "Gruppenavn" in df.columns:
        mask_test_demo_group = df["Gruppenavn"].apply(name_matches_test_demo)
        mask_gruppeledere_group = df["Gruppenavn"].astype(str).str.strip().eq("Gruppeledere")
        mask_remove |= mask_test_demo_group | mask_gruppeledere_group

    if "Mødetitel" in df.columns:
        mask_test_demo_title = df["Mødetitel"].apply(name_matches_test_demo)
        mask_remove |= mask_test_demo_title
    else:
        mask_test_demo_title = pd.Series(False, index=df.index)

    removed_named = df[mask_remove].copy()
    if not removed_named.empty:
        removed_named["Fjernelsesårsag"] = np.select(
            [
                "Gruppenavn" in df.columns
                and df.loc[removed_named.index, "Gruppenavn"].astype(str).str.strip().eq("Gruppeledere"),
                "Gruppenavn" in df.columns
                and df.loc[removed_named.index, "Gruppenavn"].apply(name_matches_test_demo),
                "Mødetitel" in df.columns
                and df.loc[removed_named.index, "Mødetitel"].apply(name_matches_test_demo),
            ],
            [
                "Møde i Gruppeledere-gruppe",
                "Møde i test/demo-gruppe",
                "Mødetitel indeholder test/demo",
            ],
            default="Filtreret møde",
        )
        removed_parts.append(removed_named)

    kept = df[~mask_remove].copy()

    # Fjern møder for grupper der ikke findes i valid_groups_df (hvis givet)
    if valid_groups_df is not None and "Gruppenavn" in kept.columns and "Gruppenavn" in valid_groups_df.columns:
        valid_names = set(valid_groups_df["Gruppenavn"].dropna().astype(str))
        mask_invalid_group = ~kept["Gruppenavn"].astype(str).isin(valid_names)
        removed_invalid = kept[mask_invalid_group].copy()
        if not removed_invalid.empty:
            removed_invalid["Fjernelsesårsag"] = "Gruppenavn findes ikke i renset gruppeliste"
            removed_parts.append(removed_invalid)
        kept = kept[~mask_invalid_group].copy()

    removed_all = pd.concat(removed_parts, ignore_index=True) if removed_parts else pd.DataFrame()
    return kept.reset_index(drop=True), removed_all.reset_index(drop=True)


def clean_seats_df(df):
    """
    Rens sæder/medlemmer:
    - standardiser kolonnenavne
    - konverter 'Antal møder'
    - fjern brugere der er medlem af 'Gruppeledere' flere gange
    - for de resterende: fjern 'Gruppeledere' fra Medlemskaber (aktivitet i gruppen skal ikke tælle)
    - fjern også medlemskaber der hedder noget med 'test' eller 'demo'
    Returnerer (renset_df, fjernet_df)
    """
    if df is None:
        return None, pd.DataFrame()

    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    if "Antal møder" in df.columns:
        df["Antal møder"] = (
            pd.to_numeric(df["Antal møder"], errors="coerce").fillna(0).astype(int)
        )

    if "Medlemskaber" not in df.columns:
        return df.reset_index(drop=True), pd.DataFrame()

    # Brugere med Gruppeledere flere gange
    medlemskaber_str = df["Medlemskaber"].fillna("").astype(str)
    gl_count = medlemskaber_str.str.count(r"\bGruppeledere\b")
    mask_gl_multi = gl_count > 1

    removed_gl_multi = df[mask_gl_multi].copy()
    if not removed_gl_multi.empty:
        removed_gl_multi["Fjernelsesårsag"] = "Medlem af Gruppeledere flere gange"

    kept = df[~mask_gl_multi].copy()

    # Rens Medlemskaber for Gruppeledere og test/demo-grupper
    def clean_memberships(s: str) -> str:
        parts = [p.strip() for p in str(s).split(",") if p.strip() != ""]
        cleaned = []
        for p in parts:
            if p == "Gruppeledere":
                continue
            if name_matches_test_demo(p):
                continue
            cleaned.append(p)
        return ", ".join(cleaned)

    kept["Medlemskaber"] = kept["Medlemskaber"].fillna("").apply(clean_memberships)

    return kept.reset_index(drop=True), removed_gl_multi.reset_index(drop=True)


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

    # Engelske ugedage
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

    def count_groups(s: str) -> int:
        parts = [p.strip() for p in str(s).split(",") if p.strip() != ""]
        return len(parts)

    df["Antal grupper"] = df["Medlemskaber"].fillna("").apply(count_groups)
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
            fig = plot_bar(
                meetings_by_participant_bins,
                "Deltagerkategori",
                "Antal møder",
                "Mødedeltagelse",
                rotation=0,
            )
            pdf.savefig(fig)
            plt.close(fig)

        # Side 4: Gruppernes mødeaktivitet
        if not meetings_per_group.empty:
            fig = plot_bar(
                meetings_per_group,
                "Mødekategori",
                "Antal grupper",
                "Gruppernes mødeaktivitet",
                rotation=0,
            )
            pdf.savefig(fig)
            plt.close(fig)

        # Side 5: Gruppestørrelse
        if not group_size_dist.empty:
            fig = plot_bar(
                group_size_dist,
                "Størrelseskategori",
                "Antal grupper",
                "Gruppestørrelse",
                rotation=0,
            )
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
            fig = plot_bar(
                member_types,
                "Stillingsbetegnelse",
                "Antal",
                "Medlemstyper",
                rotation=90,
            )
            pdf.savefig(fig)
            plt.close(fig)

        # Side 12: Antal grupper man er medlem af
        if groups_per_person is not None and not groups_per_person.empty:
            fig = plot_bar(
                groups_per_person,
                "Antal grupper",
                "Antal personer",
                "Antal grupper man er medlem af",
                rotation=0,
            )
            pdf.savefig(fig)
            plt.close(fig)

        # Side 13: Grupper pr. region pr. 100.000 borgere
        if groups_per_region_norm is not None and not groups_per_region_norm.empty:
            fig = plot_bar(
                groups_per_region_norm,
                "Region",
                "Grupper pr. 100.000 borgere",
                "Grupper pr. 100.000 borgere",
                rotation=45,
            )
            pdf.savefig(fig)
            plt.close(fig)

    buf.seek(0)
    return buf


def build_removed_pdf(removed_groups, removed_meetings, removed_seats):
    """
    PDF med de data der er sorteret fra (grupper, møder, sæder).
    """
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        # Fjernede grupper
        if removed_groups is not None and not removed_groups.empty:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title("Fjernede grupper", loc="left")
            y = 0.95
            for _, row in removed_groups.iterrows():
                name = row.get("Gruppenavn", "")
                reason = row.get("Fjernelsesårsag", "")
                ax.text(0.05, y, f"- {name} ({reason})", fontsize=9, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

        # Fjernede møder
        if removed_meetings is not None and not removed_meetings.empty:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title("Fjernede møder", loc="left")
            y = 0.95
            for _, row in removed_meetings.iterrows():
                g = row.get("Gruppenavn", "")
                title = row.get("Mødetitel", "")
                reason = row.get("Fjernelsesårsag", "")
                ax.text(0.05, y, f"- {g}: {title} ({reason})", fontsize=8, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

        # Fjernede sæder/medlemmer
        if removed_seats is not None and not removed_seats.empty:
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis("off")
            ax.set_title("Fjernede medlemmer (sæder)", loc="left")
            y = 0.95
            for _, row in removed_seats.iterrows():
                navn = f"{row.get('Fornavn', '')} {row.get('Efternavn', '')}".strip()
                medlemskaber = row.get("Medlemskaber", "")
                reason = row.get("Fjernelsesårsag", "")
                ax.text(0.05, y, f"- {navn}: {medlemskaber} ({reason})", fontsize=8, va="top")
                y -= 0.02
                if y < 0.05:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax = plt.subplots(figsize=(8.27, 11.69))
                    ax.axis("off")
                    y = 0.95
            pdf.savefig(fig)
            plt.close(fig)

    buf.seek(0)
    return buf


# ---------- STREAMLIT-APP ----------

def main():
    st.title("DGE-rapportgenerator (Regioner, grupper og møder)")

    st.markdown(
        "Upload de tre Excel-filer (grupper, møder, medlemmer) og vælg periode og regioner.\n\n"
        "Appen genererer både visuelle figurer og en samlet PDF-rapport, der ligner din eksisterende rapport.\n\n"
        "_Det er kolonneoverskrifterne der er vigtige – rækkefølgen af kolonnerne er ikke afgørende._"
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

    # Indlæs rå data
    groups_df_raw = load_excel(groups_file)
    meetings_df_raw = load_excel(meetings_file)
    seats_df_raw = load_excel(seats_file)

    # Rens data (inkl. fjernede rækker)
    groups_df, removed_groups = clean_groups_df(groups_df_raw)
    meetings_df, removed_meetings = clean_meetings_df(meetings_df_raw, valid_groups_df=groups_df)
    seats_df, removed_seats = clean_seats_df(seats_df_raw)

    # Vælg periode
    st.subheader("Periodevalg")
    if meetings_df is None or meetings_df.empty or "Starttidspunkt" not in meetings_df.columns:
        st.error("Kunne ikke finde gyldige mødedatoer i mødefilen.")
        return

    min_date = meetings_df["Starttidspunkt"].min()
    max_date = meetings_df["Starttidspunkt"].max()
    col_start, col_end = st.columns(2)
    with col_start:
        start_date = st.date_input(
            "Startdato",
            value=min_date.date() if pd.notnull(min_date) else datetime.today().date(),
        )
    with col_end:
        end_date = st.date_input(
            "Slutdato",
            value=max_date.date() if pd.notnull(max_date) else datetime.today().date(),
        )

    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())

    # Vælg region(er)
    st.subheader("Regioner")
    region_options = get_region_options(groups_df, meetings_df, seats_df)
    selected_regions = st.multiselect("Vælg region(er)", options=region_options, default=region_options)

    # Filtrér på regioner
    if selected_regions:
        if groups_df is not None and "Region" in groups_df.columns:
            groups_df = groups_df[groups_df["Region"].isin(selected_regions)]
        if meetings_df is not None and "Region" in meetings_df.columns:
            meetings_df = meetings_df[meetings_df["Region"].isin(selected_regions)]
        if seats_df is not None and "Region" in seats_df.columns:
            seats_df = seats_df[seats_df["Region"].isin(selected_regions)]

    # Filtrér møder på periode
    meetings_period_df = filter_by_period(meetings_df, start_dt, end_dt)

    if meetings_period_df is None or meetings_period_df.empty:
        st.warning("Ingen møder i den valgte periode/region(er). Prøv at ændre periode eller region.")
        return

    # ---------- BEREGN ALLE AGGREGATER ----------

    basic_stats = compute_basic_stats(meetings_period_df)
    meetings_by_type = compute_meetings_by_type(meetings_period_df)
    meetings_by_participant_bins = compute_meetings_by_participant_bins(meetings_period_df)
    meetings_per_group = compute_meetings_per_group(meetings_period_df)
    group_size_dist = compute_group_size_distribution(groups_df)
    group_size_by_type = compute_group_size_by_type(groups_df)
    meetings_by_weekday = compute_meetings_by_weekday(meetings_period_df)
    meeting_status = compute_meeting_status(meetings_period_df)
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

    st.subheader("Møder fordelt på type")
    if not meetings_by_type.empty:
        st.bar_chart(meetings_by_type.set_index("Mødetype")["Antal møder"])

    st.subheader("Mødedeltagelse (antal deltagere pr. møde)")
    if not meetings_by_participant_bins.empty:
        st.bar_chart(meetings_by_participant_bins.set_index("Deltagerkategori")["Antal møder"])

    st.subheader("Gruppernes mødeaktivitet (antal møder pr. gruppe)")
    if not meetings_per_group.empty:
        st.bar_chart(meetings_per_group.set_index("Mødekategori")["Antal grupper"])

    st.subheader("Gruppestørrelse")
    if not group_size_dist.empty:
        st.bar_chart(group_size_dist.set_index("Størrelseskategori")["Antal grupper"])

    st.subheader("Gruppestørrelse fordelt på gruppetype")
    if not group_size_by_type.empty:
        pivot = group_size_by_type.pivot(
            index="Størrelseskategori",
            columns="Gruppetyper",
            values="Antal grupper",
        ).fillna(0)
        st.bar_chart(pivot)

    st.subheader("Mødedage")
    if not meetings_by_weekday.empty:
        st.bar_chart(meetings_by_weekday.set_index("Ugedag")["Antal møder"])

    st.subheader("Mødestatus")
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

    # ---------- PDF DOWNLOAD (HOVEDRAPPORT) ----------

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

    # ---------- PDF DOWNLOAD (FJERNEDE DATA) ----------

    st.subheader("Download PDF med frasorterede data")
    removed_pdf_buffer = build_removed_pdf(removed_groups, removed_meetings, removed_seats)
    st.download_button(
        label="Download PDF med frasorterede grupper/møder/medlemmer",
        data=removed_pdf_buffer,
        file_name="dge_frasorterede_data.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
