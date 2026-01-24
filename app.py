# ---------------------------------------------
# DGE RAPPORTGENERATOR – KOMPLET NY VERSION
# Med datarensning, PDF med frasorterede data,
# ugedage uden locale, og robust filtrering.
# ---------------------------------------------

import io
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import streamlit as st

# ---------------------------------------------
# KONFIG
# ---------------------------------------------

st.set_page_config(page_title="DGE-rapport", layout="wide")

REGION_POPULATION = {
    "Region Nordjylland": 590000,
    "Region Midtjylland": 1330000,
    "Region Syddanmark": 1220000,
    "Region Hovedstaden": 1870000,
    "Region Sjælland": 840000,
}

TEST_PATTERNS = ["test", "demo", "euv"]

# ---------------------------------------------
# HJÆLPEFUNKTIONER
# ---------------------------------------------

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
        df["Antal deltagere"] = pd.to_numeric(df["Antal deltagere"], errors="coerce").fillna(0).astype(int)
    return df

def clean_seats_df(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    if "Antal møder" in df.columns:
        df["Antal møder"] = pd.to_numeric(df["Antal møder"], errors="coerce").fillna(0).astype(int)
    return df

# ---------------------------------------------
# D A T A R E N S N I N G
# ---------------------------------------------

def is_test_group(name):
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

    removed_gl = df[df["gruppeledere_count"] > 1]

    df = df[df["gruppeledere_count"] <= 1]

    removed_test = df[df["Medlemskaber"].str.lower().apply(is_test_group)]

    df = df[~df["Medlemskaber"].str.lower().apply(is_test_group)]

    df = df.drop(columns=["Medlemsliste", "gruppeledere_count"])

    removed = pd.concat([removed_gl, removed_test]).drop_duplicates()

    return df, removed

# ---------------------------------------------
# U G E D A G E  (uden locale)
# ---------------------------------------------

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

# ---------------------------------------------
# PDF MED FRASORTEREDE DATA
# ---------------------------------------------

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

        write_list("Frasorterede grupper", removed_groups["Gruppenavn"].tolist())
        write_list("Frasorterede møder (gruppenavne)", removed_meetings["Gruppenavn"].tolist())
        write_list("Frasorterede brugere (navne hvis muligt)", removed_seats.get("Fornavn", pd.Series()).astype(str).tolist())

    buf.seek(0)
    return buf

# ---------------------------------------------
# STREAMLIT APP
# ---------------------------------------------

def main():
    st.title("DGE-rapportgenerator – NY VERSION")

    col1, col2, col3 = st.columns(3)
    with col1:
        groups_file = st.file_uploader("Gruppe-data", type=["xlsx"])
    with col2:
        meetings_file = st.file_uploader("Møde-data", type=["xlsx"])
    with col3:
        seats_file = st.file_uploader("Medlemsdata", type=["xlsx"])

    if not (groups_file and meetings_file and seats_file):
        st.info("Upload alle tre filer for at fortsætte.")
        return

    groups_df_raw = clean_groups_df(load_excel(groups_file))
    meetings_df_raw = clean_meetings_df(load_excel(meetings_file))
    seats_df_raw = clean_seats_df(load_excel(seats_file))

    groups_df, removed_groups = clean_groups(groups_df_raw)
    meetings_df, removed_meetings = clean_meetings(meetings_df_raw)
    seats_df, removed_seats = clean_seats(seats_df_raw)

    st.subheader("Download frasorterede data")
    removed_pdf = build_removed_pdf(removed_groups, removed_meetings, removed_seats)
    st.download_button("Download frasorterede data (PDF)", removed_pdf, "frasorterede_data.pdf")

    st.success("Datarensning gennemført. Du kan nu fortsætte med rapportgenerering.")

    # Her indsætter du resten af din rapportlogik (samme som før)
    st.info("Resten af rapportgeneratoren indsættes her (samme som tidligere).")

if __name__ == "__main__":
    main()
