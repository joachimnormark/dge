import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

# ============================================================================
# TEKST KONFIGURATION - Rediger tekster her
# ============================================================================

TEXTS = {
    "app_title": "📊 DGE Mødeaktivitets-analyse",
    "upload_header": "Upload datafiler",
    "upload_groups": "Upload grupper-fil (Excel)",
    "upload_seats": "Upload medlemmer-fil (Excel)",
    "upload_meetings": "Upload møder-fil (Excel)",
    "period_header": "Vælg analyseperiode",
    "compare_checkbox": "Sammenlign med samme periode sidste år",
    
    # Graf titler og forklaringer
    "graph1_title": "Mødestatistik - Total og fordeling på status",
    "graph1_desc": "Viser antal møder fordelt på godkendelsestatus: Godkendt, Afventer godkendelse, Afholdt uden godkendelse, og Afvist.",
    
    "graph2_title": "Mødedage - Fordeling på ugedage",
    "graph2_desc": "Viser hvilke ugedage møder typisk afholdes på.",
    
    "graph3_title": "Mødetyper - Fordeling",
    "graph3_desc": "Viser antallet af forskellige mødetyper (DGE-møde, Supervision, moduler mv.).",
    
    "graph4_title": "Mødedeltagelse - Antal deltagere per møde",
    "graph4_desc": "Viser fordelingen af møder efter antal deltagere.",
    
    "graph5_title": "Medlemstyper - Fordeling i clusters",
    "graph5_desc": "Viser hvordan medlemmer fordeler sig på: Praktiserende læger, §-ansatte/vikarer, Uddannelseslæger, Andre, og Ej registreret.",
    
    "graph6_title": "Grupper med få møder (<4 i perioden)",
    "graph6_desc": "Liste over grupper der har holdt færre end 4 møder i perioden, sorteret efter antal møder.",
    
    "graph7_title": "Lukkede grupper i perioden",
    "graph7_desc": "Oversigt over grupper der er blevet arkiveret/lukket i den valgte periode.",
    
    "graph8_title": "Gruppestørrelse fordelt på gruppetype",
    "graph8_desc": "Viser fordelingen af gruppestørrelser for DGE, Supervision og Junior grupper.",
}

# Farveskema
COLORS = {
    "DGE": "#4169E1",           # Blå
    "Supervision": "#DC143C",   # Rød  
    "Junior": "#228B22",        # Grøn
    "Other": "#FFD700",         # Guld
    "P1": "#4169E1",            # Blå for periode 1
    "P2": "#87CEEB",            # Lyseblå for periode 2
}

# ============================================================================
# HJÆLPEFUNKTIONER
# ============================================================================

def parse_danish_date(date_str):
    """Parser danske datoformater"""
    if pd.isna(date_str) or str(date_str).strip() in ['-', '', 'nan', 'NaT']:
        return pd.NaT
    
    if isinstance(date_str, datetime):
        return date_str
    
    date_str = str(date_str).strip()
    
    # Fjern klokkeslæt
    date_str = date_str.replace('kl.', '').replace('Kl.', '').replace(',', '')
    
    # Månedsnavne
    months = {
        'januar': '01', 'februar': '02', 'marts': '03', 'april': '04',
        'maj': '05', 'juni': '06', 'juli': '07', 'august': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'december': '12'
    }
    
    for dk_month, num_month in months.items():
        if dk_month in date_str.lower():
            date_str = date_str.lower().replace(dk_month, num_month)
    
    # Erstat punktum og skråstreg med mellemrum
    date_str = date_str.replace('.', ' ').replace('/', ' ')
    
    try:
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except:
        return pd.NaT

def categorize_member_type(member_type):
    """Kategoriser medlemstyper i clusters"""
    if pd.isna(member_type) or str(member_type).strip() == "":
        return "Ej registreret"
    
    member_type = str(member_type).strip()
    
    if member_type == "Alment praktiserende læge":
        return "Praktiserende læger"
    
    speciallaege_types = [
        "Ansat speciallæge i alm. med. - § 13, stk. 2",
        "Ansat speciallæge i alm. med. - § 13, stk. 5",
        "Ansat speciallæge i alm. med. - § 23, stk. 1",
        "Ansat speciallæge i alm. med. - § 23, stk. 2",
        "Ansat speciallæge i alm. med. - § 24",
        "Ansat speciallæge i alm. med. - § 26",
        "Assisterende speciallæge",
        "Vikar i almen praksis",
    ]
    if member_type in speciallaege_types:
        return "§-ansatte, vikarer mv"
    
    udd_types = [
        "Praksisamanuensis (Fase 1)",
        "Praksisamanuensis (Fase 2)",
        "Praksisamanuensis (Fase 3)",
        "Introduktionsamanuensis (Almen Praksis)",
        "KBU - Læge (trin 1)",
        "Læge (trin 1)",
        "Læge (trin 2)",
        "Hoveduddannelsesstilling - Læge (trin 1)",
        "Hoveduddannelsesstilling - Læge (trin 2)",
    ]
    if member_type in udd_types:
        return "Uddannelseslæger"
    
    return "Andre"

def filter_gruppeledere(df, group_name_col='Gruppenavn'):
    """Filtrer 'Gruppeledere' gruppe fra dataframe"""
    if df is None or df.empty:
        return df
    return df[df[group_name_col].astype(str).str.strip().str.lower() != 'gruppeledere'].copy()

def filter_members_gruppeledere(seats_df):
    """Filtrer medlemskaber i Gruppeledere fra seats"""
    if seats_df is None or seats_df.empty:
        return seats_df
    
    df = seats_df.copy()
    if 'Medlemskaber' in df.columns:
        # Fjern "Gruppeledere" fra medlemskabslisten
        df['Medlemskaber'] = df['Medlemskaber'].astype(str).apply(
            lambda x: ','.join([g.strip() for g in x.split(',') if 'gruppeledere' not in g.lower()])
        )
    return df

def get_group_type_color(group_type):
    """Få farve for gruppetype"""
    if pd.isna(group_type):
        return COLORS["Other"]
    
    group_type = str(group_type).strip()
    
    if "DGE" in group_type or "dge" in group_type.lower():
        return COLORS["DGE"]
    elif "Supervision" in group_type or "supervision" in group_type.lower():
        return COLORS["Supervision"]
    elif "Junior" in group_type or "junior" in group_type.lower():
        return COLORS["Junior"]
    else:
        return COLORS["Other"]

# ============================================================================
# ANALYSE FUNKTIONER
# ============================================================================

def analyze_meeting_status(meetings_df):
    """Analysér møder fordelt på status"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    # Map status til dansk
    status_map = {
        'Godkendt': 'Godkendt',
        'godkendt': 'Godkendt',
        '-': 'Afventer godkendelse',
        'Afvist': 'Afvist',
        'afvist': 'Afvist',
        'Afsluttet': 'Afholdt uden godkendelse',
        'afsluttet': 'Afholdt uden godkendelse',
    }
    
    df = meetings_df.copy()
    df['Status_mapped'] = df['Status'].map(status_map).fillna('Andet')
    
    return df['Status_mapped'].value_counts().reset_index()

def analyze_weekday_distribution(meetings_df):
    """Analysér møder fordelt på ugedage"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    df = meetings_df.copy()
    df['Weekday'] = pd.to_datetime(df['Starttidspunkt']).dt.day_name()
    
    weekday_map = {
        'Monday': 'Mandag',
        'Tuesday': 'Tirsdag',
        'Wednesday': 'Onsdag',
        'Thursday': 'Torsdag',
        'Friday': 'Fredag',
        'Saturday': 'Lørdag',
        'Sunday': 'Søndag'
    }
    
    df['Weekday_DK'] = df['Weekday'].map(weekday_map)
    
    # Sorter i ugedags-rækkefølge
    weekday_order = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag', 'Søndag']
    result = df['Weekday_DK'].value_counts().reindex(weekday_order, fill_value=0).reset_index()
    result.columns = ['Ugedag', 'Antal']
    
    return result

def analyze_meeting_types(meetings_df):
    """Analysér møder fordelt på type"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    result = meetings_df['Mødetype'].value_counts().reset_index()
    result.columns = ['Mødetype', 'Antal']
    return result

def analyze_participant_distribution(meetings_df):
    """Analysér møder fordelt på antal deltagere"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    bins = [0, 4, 6, 8, 10, 12, 14, 999]
    labels = ['2-4', '5-6', '7-8', '9-10', '11-12', '13-14', '15+']
    
    df = meetings_df.copy()
    df['Deltagerkategori'] = pd.cut(df['Antal deltagere'], bins=bins, labels=labels, right=True)
    
    result = df['Deltagerkategori'].value_counts().sort_index().reset_index()
    result.columns = ['Deltagerkategori', 'Antal']
    return result

def analyze_member_types(seats_df):
    """Analysér medlemstyper i clusters"""
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    
    df = seats_df.copy()
    if 'Stillingsbetegnelse' not in df.columns:
        return pd.DataFrame()
    
    df['Cluster'] = df['Stillingsbetegnelse'].apply(categorize_member_type)
    
    result = df['Cluster'].value_counts().reset_index()
    result.columns = ['Medlemstype', 'Antal']
    return result

def analyze_groups_with_few_meetings(groups_df, meetings_df, start_date, end_date):
    """Find grupper med færre end 4 møder"""
    if groups_df is None or groups_df.empty or meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    # Tæl møder per gruppe
    meeting_counts = meetings_df.groupby('Gruppenavn').size().reset_index(name='Antal møder')
    
    # Merge med groups for at få gruppetype
    result = groups_df[['Gruppenavn', 'Gruppetyper', 'Status', 'Dato for arkivering']].merge(
        meeting_counts, on='Gruppenavn', how='left'
    )
    
    result['Antal møder'] = result['Antal møder'].fillna(0).astype(int)
    
    # Filtrer til <4 møder
    result = result[result['Antal møder'] < 4].copy()
    
    # Tjek om arkiveret i perioden
    result['Arkiveret i periode'] = result['Dato for arkivering'].apply(
        lambda x: 'Ja' if pd.notna(x) and start_date <= x <= end_date else 'Nej'
    )
    
    # Sorter efter antal møder
    result = result.sort_values('Antal møder', ascending=True)
    
    return result[['Gruppenavn', 'Gruppetyper', 'Antal møder', 'Status', 'Arkiveret i periode']]

def analyze_closed_groups(groups_df, start_date, end_date):
    """Find grupper lukket i perioden"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    df = groups_df.copy()
    
    if 'Dato for arkivering' not in df.columns:
        return pd.DataFrame()
    
    # Filtrer til grupper arkiveret i perioden
    mask = (df['Dato for arkivering'] >= start_date) & (df['Dato for arkivering'] <= end_date)
    result = df[mask].copy()
    
    if result.empty:
        return pd.DataFrame()
    
    return result[['Gruppenavn', 'Gruppetyper', 'Dato for arkivering', 'Antal medlemmer']].sort_values('Dato for arkivering')

def analyze_group_size_by_type(groups_df):
    """Analysér gruppestørrelse fordelt på type"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    df = groups_df.copy()
    
    if 'Antal medlemmer' not in df.columns or 'Gruppetyper' not in df.columns:
        return pd.DataFrame()
    
    bins = [0, 5, 8, 10, 12, 999]
    labels = ['1-5', '6-8', '9-10', '11-12', '13+']
    
    df['Størrelseskategori'] = pd.cut(df['Antal medlemmer'], bins=bins, labels=labels, right=True)
    
    result = df.groupby(['Størrelseskategori', 'Gruppetyper']).size().reset_index(name='Antal')
    
    return result

# ============================================================================
# VISUALISERINGS FUNKTIONER
# ============================================================================

def create_comparison_bar_chart(data_p1, data_p2, x_col, y_col, title, description, color=None):
    """Lav sammenlignende søjlediagram"""
    fig = go.Figure()
    
    if color is None:
        color = COLORS["P1"]
    
    # Merge data
    merged = data_p1.merge(data_p2, on=x_col, how='outer', suffixes=('_P1', '_P2')).fillna(0)
    
    y_col_p1 = f"{y_col}_P1"
    y_col_p2 = f"{y_col}_P2"
    
    # P1 bars
    fig.add_trace(go.Bar(
        x=merged[x_col],
        y=merged[y_col_p1],
        name='Periode 1',
        marker_color=COLORS["P1"]
    ))
    
    # P2 bars
    fig.add_trace(go.Bar(
        x=merged[x_col],
        y=merged[y_col_p2],
        name='Periode 2',
        marker_color=COLORS["P2"]
    ))
    
    fig.update_layout(
        title=title,
        xaxis_title=x_col,
        yaxis_title=y_col,
        barmode='group',
        height=500,
        showlegend=True,
        hovermode='x unified'
    )
    
    return fig, description

def create_single_bar_chart(data, x_col, y_col, title, description, color_col=None):
    """Lav enkelt søjlediagram"""
    fig = go.Figure()
    
    if color_col and color_col in data.columns:
        colors = data[color_col].apply(get_group_type_color)
    else:
        colors = COLORS["P1"]
    
    fig.add_trace(go.Bar(
        x=data[x_col],
        y=data[y_col],
        marker_color=colors,
        showlegend=False
    ))
    
    fig.update_layout(
        title=title,
        xaxis_title=x_col,
        yaxis_title=y_col,
        height=500,
        hovermode='x'
    )
    
    return fig, description

def create_pie_chart(data, names_col, values_col, title, description):
    """Lav pie chart"""
    fig = px.pie(
        data,
        names=names_col,
        values=values_col,
        title=title,
        hole=0.3
    )
    
    fig.update_layout(height=500)
    
    return fig, description

# ============================================================================
# PDF GENERATION
# ============================================================================

def generate_pdf(charts_data, period_info):
    """Generér PDF rapport"""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    
    # Forside
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, height - 100, "DGE Moede aktivitets-rapport")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 130, f"Periode: {period_info}")
    c.drawString(50, height - 150, f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    
    # Tilføj grafer (simplified - kun titles da images kræver mere arbejde)
    y_pos = height - 200
    for i, (title, desc) in enumerate(charts_data):
        c.drawString(50, y_pos, f"{i+1}. {title}")
        y_pos -= 30
        if y_pos < 100:
            c.showPage()
            y_pos = height - 50
    
    c.save()
    buffer.seek(0)
    return buffer

# ============================================================================
# MAIN APP
# ============================================================================

st.set_page_config(page_title=TEXTS["app_title"], layout="wide")

def main():
    st.title(TEXTS["app_title"])
    
    # ========== FILE UPLOAD ==========
    st.header(TEXTS["upload_header"])
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        groups_file = st.file_uploader(TEXTS["upload_groups"], type=['xlsx', 'xls'])
    with col2:
        seats_file = st.file_uploader(TEXTS["upload_seats"], type=['xlsx', 'xls'])
    with col3:
        meetings_file = st.file_uploader(TEXTS["upload_meetings"], type=['xlsx', 'xls'])
    
    if not all([groups_file, seats_file, meetings_file]):
        st.info("👆 Upload alle 3 datafiler for at fortsætte")
        return
    
    # Load data
    try:
        groups_df = pd.read_excel(groups_file)
        seats_df = pd.read_excel(seats_file)
        meetings_df = pd.read_excel(meetings_file)
        
        st.success(f"✅ Data indlæst: {len(groups_df)} grupper, {len(seats_df)} medlemmer, {len(meetings_df)} møder")
    except Exception as e:
        st.error(f"Fejl ved indlæsning af data: {e}")
        return
    
    # ========== DATA CLEANING ==========
    
    # Parse datoer
    if 'Starttidspunkt' in meetings_df.columns:
        meetings_df['Starttidspunkt'] = meetings_df['Starttidspunkt'].apply(parse_danish_date)
    
    if 'Dato for arkivering' in groups_df.columns:
        groups_df['Dato for arkivering'] = groups_df['Dato for arkivering'].apply(parse_danish_date)
    
    # Filtrer Gruppeledere
    groups_df = filter_gruppeledere(groups_df)
    meetings_df = filter_gruppeledere(meetings_df)
    seats_df = filter_members_gruppeledere(seats_df)
    
    # ========== PERIODE VALG ==========
    st.header(TEXTS["period_header"])
    
    col_start, col_end = st.columns(2)
    
    with col_start:
        start_date = st.date_input("Startdato", value=datetime(2024, 1, 1))
    with col_end:
        end_date = st.date_input("Slutdato", value=datetime(2024, 12, 31))
    
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())
    
    # Sammenligning med sidste år
    compare_previous_year = st.checkbox(TEXTS["compare_checkbox"])
    
    if compare_previous_year:
        start_dt_p2 = start_dt - relativedelta(years=1)
        end_dt_p2 = end_dt - relativedelta(years=1)
    
    # ========== FILTER DATA ==========
    
    meetings_p1 = meetings_df[
        (meetings_df['Starttidspunkt'] >= start_dt) & 
        (meetings_df['Starttidspunkt'] <= end_dt)
    ].copy()
    
    if compare_previous_year:
        meetings_p2 = meetings_df[
            (meetings_df['Starttidspunkt'] >= start_dt_p2) & 
            (meetings_df['Starttidspunkt'] <= end_dt_p2)
        ].copy()
    
    if meetings_p1.empty:
        st.warning("Ingen møder fundet i den valgte periode")
        return
    
    # ========== ANALYSER ==========
    
    st.header("Analyseresultater")
    
    # Metrics
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric("Antal møder (P1)", len(meetings_p1))
        if compare_previous_year:
            st.metric("Antal møder (P2)", len(meetings_p2))
    
    with col_m2:
        st.metric("Antal deltagerdage (P1)", meetings_p1['Antal deltagere'].sum())
        if compare_previous_year:
            st.metric("Antal deltagerdage (P2)", meetings_p2['Antal deltagere'].sum())
    
    with col_m3:
        st.metric("Unikke grupper (P1)", meetings_p1['Gruppenavn'].nunique())
        if compare_previous_year:
            st.metric("Unikke grupper (P2)", meetings_p2['Gruppenavn'].nunique())
    
    st.markdown("---")
    
    # Graf 1: Møde status
    st.subheader(TEXTS["graph1_title"])
    st.caption(TEXTS["graph1_desc"])
    
    status_p1 = analyze_meeting_status(meetings_p1)
    
    if compare_previous_year:
        status_p2 = analyze_meeting_status(meetings_p2)
        fig, desc = create_comparison_bar_chart(
            status_p1, status_p2, 'Status_mapped', 'count',
            TEXTS["graph1_title"], TEXTS["graph1_desc"]
        )
    else:
        status_p1.columns = ['Status', 'Antal']
        fig, desc = create_single_bar_chart(
            status_p1, 'Status', 'Antal',
            TEXTS["graph1_title"], TEXTS["graph1_desc"]
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Graf 2: Ugedage
    st.subheader(TEXTS["graph2_title"])
    st.caption(TEXTS["graph2_desc"])
    
    weekday_p1 = analyze_weekday_distribution(meetings_p1)
    
    if compare_previous_year:
        weekday_p2 = analyze_weekday_distribution(meetings_p2)
        fig, desc = create_comparison_bar_chart(
            weekday_p1, weekday_p2, 'Ugedag', 'Antal',
            TEXTS["graph2_title"], TEXTS["graph2_desc"]
        )
    else:
        fig, desc = create_single_bar_chart(
            weekday_p1, 'Ugedag', 'Antal',
            TEXTS["graph2_title"], TEXTS["graph2_desc"]
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Graf 3: Mødetyper
    st.subheader(TEXTS["graph3_title"])
    st.caption(TEXTS["graph3_desc"])
    
    types_p1 = analyze_meeting_types(meetings_p1)
    
    if compare_previous_year:
        types_p2 = analyze_meeting_types(meetings_p2)
        fig, desc = create_comparison_bar_chart(
            types_p1, types_p2, 'Mødetype', 'Antal',
            TEXTS["graph3_title"], TEXTS["graph3_desc"]
        )
    else:
        fig, desc = create_single_bar_chart(
            types_p1, 'Mødetype', 'Antal',
            TEXTS["graph3_title"], TEXTS["graph3_desc"]
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Graf 4: Deltagere
    st.subheader(TEXTS["graph4_title"])
    st.caption(TEXTS["graph4_desc"])
    
    participants_p1 = analyze_participant_distribution(meetings_p1)
    
    if compare_previous_year:
        participants_p2 = analyze_participant_distribution(meetings_p2)
        fig, desc = create_comparison_bar_chart(
            participants_p1, participants_p2, 'Deltagerkategori', 'Antal',
            TEXTS["graph4_title"], TEXTS["graph4_desc"]
        )
    else:
        fig, desc = create_single_bar_chart(
            participants_p1, 'Deltagerkategori', 'Antal',
            TEXTS["graph4_title"], TEXTS["graph4_desc"]
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Graf 5: Medlemstyper
    st.subheader(TEXTS["graph5_title"])
    st.caption(TEXTS["graph5_desc"])
    
    member_types = analyze_member_types(seats_df)
    
    if not member_types.empty:
        fig, desc = create_pie_chart(
            member_types, 'Medlemstype', 'Antal',
            TEXTS["graph5_title"], TEXTS["graph5_desc"]
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Ingen medlemstype-data tilgængelig")
    
    # Graf 6: Grupper med få møder
    st.subheader(TEXTS["graph6_title"])
    st.caption(TEXTS["graph6_desc"])
    
    few_meetings = analyze_groups_with_few_meetings(groups_df, meetings_p1, start_dt, end_dt)
    
    if not few_meetings.empty:
        st.dataframe(few_meetings, use_container_width=True)
    else:
        st.success("Alle grupper har afholdt mindst 4 møder i perioden!")
    
    # Graf 7: Lukkede grupper
    st.subheader(TEXTS["graph7_title"])
    st.caption(TEXTS["graph7_desc"])
    
    closed = analyze_closed_groups(groups_df, start_dt, end_dt)
    
    if not closed.empty:
        st.dataframe(closed, use_container_width=True)
    else:
        st.success("Ingen grupper blev lukket i perioden!")
    
    # Graf 8: Gruppestørrelse
    st.subheader(TEXTS["graph8_title"])
    st.caption(TEXTS["graph8_desc"])
    
    size_dist = analyze_group_size_by_type(groups_df)
    
    if not size_dist.empty:
        # Pivot for grouped bar chart
        pivot = size_dist.pivot(index='Størrelseskategori', columns='Gruppetyper', values='Antal').fillna(0)
        
        fig = go.Figure()
        
        for col in pivot.columns:
            color = get_group_type_color(col)
            fig.add_trace(go.Bar(
                name=col,
                x=pivot.index,
                y=pivot[col],
                marker_color=color
            ))
        
        fig.update_layout(
            title=TEXTS["graph8_title"],
            xaxis_title="Gruppestørrelse",
            yaxis_title="Antal grupper",
            barmode='group',
            height=500
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # ========== PDF DOWNLOAD ==========
    st.markdown("---")
    st.header("📥 Download rapport")
    
    if st.button("Generer PDF-rapport", type="primary"):
        with st.spinner("Genererer PDF..."):
            charts_data = [
                (TEXTS["graph1_title"], TEXTS["graph1_desc"]),
                (TEXTS["graph2_title"], TEXTS["graph2_desc"]),
                (TEXTS["graph3_title"], TEXTS["graph3_desc"]),
                (TEXTS["graph4_title"], TEXTS["graph4_desc"]),
                (TEXTS["graph5_title"], TEXTS["graph5_desc"]),
                (TEXTS["graph6_title"], TEXTS["graph6_desc"]),
                (TEXTS["graph7_title"], TEXTS["graph7_desc"]),
                (TEXTS["graph8_title"], TEXTS["graph8_desc"]),
            ]
            
            period_str = f"{start_date.strftime('%d-%m-%Y')} til {end_date.strftime('%d-%m-%Y')}"
            pdf_buffer = generate_pdf(charts_data, period_str)
            
            st.download_button(
                label="⬇️ Download PDF",
                data=pdf_buffer,
                file_name=f"dge_rapport_{start_date.strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )
            
            st.success("✅ PDF klar til download!")

if __name__ == "__main__":
    main()
