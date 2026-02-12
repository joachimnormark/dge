"""
DGE Mødeaktivitets-analyse
Streamlit app til analyse af DGE-gruppers mødeaktivitet
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image

# ============================================================================
# KONFIGURATION
# ============================================================================

TEXTS = {
    "app_title": "📊 DGE Mødeaktivitets-analyse",
    "upload_header": "Upload datafiler",
    "upload_files": "Upload 3 Excel-filer (grupper, medlemmer, møder)",
    "period_header": "Vælg analyseperiode",
    "compare_checkbox": "Sammenlign med samme periode sidste år",
    
    "table1_title": "Tabel 1: Mødestatistik - Total og fordeling på status",
    "table1_desc": "Viser antal møder fordelt på godkendelsestatus og gruppetype.",
    
    "table2_title": "Tabel 2: Mødedage - Fordeling på ugedage",
    "table2_desc": "Viser hvilke ugedage møder afholdes på, fordelt på gruppetype.",
    
    "table3_title": "Tabel 3: Mødetyper - Fordeling",
    "table3_desc": "Viser antallet af forskellige mødetyper.",
    
    "table4_title": "Tabel 4: Mødedeltagelse - Antal deltagere per møde",
    "table4_desc": "Viser fordelingen af møder efter antal deltagere, fordelt på gruppetype.",
    
    "table5_title": "Tabel 5: Gruppestørrelse fordelt på gruppetype",
    "table5_desc": "Viser fordelingen af gruppestørrelser for DGE, Supervision og Junior grupper.",
    
    "table6_title": "Tabel 6: Gruppernes mødeaktivitet",
    "table6_desc": "Viser hvor mange grupper der har holdt 0, 1, 2, 3... møder i perioden.",
    
    "table7_title": "Tabel 7: Antal grupper medlemmer er medlem af",
    "table7_desc": "Viser hvor mange grupper hvert medlem deltager i (ekskl. Gruppeledere).",
    
    "table8_title": "Tabel 8: Medlemstyper - Fordeling i clusters",
    "table8_desc": "Viser hvordan medlemmer fordeler sig på kategorier.",
    
    "table9_title": "Tabel 9: Supervisors og deres gruppefordeling",
    "table9_desc": "Viser hvilke supervisors der har ansvar for hvilke grupper.",
    
    "table10_title": "Tabel 10: Lukkede grupper i perioden",
    "table10_desc": "Oversigt over grupper der er blevet arkiveret/lukket i den valgte periode.",
}

COLORS = {
    # Gruppetyper
    "DGE": "#2B6CB0",           # kølig blå (autoritet, ro)
    "Supervision": "#C53030",   # dyb rød (kontrast, opmærksomhed)
    "Junior": "#2F855A",        # dæmpet grøn (balance, friskhed)
    
    # Medlemstyper
    "Praktiserende læger": "#1F77B4",      # klassisk blå
    "§-ansatte, vikarer mv": "#FF7F0E",    # varm orange
    "Uddannelseslæger": "#2CA02C",         # klar grøn
    "Ej registreret": "#9467BD",           # lilla
    "Andre": "#8C564B"                     # brunlig
}

# ============================================================================
# HJÆLPEFUNKTIONER - DATA PARSING OG VALIDERING
# ============================================================================

def parse_danish_date(date_str):
    """Parser danske datoformater til datetime objekter"""
    if pd.isna(date_str) or str(date_str).strip() in ['-', '', 'nan', 'NaT']:
        return pd.NaT
    
    if isinstance(date_str, datetime):
        return date_str
    
    date_str = str(date_str).strip().replace('kl.', '').replace('Kl.', '').replace(',', '')
    
    months = {
        'januar': '01', 'februar': '02', 'marts': '03', 'april': '04',
        'maj': '05', 'juni': '06', 'juli': '07', 'august': '08',
        'september': '09', 'oktober': '10', 'november': '11', 'december': '12'
    }
    
    for dk_month, num_month in months.items():
        if dk_month in date_str.lower():
            date_str = date_str.lower().replace(dk_month, num_month)
    
    date_str = date_str.replace('.', ' ').replace('/', ' ')
    
    try:
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except:
        return pd.NaT

def identify_dataframe_type(df):
    """Identificer hvilken type dataframe baseret på kolonner"""
    columns = set(df.columns)
    
    if 'Gruppenavn' in columns and 'Gruppetyper' in columns and 'Antal medlemmer' in columns:
        return 'groups'
    elif 'Medlemskaber' in columns and 'Stillingsbetegnelse' in columns:
        return 'seats'
    elif 'Mødetype' in columns and 'Starttidspunkt' in columns and 'Status' in columns:
        return 'meetings'
    
    return 'unknown'

def standardize_group_type(gtype):
    """Standardiser gruppetype til DGE, Supervision, Junior eller Andre"""
    if pd.isna(gtype):
        return "Andre"
    
    gtype = str(gtype).strip()
    
    if "DGE" in gtype:
        return "DGE"
    elif "Supervision" in gtype:
        return "Supervision"
    elif "Junior" in gtype:
        return "Junior"
    
    return "Andre"

def categorize_member_type(member_type):
    """Kategoriser medlemstyper i definerede clusters"""
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

# ============================================================================
# HJÆLPEFUNKTIONER - DATA FILTRERING
# ============================================================================

def filter_gruppeledere(df, group_name_col='Gruppenavn'):
    """Filtrer Gruppeledere-gruppen fra dataframe"""
    if df is None or df.empty:
        return df
    
    return df[df[group_name_col].astype(str).str.strip().str.lower() != 'gruppeledere'].copy()

def filter_members_gruppeledere(seats_df):
    """Fjern Gruppeledere fra medlemmers gruppemedlemskaber"""
    if seats_df is None or seats_df.empty:
        return seats_df
    
    df = seats_df.copy()
    
    if 'Medlemskaber' in df.columns:
        df['Medlemskaber'] = df['Medlemskaber'].astype(str).apply(
            lambda x: ','.join([
                g.strip() 
                for g in x.split(',') 
                if g.strip() and 'gruppeledere' not in g.lower()
            ])
        )
    
    return df

def get_group_type_from_meeting(meetings_df, group_df):
    """Tilføj gruppetype til møder via merge med groups"""
    if meetings_df is None or group_df is None:
        return meetings_df
    
    df = meetings_df.copy()
    group_types = group_df[['Gruppenavn', 'Gruppetyper']].drop_duplicates()
    df = df.merge(group_types, on='Gruppenavn', how='left')
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    return df

# ============================================================================
# ANALYSE FUNKTIONER
# ============================================================================

def analyze_meeting_status_by_type(meetings_df):
    """Analyser mødestatus fordelt på gruppetype"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    status_map = {
        'Godkendt': 'Godkendt',
        'godkendt': 'Godkendt',
        '-': 'Afventer',
        'Afvist': 'Afvist',
        'afvist': 'Afvist',
        'Afsluttet': 'Afholdt u. godk.',
        'afsluttet': 'Afholdt u. godk.',
    }
    
    df = meetings_df.copy()
    df['Status_mapped'] = df['Status'].map(status_map).fillna('Andet')
    
    result = df.groupby(['Status_mapped', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    result['Status_mapped'] = result['Status_mapped'].astype(str)
    
    return result

def analyze_weekday_by_type(meetings_df):
    """Analyser møder fordelt på ugedag og gruppetype"""
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
    
    df['Ugedag'] = df['Weekday'].map(weekday_map)
    
    result = df.groupby(['Ugedag', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    result['Ugedag'] = result['Ugedag'].astype(str)
    
    return result

def analyze_participants_by_type(meetings_df):
    """Analyser møder fordelt på antal deltagere og gruppetype"""
    if meetings_df is None or meetings_df.empty:
        return pd.DataFrame()
    
    bins = [0, 4, 7, 10, 13, 999]
    labels = ['1-4', '5-7', '8-10', '11-13', '14+']
    
    df = meetings_df.copy()
    df['Deltagerkategori'] = pd.cut(df['Antal deltagere'], bins=bins, labels=labels, right=True)
    
    result = df.groupby(['Deltagerkategori', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    result['Deltagerkategori'] = result['Deltagerkategori'].astype(str)
    
    return result

def analyze_group_size_by_type(groups_df, period_start):
    """Analyser gruppestørrelser fordelt på gruppetype"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    df = groups_df.copy()
    
    # Filtrer grupper arkiveret før periodens start
    if 'Dato for arkivering' in df.columns:
        df = df[df['Dato for arkivering'].isna() | (df['Dato for arkivering'] >= period_start)]
    
    if 'Antal medlemmer' not in df.columns or 'Gruppetyper' not in df.columns:
        return pd.DataFrame()
    
    bins = [0, 4, 6, 8, 10, 12, 14, 999]
    labels = ['1-4', '5-6', '7-8', '9-10', '11-12', '13-14', '15+']
    
    df['Størrelseskategori'] = pd.cut(df['Antal medlemmer'], bins=bins, labels=labels, right=True)
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    result = df.groupby(['Størrelseskategori', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    result['Størrelseskategori'] = result['Størrelseskategori'].astype(str)
    
    return result

def analyze_group_meeting_activity(groups_df, meetings_df):
    """Analyser hvor mange møder hver gruppe har holdt"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    # Tæl møder per gruppe
    if meetings_df is not None and not meetings_df.empty:
        meeting_counts = meetings_df.groupby('Gruppenavn').size().to_dict()
    else:
        meeting_counts = {}
    
    df = groups_df[['Gruppenavn', 'Gruppetyper']].copy()
    df['Antal_møder'] = df['Gruppenavn'].map(meeting_counts).fillna(0).astype(int)
    
    # Kategoriser antal møder
    def categorize_meetings(n):
        return '10+' if n >= 10 else str(int(n))
    
    df['Mødekategori'] = df['Antal_møder'].apply(categorize_meetings)
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    grouped = df.groupby(['Mødekategori', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal_grupper')
    
    # Sikr ALLE kategorier er med (0-10+)
    all_cats = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10+']
    all_types = ['DGE', 'Supervision', 'Junior']
    
    complete = []
    for cat in all_cats:
        for gtype in all_types:
            existing = grouped[(grouped['Mødekategori'] == cat) & (grouped['Gruppetype_std'] == gtype)]
            val = existing['Antal_grupper'].values[0] if len(existing) > 0 else 0
            complete.append({
                'Mødekategori': cat,
                'Gruppetype_std': gtype,
                'Antal_grupper': val
            })
    
    return pd.DataFrame(complete)

def analyze_groups_per_member(seats_df):
    """Analyser hvor mange grupper hvert medlem er med i"""
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    
    df = seats_df.copy()
    
    if 'Medlemskaber' not in df.columns:
        return pd.DataFrame()
    
    # Tæl antal grupper per medlem
    def count_groups(medlemskaber_str):
        if pd.isna(medlemskaber_str):
            return 0
        
        s = str(medlemskaber_str).strip()
        
        if s == '' or s.lower() == 'nan':
            return 0
        
        groups = [g.strip() for g in s.split(',') if g.strip() and g.strip().lower() != 'nan']
        return len(groups)
    
    df['Antal_grupper_int'] = df['Medlemskaber'].apply(count_groups)
    
    # Kategoriser
    def categorize_groups(n):
        n_int = int(n)
        return '4+' if n_int >= 4 else str(n_int)
    
    df['Gruppe_kategori'] = df['Antal_grupper_int'].apply(categorize_groups)
    
    result = df['Gruppe_kategori'].value_counts().reset_index()
    result.columns = ['Antal grupper', 'Antal medlemmer']
    
    # Sorter korrekt
    order = ['0', '1', '2', '3', '4+']
    result['Antal grupper'] = pd.Categorical(result['Antal grupper'], categories=order, ordered=True)
    result = result.sort_values('Antal grupper').reset_index(drop=True)
    result['Antal grupper'] = result['Antal grupper'].astype(str)
    
    return result

def analyze_member_types(seats_df):
    """Analyser medlemstyper"""
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    
    df = seats_df.copy()
    
    if 'Stillingsbetegnelse' not in df.columns:
        return pd.DataFrame()
    
    df['Cluster'] = df['Stillingsbetegnelse'].apply(categorize_member_type)
    
    result = df['Cluster'].value_counts().reset_index()
    result.columns = ['Medlemstype', 'Antal']
    
    return result

def analyze_supervisor_groups(groups_df):
    """NYT: Analyser supervisors og deres gruppefordeling"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    if 'Supervisor' not in groups_df.columns:
        return pd.DataFrame()
    
    df = groups_df.copy()
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    # Group by supervisor og gruppetype
    supervisor_stats = df.groupby(['Supervisor', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    
    # Pivot til at få en række per supervisor
    pivot = supervisor_stats.pivot_table(
        index='Supervisor',
        columns='Gruppetype_std',
        values='Antal',
        fill_value=0,
        aggfunc='sum'
    ).reset_index()
    
    # Sikr alle gruppetyper er kolonner
    for gtype in ['DGE', 'Supervision', 'Junior']:
        if gtype not in pivot.columns:
            pivot[gtype] = 0
    
    # Tilføj total kolonne
    pivot['Total'] = pivot[['DGE', 'Supervision', 'Junior']].sum(axis=1)
    
    # Sorter efter total (højest først)
    pivot = pivot.sort_values('Total', ascending=False)
    
    # Filtrer supervisors uden navn ("-" eller tomme)
    pivot = pivot[pivot['Supervisor'].notna() & (pivot['Supervisor'] != '-')]
    
    return pivot

def analyze_closed_groups(groups_df, start_date, end_date):
    """Analyser lukkede grupper i perioden"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    df = groups_df.copy()
    
    if 'Dato for arkivering' not in df.columns:
        return pd.DataFrame()
    
    mask = (df['Dato for arkivering'] >= start_date) & (df['Dato for arkivering'] <= end_date)
    result = df[mask].copy()
    
    if result.empty:
        return pd.DataFrame()
    
    return result[['Gruppenavn', 'Gruppetyper', 'Dato for arkivering', 'Antal medlemmer']].sort_values('Dato for arkivering')

# ============================================================================
# VISUALISERING
# ============================================================================

def create_stacked_bar_chart(data, x_col, y_col, title, description, ordered_categories=None):
    """Lav stacked bar chart med DGE (bund), Supervision (midt), Junior (top)"""
    if data.empty:
        return None, description
    
    if ordered_categories is None:
        ordered_categories = sorted(data[x_col].unique())
    
    # Byg data direkte for plotly
    traces_data = {'DGE': [], 'Supervision': [], 'Junior': []}
    
    for cat in ordered_categories:
        for gtype in ['DGE', 'Supervision', 'Junior']:
            mask = (data[x_col] == cat) & (data['Gruppetype_std'] == gtype)
            matching = data[mask]
            val = matching[y_col].sum() if len(matching) > 0 else 0
            traces_data[gtype].append(val)
    
    fig = go.Figure()
    
    # Tilføj traces i rækkefølge: DGE (bund), Supervision, Junior (top)
    for gtype in ['DGE', 'Supervision', 'Junior']:
        fig.add_trace(go.Bar(
            name=gtype,
            x=ordered_categories,
            y=traces_data[gtype],
            marker_color=COLORS[gtype]
        ))
    
    fig.update_layout(
        title=title,
        xaxis_title=x_col,
        yaxis_title='Antal',
        barmode='stack',
        height=500,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(type='category')
    )
    
    return fig, description

def create_comparison_stacked_bar(data_p1, data_p2, x_col, y_col, title, description, ordered_categories=None):
    """Lav sammenlignende stacked bar chart for to perioder"""
    if data_p1.empty:
        return None, description
    
    data_p1 = data_p1.copy()
    data_p2 = data_p2.copy()
    data_p1['Periode'] = 'P1'
    data_p2['Periode'] = 'P2'
    
    combined = pd.concat([data_p1, data_p2], ignore_index=True)
    combined[x_col] = combined[x_col].astype(str)
    
    pivot = combined.pivot_table(
        index=[x_col, 'Periode'],
        columns='Gruppetype_std',
        values=y_col,
        aggfunc='sum',
        fill_value=0,
        observed=True
    ).reset_index()
    
    fig = go.Figure()
    
    x_categories = pivot[x_col].unique()
    if ordered_categories:
        x_categories = [c for c in ordered_categories if c in x_categories]
    
    x_labels = []
    for cat in x_categories:
        x_labels.append(f"{cat} P1")
        x_labels.append(f"{cat} P2")
    
    for gtype in ['DGE', 'Supervision', 'Junior']:
        if gtype not in pivot.columns:
            continue
        
        y_values = []
        for cat in x_categories:
            p1_val = pivot[(pivot[x_col] == cat) & (pivot['Periode'] == 'P1')][gtype].values
            p2_val = pivot[(pivot[x_col] == cat) & (pivot['Periode'] == 'P2')][gtype].values
            
            y_values.append(p1_val[0] if len(p1_val) > 0 else 0)
            y_values.append(p2_val[0] if len(p2_val) > 0 else 0)
        
        fig.add_trace(go.Bar(
            name=gtype,
            x=x_labels,
            y=y_values,
            marker_color=COLORS[gtype]
        ))
    
    fig.update_layout(
        title=title,
        xaxis_title=x_col,
        yaxis_title='Antal',
        barmode='stack',
        height=500,
        showlegend=True,
        xaxis=dict(tickangle=-45, type='category'),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    return fig, description

def create_simple_bar_chart(data, x_col, y_col, title, description):
    """Lav simpel bar chart uden gruppetype-opdeling"""
    if data.empty:
        return None, description
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=data[x_col].astype(str),
        y=data[y_col],
        marker_color=COLORS["DGE"]
    ))
    
    fig.update_layout(
        title=title,
        xaxis_title=x_col,
        yaxis_title=y_col,
        height=500,
        xaxis=dict(type='category')
    )
    
    return fig, description

# ============================================================================
# PDF GENERATION
# ============================================================================

def generate_pdf_with_charts(all_charts, period_info):
    """Generér PDF med grafer"""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    
    # Forside
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, height - 100, "DGE Moede aktivitets-rapport")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 130, f"Periode: {period_info}")
    c.drawString(50, height - 150, f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    c.showPage()
    
    # Tilføj grafer
    for i, (fig, title, desc) in enumerate(all_charts):
        if fig is None:
            continue
        
        try:
            # Export figure som PNG
            img_bytes = fig.to_image(format="png", width=1000, height=500, engine='kaleido')
            img = Image.open(io.BytesIO(img_bytes))
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            # Tilføj til PDF
            c.setFont("Helvetica-Bold", 14)
            c.drawString(50, height - 50, title)
            c.setFont("Helvetica", 10)
            
            # Wrap description
            desc_lines = []
            words = desc.split()
            current_line = ""
            for word in words:
                if len(current_line + " " + word) < 100:
                    current_line += " " + word if current_line else word
                else:
                    desc_lines.append(current_line)
                    current_line = word
            if current_line:
                desc_lines.append(current_line)
            
            y_pos = height - 70
            for line in desc_lines[:2]:
                c.drawString(50, y_pos, line)
                y_pos -= 15
            
            # Indsæt billede
            c.drawImage(ImageReader(img_buffer), 50, height - 500, width=700, height=350, preserveAspectRatio=True)
            c.showPage()
            
        except Exception as e:
            # Fallback hvis billede ikke kan eksporteres
            c.setFont("Helvetica-Bold", 14)
            c.drawString(50, height - 100, title)
            c.setFont("Helvetica", 10)
            c.drawString(50, height - 120, desc[:100])
            c.showPage()
    
    c.save()
    buffer.seek(0)
    return buffer

def generate_pdf_details(meetings_df, groups_df, period_str, start_date, end_date):
    """Generér detaljeret PDF med mødelister"""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Forside
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "DGE Detaljeret Moedeoversigt")
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 80, f"Periode: {period_str}")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 100, f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
    
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 140, "Godkendte moeder i perioden:")
    
    # Liste møder
    y_position = height - 170
    c.setFont("Helvetica", 9)
    
    for idx, row in meetings_df.iterrows():
        if y_position < 50:
            c.showPage()
            y_position = height - 50
            c.setFont("Helvetica", 9)
        
        meeting_info = f"{row['Gruppenavn'][:40]} - {row['Starttidspunkt'].strftime('%d-%m-%Y')} - {row['Antal deltagere']} deltagere"
        c.drawString(50, y_position, meeting_info)
        y_position -= 15
    
    # Liste grupper med 0-4 møder
    c.showPage()
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 50, "Grupper med 0-4 moeder:")
    
    meeting_counts = meetings_df.groupby('Gruppenavn').size().to_dict()
    few_meeting_groups = groups_df[
        groups_df['Gruppenavn'].map(lambda x: meeting_counts.get(x, 0) <= 4)
    ].copy()
    few_meeting_groups['Antal_møder'] = few_meeting_groups['Gruppenavn'].map(lambda x: meeting_counts.get(x, 0))
    few_meeting_groups = few_meeting_groups.sort_values('Antal_møder')
    
    y_position = height - 80
    c.setFont("Helvetica", 9)
    
    for idx, row in few_meeting_groups.iterrows():
        if y_position < 50:
            c.showPage()
            y_position = height - 50
            c.setFont("Helvetica", 9)
        
        group_info = f"{row['Gruppenavn'][:50]} - {row['Antal_møder']} møder - {row['Gruppetyper']}"
        c.drawString(50, y_position, group_info)
        y_position -= 15
    
    c.save()
    buffer.seek(0)
    return buffer

# ============================================================================
# STREAMLIT APP
# ============================================================================

st.set_page_config(page_title=TEXTS["app_title"], layout="wide")

def main():
    st.title(TEXTS["app_title"])
    
    # ========== FILE UPLOAD ==========
    st.header(TEXTS["upload_header"])
    
    uploaded_files = st.file_uploader(
        TEXTS["upload_files"],
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if len(uploaded_files) != 3:
        st.info("👆 Upload præcis 3 Excel-filer")
        return
    
    # Identificer og load filer
    groups_df = None
    seats_df = None
    meetings_df = None
    
    try:
        for file in uploaded_files:
            df = pd.read_excel(file)
            df_type = identify_dataframe_type(df)
            
            if df_type == 'groups':
                groups_df = df
            elif df_type == 'seats':
                seats_df = df
            elif df_type == 'meetings':
                meetings_df = df
        
        if groups_df is None or seats_df is None or meetings_df is None:
            st.error("Kunne ikke identificere alle 3 filer. Tjek at de indeholder de rigtige kolonner.")
            return
        
        st.success(f"✅ Data indlæst: {len(groups_df)} grupper, {len(seats_df)} medlemmer, {len(meetings_df)} møder")
    
    except Exception as e:
        st.error(f"Fejl ved indlæsning af filer: {e}")
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
    
    current_year = datetime.now().year
    last_year = current_year - 1
    
    col_year, col_custom = st.columns([1, 2])
    
    with col_year:
        use_full_year = st.checkbox("Brug helt år", value=True)
        if use_full_year:
            selected_year = st.selectbox(
                "Vælg år",
                list(range(2020, current_year + 1)),
                index=list(range(2020, current_year + 1)).index(last_year)
            )
            start_date = datetime(selected_year, 1, 1).date()
            end_date = datetime(selected_year, 12, 31).date()
        else:
            start_date = None
            end_date = None
    
    if not use_full_year:
        with col_custom:
            col_s, col_e = st.columns(2)
            with col_s:
                start_date = st.date_input("Startdato", value=datetime(last_year, 1, 1))
            with col_e:
                end_date = st.date_input("Slutdato", value=datetime(last_year, 12, 31))
    else:
        st.info(f"Valgt periode: {start_date.strftime('%d-%m-%Y')} til {end_date.strftime('%d-%m-%Y')}")
    
    start_dt = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.max.time())
    
    # Filtrer aktive grupper (ikke arkiveret før periodens start)
    groups_df_active = groups_df.copy()
    if 'Dato for arkivering' in groups_df_active.columns:
        groups_df_active = groups_df_active[
            groups_df_active['Dato for arkivering'].isna() | 
            (groups_df_active['Dato for arkivering'] >= start_dt)
        ]
    
    # Sammenligning
    compare_previous_year = st.checkbox(TEXTS["compare_checkbox"])
    
    if compare_previous_year:
        start_dt_p2 = start_dt - relativedelta(years=1)
        end_dt_p2 = end_dt - relativedelta(years=1)
    
    # ========== FILTER MØDER ==========
    
    meetings_p1_all = meetings_df[
        (meetings_df['Starttidspunkt'] >= start_dt) & 
        (meetings_df['Starttidspunkt'] <= end_dt)
    ].copy()
    
    # KUN GODKENDTE MØDER til de fleste analyser
    meetings_p1 = meetings_p1_all[
        meetings_p1_all['Status'].astype(str).str.strip().str.lower() == 'godkendt'
    ].copy()
    
    if compare_previous_year:
        meetings_p2_all = meetings_df[
            (meetings_df['Starttidspunkt'] >= start_dt_p2) & 
            (meetings_df['Starttidspunkt'] <= end_dt_p2)
        ].copy()
        
        meetings_p2 = meetings_p2_all[
            meetings_p2_all['Status'].astype(str).str.strip().str.lower() == 'godkendt'
        ].copy()
    
    if meetings_p1.empty and meetings_p1_all.empty:
        st.warning("Ingen møder fundet i den valgte periode")
        return
    
    # Tilføj gruppetype til møder
    meetings_p1_all = get_group_type_from_meeting(meetings_p1_all, groups_df)
    meetings_p1 = get_group_type_from_meeting(meetings_p1, groups_df)
    
    if compare_previous_year:
        meetings_p2_all = get_group_type_from_meeting(meetings_p2_all, groups_df)
        meetings_p2 = get_group_type_from_meeting(meetings_p2, groups_df)
    
    # ========== METRICS ==========
    st.header("Analyseresultater")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    
    with col_m1:
        st.metric("Antal møder (godkendte, P1)", len(meetings_p1))
        if compare_previous_year:
            st.metric("Antal møder (godkendte, P2)", len(meetings_p2))
    
    with col_m2:
        st.metric("Deltagerdage (godkendte, P1)", meetings_p1['Antal deltagere'].sum())
        if compare_previous_year:
            st.metric("Deltagerdage (godkendte, P2)", meetings_p2['Antal deltagere'].sum())
    
    with col_m3:
        st.metric("Unikke grupper (P1)", meetings_p1['Gruppenavn'].nunique())
        if compare_previous_year:
            st.metric("Unikke grupper (P2)", meetings_p2['Gruppenavn'].nunique())
    
    st.markdown("---")
    
    all_charts = []  # For PDF export
    
    # ========== TABEL 1: MØDESTATISTIK ==========
    st.subheader(TEXTS["table1_title"])
    st.caption(TEXTS["table1_desc"])
    
    status_p1 = analyze_meeting_status_by_type(meetings_p1_all)
    
    if compare_previous_year:
        status_p2 = analyze_meeting_status_by_type(meetings_p2_all)
        fig, desc = create_comparison_stacked_bar(
            status_p1, status_p2, 'Status_mapped', 'Antal',
            TEXTS["table1_title"], TEXTS["table1_desc"],
            ordered_categories=['Godkendt', 'Afventer', 'Afholdt u. godk.', 'Afvist', 'Andet']
        )
    else:
        fig, desc = create_stacked_bar_chart(
            status_p1, 'Status_mapped', 'Antal',
            TEXTS["table1_title"], TEXTS["table1_desc"],
            ordered_categories=['Godkendt', 'Afventer', 'Afholdt u. godk.', 'Afvist', 'Andet']
        )
    
    if fig:
        fig.update_layout(xaxis_title='Mødestatus', yaxis_title='Antal møder')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table1_title"], TEXTS["table1_desc"]))
    
    # ========== TABEL 2: UGEDAGE ==========
    st.subheader(TEXTS["table2_title"])
    st.caption(TEXTS["table2_desc"])
    
    weekday_p1 = analyze_weekday_by_type(meetings_p1)
    weekday_order = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag', 'Søndag']
    
    if compare_previous_year:
        weekday_p2 = analyze_weekday_by_type(meetings_p2)
        fig, desc = create_comparison_stacked_bar(
            weekday_p1, weekday_p2, 'Ugedag', 'Antal',
            TEXTS["table2_title"], TEXTS["table2_desc"],
            ordered_categories=weekday_order
        )
    else:
        fig, desc = create_stacked_bar_chart(
            weekday_p1, 'Ugedag', 'Antal',
            TEXTS["table2_title"], TEXTS["table2_desc"],
            ordered_categories=weekday_order
        )
    
    if fig:
        fig.update_layout(xaxis_title='Mødedag på ugen', yaxis_title='Antal møder')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table2_title"], TEXTS["table2_desc"]))
    
    # ========== TABEL 3: MØDETYPER ==========
    st.subheader(TEXTS["table3_title"])
    st.caption(TEXTS["table3_desc"])
    
    types_p1 = meetings_p1['Mødetype'].value_counts().reset_index()
    types_p1.columns = ['Mødetype', 'Antal']
    
    fig, desc = create_simple_bar_chart(
        types_p1, 'Mødetype', 'Antal',
        TEXTS["table3_title"], TEXTS["table3_desc"]
    )
    
    if fig:
        fig.update_layout(xaxis_title='Mødetype', yaxis_title='Antal møder')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table3_title"], TEXTS["table3_desc"]))
    
    # ========== TABEL 4: MØDEDELTAGELSE ==========
    st.subheader(TEXTS["table4_title"])
    st.caption(TEXTS["table4_desc"])
    
    participants_p1 = analyze_participants_by_type(meetings_p1)
    participant_order = ['1-4', '5-7', '8-10', '11-13', '14+']
    
    if compare_previous_year:
        participants_p2 = analyze_participants_by_type(meetings_p2)
        fig, desc = create_comparison_stacked_bar(
            participants_p1, participants_p2, 'Deltagerkategori', 'Antal',
            TEXTS["table4_title"], TEXTS["table4_desc"],
            ordered_categories=participant_order
        )
    else:
        fig, desc = create_stacked_bar_chart(
            participants_p1, 'Deltagerkategori', 'Antal',
            TEXTS["table4_title"], TEXTS["table4_desc"],
            ordered_categories=participant_order
        )
    
    if fig:
        fig.update_layout(xaxis_title='Antal mødedeltagere', yaxis_title='Antal møder')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table4_title"], TEXTS["table4_desc"]))
    
    # ========== TABEL 5: GRUPPESTØRRELSE ==========
    st.subheader(TEXTS["table5_title"])
    st.caption(TEXTS["table5_desc"])
    
    size_dist = analyze_group_size_by_type(groups_df_active, start_dt)
    size_order = ['1-4', '5-6', '7-8', '9-10', '11-12', '13-14', '15+']
    
    fig, desc = create_stacked_bar_chart(
        size_dist, 'Størrelseskategori', 'Antal',
        TEXTS["table5_title"], TEXTS["table5_desc"],
        ordered_categories=size_order
    )
    
    if fig:
        fig.update_layout(xaxis_title='Antal medlemmer i gruppen', yaxis_title='Antal grupper')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table5_title"], TEXTS["table5_desc"]))
    
    # ========== TABEL 6: GRUPPERNES MØDEAKTIVITET ==========
    st.subheader(TEXTS["table6_title"])
    st.caption(TEXTS["table6_desc"])
    
    activity = analyze_group_meeting_activity(groups_df_active, meetings_p1)
    activity_order = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10+']
    
    fig, desc = create_stacked_bar_chart(
        activity, 'Mødekategori', 'Antal_grupper',
        TEXTS["table6_title"], TEXTS["table6_desc"],
        ordered_categories=activity_order
    )
    
    if fig:
        fig.update_layout(xaxis_title='Antal møder i perioden', yaxis_title='Antal grupper')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table6_title"], TEXTS["table6_desc"]))
    
    # ========== TABEL 7: ANTAL GRUPPER PER MEDLEM ==========
    st.subheader(TEXTS["table7_title"])
    st.caption(TEXTS["table7_desc"])
    
    groups_per_member = analyze_groups_per_member(seats_df)
    
    fig, desc = create_simple_bar_chart(
        groups_per_member, 'Antal grupper', 'Antal medlemmer',
        TEXTS["table7_title"], TEXTS["table7_desc"]
    )
    
    if fig:
        fig.update_layout(xaxis_title='Antal medlemsskaber (Gruppeledergruppe undtaget)', yaxis_title='Antal medlemmer')
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table7_title"], TEXTS["table7_desc"]))
    
    # ========== TABEL 8: MEDLEMSTYPER ==========
    st.subheader(TEXTS["table8_title"])
    st.caption(TEXTS["table8_desc"])
    
    member_types = analyze_member_types(seats_df)
    
    if not member_types.empty:
        # Trim og hent unikke navne
        member_types['Medlemstype'] = member_types['Medlemstype'].astype(str).str.strip()
        unique_clusters = member_types['Medlemstype'].unique().tolist()
        
        # Lav color mapping
        fallback_colors = ['#FFD700', '#00CED1', '#FF8C00', '#A9A9A9']
        color_discrete_map = {}
        colors_list = []
        fb_idx = 0
        
        for cluster in unique_clusters:
            if cluster in COLORS and COLORS[cluster]:
                color = COLORS[cluster]
            else:
                color = fallback_colors[fb_idx % len(fallback_colors)]
                fb_idx += 1
            color_discrete_map[cluster] = color
            colors_list.append(color)
        
        # Opret pie chart
        fig = px.pie(
            member_types,
            names='Medlemstype',
            values='Antal',
            title=TEXTS["table8_title"],
            hole=0.3,
            color_discrete_map=color_discrete_map
        )
        
        fig.update_traces(
            textinfo='percent+label',
            marker=dict(colors=colors_list, line=dict(color='#FFFFFF', width=1))
        )
        fig.update_layout(height=500)
        
        st.plotly_chart(fig, use_container_width=True)
        all_charts.append((fig, TEXTS["table8_title"], TEXTS["table8_desc"]))
    
    # ========== NY TABEL 9: SUPERVISORS OG DERES GRUPPEFORDELING ==========
    st.subheader(TEXTS["table9_title"])
    st.caption(TEXTS["table9_desc"])
    
    supervisor_stats = analyze_supervisor_groups(groups_df_active)
    
    if not supervisor_stats.empty:
        # Vis som tabel med formattering
        st.dataframe(
            supervisor_stats[['Supervisor', 'DGE', 'Supervision', 'Junior', 'Total']],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Ingen supervisor-data tilgængelig")
    
    # BEMÆRK: Denne tabel inkluderes IKKE i PDF (kun web)
    
    # ========== TABEL 10: LUKKEDE GRUPPER ==========
    st.subheader(TEXTS["table10_title"])
    st.caption(TEXTS["table10_desc"])
    
    closed = analyze_closed_groups(groups_df, start_dt, end_dt)
    
    if not closed.empty:
        st.dataframe(closed, use_container_width=True)
    else:
        st.success("Ingen grupper blev lukket i perioden!")
    
    # ========== PDF DOWNLOAD ==========
    st.markdown("---")
    st.header("📥 Download rapport")
    
    col_pdf1, col_pdf2 = st.columns(2)
    
    with col_pdf1:
        if st.button("Hent PDF med grafer", type="primary"):
            with st.spinner("Genererer PDF med grafer..."):
                period_str = f"{start_date.strftime('%d-%m-%Y')} til {end_date.strftime('%d-%m-%Y')}"
                
                try:
                    pdf_buffer = generate_pdf_with_charts(all_charts, period_str)
                    
                    st.download_button(
                        label="⬇️ Download PDF med grafer",
                        data=pdf_buffer,
                        file_name=f"dge_rapport_{start_date.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
                    
                    st.success("✅ PDF klar!")
                except Exception as e:
                    st.error(f"PDF-fejl: {e}")
    
    with col_pdf2:
        if st.button("Hent PDF med detaljer"):
            with st.spinner("Genererer detaljeret mødeliste..."):
                period_str = f"{start_date.strftime('%d-%m-%Y')} til {end_date.strftime('%d-%m-%Y')}"
                
                try:
                    pdf_buf = generate_pdf_details(
                        meetings_p1,
                        groups_df_active,
                        period_str,
                        start_date=start_dt,
                        end_date=end_dt
                    )
                    
                    st.download_button(
                        label="⬇️ Download PDF med detaljer",
                        data=pdf_buf,
                        file_name=f"DGE_detaljer_{start_date.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
                    
                    st.success("✅ PDF klar!")
                except Exception as e:
                    st.error(f"Fejl ved generering af detaljer: {e}")

if __name__ == "__main__":
    main()
