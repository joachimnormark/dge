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
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.units import inch
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
    
    "table9_title": "Tabel 9: Vejledere og deres gruppefordeling",
    "table9_desc": "Viser hvilke vejledere der har ansvar for hvilke grupper.",
    
    "table10_title": "Tabel 10: Lukkede grupper i perioden",
    "table10_desc": "Oversigt over grupper der er blevet arkiveret/lukket i den valgte periode.",
}

COLORS = {
    # Gruppetyper
    "DGE": "#2B6CB0",
    "Supervision": "#C53030",
    "Junior": "#2F855A",
    
    # Medlemstyper
    "Praktiserende læger": "#1F77B4",
    "§-ansatte, vikarer mv": "#FF7F0E",
    "Uddannelseslæger": "#2CA02C",
    "Ej registreret": "#9467BD",
    "Andre": "#8C564B"
}

# ============================================================================
# HJÆLPEFUNKTIONER - DATA PARSING
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
    """Identificer hvilken type dataframe"""
    columns = set(df.columns)
    
    if 'Gruppenavn' in columns and 'Gruppetyper' in columns and 'Antal medlemmer' in columns:
        return 'groups'
    elif 'Medlemskaber' in columns and 'Stillingsbetegnelse' in columns:
        return 'seats'
    elif 'Mødetype' in columns and 'Starttidspunkt' in columns and 'Status' in columns:
        return 'meetings'
    
    return 'unknown'

def standardize_group_type(gtype):
    """Standardiser gruppetype"""
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
    """Kategoriser medlemstyper"""
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
    """Filtrer Gruppeledere-gruppen"""
    if df is None or df.empty:
        return df
    
    return df[df[group_name_col].astype(str).str.strip().str.lower() != 'gruppeledere'].copy()

def filter_members_gruppeledere(seats_df):
    """Fjern Gruppeledere fra medlemskaber"""
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
    """Tilføj gruppetype til møder"""
    if meetings_df is None or group_df is None:
        return meetings_df
    
    df = meetings_df.copy()
    
    # Merge gruppetype og supervisor
    group_info = group_df[['Gruppenavn', 'Gruppetyper', 'Supervisor']].drop_duplicates()
    df = df.merge(group_info, on='Gruppenavn', how='left')
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    # Filtrer Gruppeledere EFTER merge
    df = df[df['Gruppenavn'].astype(str).str.strip().str.lower() != 'gruppeledere'].copy()
    
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
    """Analyser møder fordelt på ugedag"""
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
    """Analyser møder fordelt på antal deltagere"""
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
    """Analyser gruppestørrelser"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    df = groups_df.copy()
    
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
    """Analyser gruppernes mødeaktivitet"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    # Ekstra sikkerhedscheck: Filtrer Gruppeledere
    groups_df = filter_gruppeledere(groups_df.copy())
    
    if meetings_df is not None and not meetings_df.empty:
        # Filtrer også møder
        meetings_df = filter_gruppeledere(meetings_df.copy())
        meeting_counts = meetings_df.groupby('Gruppenavn').size().to_dict()
    else:
        meeting_counts = {}
    
    df = groups_df[['Gruppenavn', 'Gruppetyper']].copy()
    df['Antal_møder'] = df['Gruppenavn'].map(meeting_counts).fillna(0).astype(int)
    
    def categorize_meetings(n):
        return '10+' if n >= 10 else str(int(n))
    
    df['Mødekategori'] = df['Antal_møder'].apply(categorize_meetings)
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    grouped = df.groupby(['Mødekategori', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal_grupper')
    
    # Sikr ALLE kategorier er med
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
    """Analyser antal grupper per medlem"""
    if seats_df is None or seats_df.empty:
        return pd.DataFrame()
    
    df = seats_df.copy()
    
    if 'Medlemskaber' not in df.columns:
        return pd.DataFrame()
    
    def count_groups(medlemskaber_str):
        if pd.isna(medlemskaber_str):
            return 0
        
        s = str(medlemskaber_str).strip()
        
        if s == '' or s.lower() == 'nan':
            return 0
        
        # Filtrer både 'nan' og 'gruppeledere'
        groups = [
            g.strip() 
            for g in s.split(',') 
            if g.strip() 
            and g.strip().lower() != 'nan'
            and 'gruppeledere' not in g.strip().lower()
        ]
        return len(groups)
    
    df['Antal_grupper_int'] = df['Medlemskaber'].apply(count_groups)
    
    def categorize_groups(n):
        n_int = int(n)
        return '4+' if n_int >= 4 else str(n_int)
    
    df['Gruppe_kategori'] = df['Antal_grupper_int'].apply(categorize_groups)
    
    result = df['Gruppe_kategori'].value_counts().reset_index()
    result.columns = ['Antal grupper', 'Antal medlemmer']
    
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
    """Analyser supervisors og deres gruppefordeling"""
    if groups_df is None or groups_df.empty:
        return pd.DataFrame()
    
    if 'Supervisor' not in groups_df.columns:
        return pd.DataFrame()
    
    df = groups_df.copy()
    df['Gruppetype_std'] = df['Gruppetyper'].apply(standardize_group_type)
    
    supervisor_stats = df.groupby(['Supervisor', 'Gruppetype_std'], observed=True).size().reset_index(name='Antal')
    
    pivot = supervisor_stats.pivot_table(
        index='Supervisor',
        columns='Gruppetype_std',
        values='Antal',
        fill_value=0,
        aggfunc='sum'
    ).reset_index()
    
    for gtype in ['DGE', 'Supervision', 'Junior']:
        if gtype not in pivot.columns:
            pivot[gtype] = 0
    
    pivot['Total'] = pivot[['DGE', 'Supervision', 'Junior']].sum(axis=1)
    pivot = pivot.sort_values('Total', ascending=False)
    pivot = pivot[pivot['Supervisor'].notna() & (pivot['Supervisor'] != '-')]
    
    return pivot

def analyze_closed_groups(groups_df, start_date, end_date):
    """Analyser lukkede grupper"""
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
    """Lav stacked bar chart"""
    if data.empty:
        return None, description
    
    if ordered_categories is None:
        ordered_categories = sorted(data[x_col].unique())
    
    traces_data = {'DGE': [], 'Supervision': [], 'Junior': []}
    
    for cat in ordered_categories:
        for gtype in ['DGE', 'Supervision', 'Junior']:
            mask = (data[x_col] == cat) & (data['Gruppetype_std'] == gtype)
            matching = data[mask]
            val = matching[y_col].sum() if len(matching) > 0 else 0
            traces_data[gtype].append(val)
    
    fig = go.Figure()
    
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
    """Lav sammenlignende stacked bar chart"""
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
    """Lav simpel bar chart"""
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
# PDF GENERATION - GRAFER
# ============================================================================

def generate_pdf_with_charts(all_charts, period_info):
    """Generér PDF med grafer - professionel styling"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=landscape(A4),
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=22,
        textColor=colors.HexColor('#2B6CB0'),
        spaceAfter=20,
        alignment=1,  # Center
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#4A5568'),
        spaceAfter=8,
        alignment=1
    )
    
    chart_title_style = ParagraphStyle(
        'ChartTitle',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#2B6CB0'),
        spaceAfter=8,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )
    
    # Forside
    story.append(Paragraph("DGE Mødeaktivitets-rapport", title_style))
    story.append(Paragraph(f"Periode: {period_info}", subtitle_style))
    story.append(Paragraph(f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_style))
    story.append(Spacer(1, 0.5*inch))
    story.append(PageBreak())
    
    # Tilføj grafer
    for i, (fig, title, desc) in enumerate(all_charts):
        if fig is None:
            continue
        
        try:
            # Export graf som PNG
            img_bytes = fig.to_image(format="png", width=1000, height=500, engine='kaleido')
            img = Image.open(io.BytesIO(img_bytes))
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            # Titel
            story.append(Paragraph(title, chart_title_style))
            
            # Beskrivelse
            story.append(Paragraph(desc, styles['Normal']))
            story.append(Spacer(1, 0.15*inch))
            
            # Billede
            from reportlab.platypus import Image as RLImage
            img_obj = RLImage(img_buffer, width=9*inch, height=5*inch)
            story.append(img_obj)
            
            story.append(PageBreak())
            
        except Exception as e:
            # Fallback
            story.append(Paragraph(title, chart_title_style))
            story.append(Paragraph(f"Kunne ikke generere graf: {str(e)}", styles['Normal']))
            story.append(PageBreak())
    
    # Byg PDF
    doc.build(story)
    buffer.seek(0)
    return buffer

# ============================================================================
# PDF GENERATION - DETALJER (NY VERSION)
# ============================================================================

def generate_pdf_details(meetings_df, groups_df, period_str, start_date, end_date):
    """Generér detaljeret PDF med mødelister og gruppestatistik"""
    
    # SIKKERHEDSCHECK: Filtrer Gruppeledere
    meetings_df = filter_gruppeledere(meetings_df.copy())
    groups_df = filter_gruppeledere(groups_df.copy())
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#2B6CB0'),
        spaceAfter=20,
        alignment=1  # Center
    )
    
    section_style = ParagraphStyle(
        'SectionTitle',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#2B6CB0'),
        spaceAfter=12,
        spaceBefore=12
    )
    
    # Style for status column - matches table typography
    status_style = ParagraphStyle(
        'StatusStyle',
        parent=styles['Normal'],
        fontSize=8,
        fontName='Helvetica',
        leading=10,
        alignment=0  # Left align
    )
    
    # Forside
    story.append(Paragraph("DGE Detaljeret Mødeoversigt", title_style))
    story.append(Paragraph(f"Periode: {period_str}", styles['Normal']))
    story.append(Paragraph(f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    story.append(PageBreak())
    
    # ========== DEL 1: MØDER GRUPPERET EFTER ANTAL DELTAGERE ==========
    
    participant_ranges = [
        ('1-4', 1, 4),
        ('5-7', 5, 7),
        ('8-10', 8, 10),
        ('11-13', 11, 13),
        ('14+', 14, 999)
    ]
    
    for range_name, min_part, max_part in participant_ranges:
        # Filtrer møder i denne deltagergruppe
        range_meetings = meetings_df[
            (meetings_df['Antal deltagere'] >= min_part) & 
            (meetings_df['Antal deltagere'] <= max_part)
        ].copy()
        
        if range_meetings.empty:
            continue
        
        # Sektionshoved
        story.append(Paragraph(f"Møder med {range_name} deltagere", section_style))
        story.append(Spacer(1, 0.1*inch))
        
        # Gruppér efter gruppetype
        for gtype in ['DGE', 'Supervision', 'Junior']:
            gtype_meetings = range_meetings[range_meetings['Gruppetype_std'] == gtype].copy()
            
            if gtype_meetings.empty:
                continue
            
            # Gruppetype overskrift
            story.append(Paragraph(f"<b>{gtype}</b>", styles['Heading3']))
            
            # Sorter efter dato
            gtype_meetings = gtype_meetings.sort_values('Starttidspunkt')
            
            # Byg tabel data
            table_data = [['Dato', 'Gruppenavn', 'Vejleder', 'Mødetype', 'Antal']]
            
            for _, row in gtype_meetings.iterrows():
                date_str = row['Starttidspunkt'].strftime('%d-%m-%Y')
                group_name = str(row['Gruppenavn'])[:35]
                vejleder = str(row.get('Supervisor', '-'))[:25]
                meeting_type = str(row['Mødetype'])[:20]
                participants = str(int(row['Antal deltagere']))
                
                table_data.append([date_str, group_name, vejleder, meeting_type, participants])
            
            # Opret tabel
            table = Table(table_data, colWidths=[1*inch, 2.3*inch, 1.5*inch, 1.3*inch, 0.9*inch])
            
            table.setStyle(TableStyle([
                # Header styling
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('TOPPADDING', (0, 0), (-1, 0), 8),
                
                # Data rows
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('ALIGN', (4, 1), (4, -1), 'CENTER'),  # Antal kolonne centered
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
                
                # Grid
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2B6CB0')),
                
                # Padding
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 1), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
            ]))
            
            story.append(table)
            story.append(Spacer(1, 0.2*inch))
        
        story.append(PageBreak())
    
    # ========== DEL 2: GRUPPER GRUPPERET EFTER ANTAL MØDER ==========
    
    # Tæl møder per gruppe
    meeting_counts = meetings_df.groupby('Gruppenavn').size().to_dict()
    
    # Tilføj til groups dataframe
    groups_with_counts = groups_df.copy()
    
    # KRITISK: Filtrer grupper der var lukket FØR perioden startede
    # Disse grupper skal IKKE fremgå i rapporten
    if 'Dato for arkivering' in groups_with_counts.columns:
        groups_with_counts = groups_with_counts[
            groups_with_counts['Dato for arkivering'].isna() |  # Ikke arkiveret
            (groups_with_counts['Dato for arkivering'] >= start_date)  # Eller arkiveret i/efter perioden
        ].copy()
    
    groups_with_counts['Antal_møder'] = groups_with_counts['Gruppenavn'].map(meeting_counts).fillna(0).astype(int)
    groups_with_counts['Gruppetype_std'] = groups_with_counts['Gruppetyper'].apply(standardize_group_type)
    
    # Find maks antal møder for at vide hvor langt vi skal gå
    max_meetings = groups_with_counts['Antal_møder'].max() if not groups_with_counts.empty else 0
    
    # Lav kategorier: 0, 1, 2, ..., op til max_meetings, derefter 10+ hvis der er grupper med 10+
    meeting_categories = list(range(0, min(max_meetings + 1, 10)))  # 0-9
    if max_meetings >= 10:
        meeting_categories.append('10+')
    
    for num_meetings in meeting_categories:
        if num_meetings == '10+':
            # Grupper med 10 eller flere møder
            groups_this_count = groups_with_counts[groups_with_counts['Antal_møder'] >= 10].copy()
        else:
            # Grupper med præcis num_meetings møder
            groups_this_count = groups_with_counts[groups_with_counts['Antal_møder'] == num_meetings].copy()
        
        if groups_this_count.empty:
            continue
        
        # Sektionshoved
        story.append(Paragraph(f"Grupper med {num_meetings} møder i perioden", section_style))
        story.append(Spacer(1, 0.1*inch))
        
        # IKKE opdelt efter gruppetype - vis alle på én gang
        # Sorter efter gruppetype så de kommer i rækkefølge
        groups_this_count = groups_this_count.sort_values(['Gruppetype_std', 'Gruppenavn'])
        
        # Byg tabel data
        table_data = [['Gruppenavn', 'Vejleder', 'Status', 'Gruppetype', 'Antal møder']]
        
        for _, row in groups_this_count.iterrows():
            group_name = str(row['Gruppenavn'])[:40]
            vejleder = str(row.get('Supervisor', '-'))[:25]
            
            # Bestem lukket status - SAMME TYPOGRAFI SOM RESTEN
            arkiv_dato = row.get('Dato for arkivering')
            if pd.notna(arkiv_dato):
                if start_date <= arkiv_dato <= end_date:
                    # Lukket i perioden - RØD (små bogstaver, ingen bold)
                    status_text = f"<font color='red'>Lukket {arkiv_dato.strftime('%d-%m-%Y')}</font>"
                else:
                    # Lukket efter perioden - GUL/ORANGE (små bogstaver, ingen bold)
                    status_text = f"<font color='#DAA520'>Lukket efter perioden {arkiv_dato.strftime('%d-%m-%Y')}</font>"
            else:
                status_text = "Aktiv"
            
            gruppetype = str(row['Gruppetyper'])[:15]
            antal = str(int(row['Antal_møder']))
            
            table_data.append([
                group_name, 
                vejleder, 
                Paragraph(status_text, status_style), 
                gruppetype, 
                antal
            ])
        
        # Opret tabel
        table = Table(table_data, colWidths=[2.2*inch, 1.4*inch, 1.8*inch, 1*inch, 1*inch])
        
        table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (4, 1), (4, -1), 'CENTER'),  # Antal kolonne centered
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
            
            # Grid
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2B6CB0')),
            
            # Padding
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.3*inch))
        
        # Hvis der er mange grupper, tilføj PageBreak efter hver kategori
        if len(table_data) > 20:
            story.append(PageBreak())
    
    # ========== DEL 3: DGE GRUPPER MED SUPERVISION MØDER ==========
    
    # Find DGE grupper (gruppetype) der har holdt møder af typen "Supervision" (mødetype)
    if not meetings_df.empty:
        # Find møder der er af typen "Supervision"
        supervision_meetings = meetings_df[
            meetings_df['Mødetype'].astype(str).str.contains('Supervision', case=False, na=False)
        ].copy()
        
        if not supervision_meetings.empty:
            # Tjek om vi allerede har Gruppetype_std (fra get_group_type_from_meeting)
            if 'Gruppetype_std' not in supervision_meetings.columns:
                # Merge med gruppetype hvis ikke allerede der
                group_types = groups_df[['Gruppenavn', 'Gruppetyper', 'Supervisor']].drop_duplicates()
                supervision_meetings = supervision_meetings.merge(
                    group_types, 
                    on='Gruppenavn', 
                    how='left'
                )
                supervision_meetings['Gruppetype_std'] = supervision_meetings['Gruppetyper'].apply(standardize_group_type)
            
            # Hvis ikke supervisor kolonne findes, hent den
            if 'Supervisor' not in supervision_meetings.columns:
                supervisor_info = groups_df[['Gruppenavn', 'Supervisor']].drop_duplicates()
                supervision_meetings = supervision_meetings.merge(
                    supervisor_info,
                    on='Gruppenavn',
                    how='left'
                )
            
            # Filtrer kun DGE grupper
            dge_with_supervision = supervision_meetings[
                supervision_meetings['Gruppetype_std'] == 'DGE'
            ].copy()
            
            if not dge_with_supervision.empty:
                # Tæl antal supervision møder per gruppe
                supervision_counts = dge_with_supervision.groupby(['Gruppenavn', 'Supervisor']).size().reset_index(name='Antal_supervision_møder')
                
                # Sorter efter antal møder (højest først)
                supervision_counts = supervision_counts.sort_values('Antal_supervision_møder', ascending=False)
                
                # Tilføj sektion
                story.append(PageBreak())
                story.append(Paragraph("DGE-grupper der har holdt Supervision-møder", section_style))
                story.append(Spacer(1, 0.1*inch))
                story.append(Paragraph(
                    "Følgende DGE-grupper har afholdt møder af typen 'Supervision' i perioden:", 
                    styles['Normal']
                ))
                story.append(Spacer(1, 0.15*inch))
                
                # Byg tabel
                table_data = [['Gruppenavn', 'Vejleder', 'Antal Supervision-møder']]
                
                for _, row in supervision_counts.iterrows():
                    group_name = str(row['Gruppenavn'])[:50]
                    vejleder = str(row.get('Supervisor', '-'))[:30]
                    antal = str(int(row['Antal_supervision_møder']))
                    
                    table_data.append([group_name, vejleder, antal])
                
                # Opret tabel
                table = Table(table_data, colWidths=[3.5*inch, 2*inch, 2*inch])
                
                table.setStyle(TableStyle([
                    # Header styling
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('TOPPADDING', (0, 0), (-1, 0), 8),
                    
                    # Data rows
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Antal kolonne centered
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
                    
                    # Grid
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2B6CB0')),
                    
                    # Padding
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 1), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
                ]))
                
                story.append(table)
                story.append(Spacer(1, 0.3*inch))
    
    # Byg PDF
    doc.build(story)
    buffer.seek(0)
    return buffer

# ============================================================================
# PDF GENERATION - VEJLEDER RAPPORT (NY)
# ============================================================================

def generate_pdf_supervisor_report(supervisor_name, groups_df, meetings_df, period_str, start_date, end_date):
    """Generér rapport for specifik vejleder"""
    
    # Filtrer Gruppeledere
    groups_df = filter_gruppeledere(groups_df.copy())
    meetings_df = filter_gruppeledere(meetings_df.copy())
    
    # Filtrer kun denne vejleders grupper
    supervisor_groups = groups_df[
        groups_df['Supervisor'].astype(str).str.strip() == supervisor_name
    ].copy()
    
    # Filtrer kun aktive grupper (ikke arkiveret)
    supervisor_groups = supervisor_groups[
        supervisor_groups['Dato for arkivering'].isna()
    ].copy()
    
    if supervisor_groups.empty:
        return None
    
    # Standardiser gruppetype for sortering
    supervisor_groups['Gruppetype_std'] = supervisor_groups['Gruppetyper'].apply(standardize_group_type)
    
    # Tæl møder først så vi kan sortere efter dem
    for idx, group_row in supervisor_groups.iterrows():
        group_meetings = meetings_df[meetings_df['Gruppenavn'] == group_row['Gruppenavn']]
        supervisor_groups.loc[idx, 'Antal_møder_temp'] = len(group_meetings)
    
    # Sorter efter gruppetype, derefter antal møder (stigende)
    supervisor_groups = supervisor_groups.sort_values(['Gruppetype_std', 'Antal_møder_temp'])
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#2B6CB0'),
        spaceAfter=20,
        alignment=1
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#4A5568'),
        spaceAfter=8,
        alignment=1
    )
    
    # Forside
    story.append(Paragraph(f"Grupperapport for vejleder", title_style))
    story.append(Paragraph(f"<b>{supervisor_name}</b>", subtitle_style))
    story.append(Paragraph(f"Periode: {period_str}", subtitle_style))
    story.append(Paragraph(f"Genereret: {datetime.now().strftime('%d-%m-%Y %H:%M')}", subtitle_style))
    story.append(Spacer(1, 0.5*inch))
    
    # Byg tabel data
    table_data = [[
        'Gruppenavn',
        'Gruppetype', 
        'Antal\nmøder',
        'Antal\nmedlemmer',
        'Deltager-\nprocent',
        'Møde-\nomkostninger',
        'Verificeret'
    ]]
    
    for _, group_row in supervisor_groups.iterrows():
        group_name = str(group_row['Gruppenavn'])
        group_type = str(group_row['Gruppetyper'])[:15]
        num_members = int(group_row['Antal medlemmer']) if pd.notna(group_row['Antal medlemmer']) else 0
        
        # Tæl møder for denne gruppe
        group_meetings = meetings_df[meetings_df['Gruppenavn'] == group_row['Gruppenavn']]
        num_meetings = len(group_meetings)
        
        # Beregn deltagerprocent
        if not group_meetings.empty and num_members > 0:
            avg_participants = group_meetings['Antal deltagere'].mean()
            participation_pct = (avg_participants / num_members) * 100
            participation_str = f"{participation_pct:.0f}%"
        else:
            participation_str = "-"
        
        # Beregn total mødeomkostninger
        if not group_meetings.empty and 'Mødeomkostninger' in group_meetings.columns:
            # Parse mødeomkostninger (format: "1.234,56 kr." eller "0,00 kr.")
            def parse_cost(cost_str):
                if pd.isna(cost_str):
                    return 0
                cost_str = str(cost_str).replace('kr.', '').replace('.', '').replace(',', '.').strip()
                try:
                    return float(cost_str)
                except:
                    return 0
            
            total_cost = group_meetings['Mødeomkostninger'].apply(parse_cost).sum()
            cost_str = f"{total_cost:,.0f} kr.".replace(',', '.')
        else:
            cost_str = "0 kr."
        
        # Parse verificeringsdato
        verif_date = group_row.get('Dato for verificering')
        if pd.notna(verif_date) and str(verif_date).strip() not in ['-', '', 'nan']:
            verif_dt = parse_danish_date(verif_date)
            if pd.notna(verif_dt):
                verif_str = verif_dt.strftime('%d-%m-%Y')
            else:
                verif_str = "-"
        else:
            verif_str = "-"
        
        table_data.append([
            group_name[:30],
            group_type,
            str(num_meetings),
            str(num_members),
            participation_str,
            cost_str,
            verif_str
        ])
    
    # Opret tabel
    table = Table(table_data, colWidths=[
        2*inch,      # Gruppenavn
        0.8*inch,    # Gruppetype
        0.6*inch,    # Antal møder
        0.6*inch,    # Antal medlemmer
        0.7*inch,    # Deltagerprocent
        0.9*inch,    # Mødeomkostninger
        0.8*inch     # Verificeret
    ])
    
    table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (2, 1), (6, -1), 'CENTER'),  # Tal-kolonner centered
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
        ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2B6CB0')),
        
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
    ]))
    
    story.append(table)
    
    # Byg PDF
    doc.build(story)
    buffer.seek(0)
    return buffer

# ============================================================================
# STREAMLIT APP (Fortsætter som før...)
# ============================================================================

st.set_page_config(page_title=TEXTS["app_title"], layout="wide")

# VERSION NUMMER
APP_VERSION = "v9.0 - Filtrerer grupper lukket før perioden - 2026-02-14"

def main():
    st.title(TEXTS["app_title"])
    st.caption(f"Version: {APP_VERSION}")
    
    # FILE UPLOAD
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
            st.error("Kunne ikke identificere alle 3 filer")
            return
        
        st.success(f"✅ Data indlæst: {len(groups_df)} grupper, {len(seats_df)} medlemmer, {len(meetings_df)} møder")
    
    except Exception as e:
        st.error(f"Fejl: {e}")
        return
    
    # ========== TJEK FOR DUBLETTER I GROUPS ==========
    duplicate_groups = groups_df[groups_df.duplicated(subset=['Gruppenavn'], keep=False)]
    
    if not duplicate_groups.empty:
        duplicate_names = duplicate_groups['Gruppenavn'].unique().tolist()
        
        st.error("🚨 **FEJL: Der findes grupper med samme navn!**")
        st.write("Følgende gruppenavne forekommer mere end én gang i datafilen:")
        
        for name in duplicate_names:
            dupes = groups_df[groups_df['Gruppenavn'] == name]
            st.write(f"\n**'{name}'** (findes {len(dupes)} gange):")
            st.dataframe(dupes[['Gruppenavn', 'Status', 'Gruppetyper', 'Supervisor', 'Dato for arkivering']])
        
        st.warning("⚠️ Ret venligst dine data så hver gruppe kun forekommer én gang, og genindlæs filerne.")
        st.info("💡 Tip: Hvis en gruppe er både aktiv og inaktiv, behold kun én af rækkerne.")
        st.stop()  # Stop programmet her
    
    # DATA CLEANING
    if 'Starttidspunkt' in meetings_df.columns:
        meetings_df['Starttidspunkt'] = meetings_df['Starttidspunkt'].apply(parse_danish_date)
    
    if 'Dato for arkivering' in groups_df.columns:
        groups_df['Dato for arkivering'] = groups_df['Dato for arkivering'].apply(parse_danish_date)
    
    groups_df = filter_gruppeledere(groups_df)
    meetings_df = filter_gruppeledere(meetings_df)
    seats_df = filter_members_gruppeledere(seats_df)
    
    # PERIODE VALG
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
    
    # Filtrer aktive grupper
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
    
    # FILTER MØDER
    meetings_p1_all = meetings_df[
        (meetings_df['Starttidspunkt'] >= start_dt) & 
        (meetings_df['Starttidspunkt'] <= end_dt)
    ].copy()
    
    meetings_p1 = meetings_p1_all[
        meetings_p1_all['Status'].astype(str).str.strip().str.lower() == 'godkendt'
    ].copy()
    
    # Ekstra sikkerhedscheck: Fjern eventuelle Gruppeledere-møder
    meetings_p1 = filter_gruppeledere(meetings_p1)
    
    if compare_previous_year:
        meetings_p2_all = meetings_df[
            (meetings_df['Starttidspunkt'] >= start_dt_p2) & 
            (meetings_df['Starttidspunkt'] <= end_dt_p2)
        ].copy()
        
        meetings_p2 = meetings_p2_all[
            meetings_p2_all['Status'].astype(str).str.strip().str.lower() == 'godkendt'
        ].copy()
        
        # Ekstra sikkerhedscheck: Fjern eventuelle Gruppeledere-møder
        meetings_p2 = filter_gruppeledere(meetings_p2)
    
    if meetings_p1.empty and meetings_p1_all.empty:
        st.warning("Ingen møder fundet")
        return
    
    # Tilføj gruppetype og supervisor til møder
    meetings_p1_all = get_group_type_from_meeting(meetings_p1_all, groups_df)
    meetings_p1 = get_group_type_from_meeting(meetings_p1, groups_df)
    
    if compare_previous_year:
        meetings_p2_all = get_group_type_from_meeting(meetings_p2_all, groups_df)
        meetings_p2 = get_group_type_from_meeting(meetings_p2, groups_df)
    
    # METRICS
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
    
    # DEBUG INFORMATION
    with st.expander("🔍 DEBUG: Tjek at tallene er korrekte"):
        st.write("**Forventede korrekte tal for 2025:**")
        st.write("- Møder: 219 (uden Gruppeledere)")
        st.write("- Deltagerdage: 1551")
        
        st.write(f"\n**Faktiske tal:**")
        st.write(f"- Møder: {len(meetings_p1)}")
        st.write(f"- Deltagerdage: {meetings_p1['Antal deltagere'].sum()}")
        
        if len(meetings_p1) == 219 and meetings_p1['Antal deltagere'].sum() == 1551:
            st.success("✅ Tallene er nu korrekte!")
        else:
            st.warning("⚠️ Tallene matcher ikke endnu")
    
    st.markdown("---")
    
    all_charts = []
    
    # TABEL 1: MØDESTATISTIK
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
    
    # TABEL 2: UGEDAGE
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
    
    # TABEL 3: MØDETYPER
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
    
    # TABEL 4: MØDEDELTAGELSE
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
    
    # TABEL 5: GRUPPESTØRRELSE
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
    
    # TABEL 6: GRUPPERNES MØDEAKTIVITET
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
    
    # TABEL 7: ANTAL GRUPPER PER MEDLEM
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
    
    # TABEL 8: MEDLEMSTYPER
    st.subheader(TEXTS["table8_title"])
    st.caption(TEXTS["table8_desc"])
    
    member_types = analyze_member_types(seats_df)
    
    if not member_types.empty:
        member_types['Medlemstype'] = member_types['Medlemstype'].astype(str).str.strip()
        unique_clusters = member_types['Medlemstype'].unique().tolist()
        
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
    
    # NY TABEL 9: SUPERVISORS
    st.subheader(TEXTS["table9_title"])
    st.caption(TEXTS["table9_desc"])
    
    supervisor_stats = analyze_supervisor_groups(groups_df_active)
    
    if not supervisor_stats.empty:
        # Omdøb kolonne til visning
        display_stats = supervisor_stats.copy()
        display_stats = display_stats.rename(columns={'Supervisor': 'Vejleder'})
        
        st.dataframe(
            display_stats[['Vejleder', 'DGE', 'Supervision', 'Junior', 'Total']],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Ingen supervisor-data tilgængelig")
    
    # TABEL 10: LUKKEDE GRUPPER
    st.subheader(TEXTS["table10_title"])
    st.caption(TEXTS["table10_desc"])
    
    closed = analyze_closed_groups(groups_df, start_dt, end_dt)
    
    if not closed.empty:
        st.dataframe(closed, use_container_width=True)
    else:
        st.success("Ingen grupper blev lukket!")
    
    # PDF DOWNLOAD
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
                        groups_df,
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
                    st.error(f"Fejl: {e}")
    
    # NY: PDF TIL VEJLEDER
    st.markdown("---")
    st.subheader("📋 PDF til vejleder")
    
    # Hent liste over vejledere med grupper
    available_supervisors = groups_df_active[
        groups_df_active['Supervisor'].notna() & 
        (groups_df_active['Supervisor'] != '-')
    ]['Supervisor'].unique().tolist()
    
    available_supervisors = sorted([s for s in available_supervisors if str(s).strip()])
    
    if available_supervisors:
        col_sup1, col_sup2 = st.columns([1, 1])
        
        with col_sup1:
            selected_supervisor = st.selectbox(
                "Vælg vejleder",
                options=available_supervisors,
                key="supervisor_select"
            )
        
        with col_sup2:
            st.write("")  # Spacing
            st.write("")  # Spacing
            if st.button("Generér rapport for vejleder", type="secondary"):
                with st.spinner(f"Genererer rapport for {selected_supervisor}..."):
                    period_str = f"{start_date.strftime('%d-%m-%Y')} til {end_date.strftime('%d-%m-%Y')}"
                    
                    try:
                        pdf_buf = generate_pdf_supervisor_report(
                            selected_supervisor,
                            groups_df,
                            meetings_p1,
                            period_str,
                            start_date=start_dt,
                            end_date=end_dt
                        )
                        
                        if pdf_buf is None:
                            st.warning(f"Ingen aktive grupper fundet for {selected_supervisor}")
                        else:
                            # Lav filnavn-venligt vejledernavn
                            safe_name = selected_supervisor.replace(' ', '_').replace('/', '_')
                            
                            st.download_button(
                                label=f"⬇️ Download rapport for {selected_supervisor}",
                                data=pdf_buf,
                                file_name=f"Vejleder_{safe_name}_{start_date.strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                key="download_supervisor"
                            )
                            
                            st.success("✅ Vejleder-rapport klar!")
                    except Exception as e:
                        st.error(f"Fejl ved generering: {e}")
    else:
        st.info("Ingen vejledere med grupper fundet i perioden")

if __name__ == "__main__":
    main()
