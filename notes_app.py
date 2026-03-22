import streamlit as st
import pandas as pd
import math
import io
from datetime import datetime

# Vérification des modules optionnels
try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

st.set_page_config(
    page_title="Suivi des Notes",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700;800;900&family=JetBrains+Mono:wght@400;600&display=swap');

:root {
    --blue:    #1565C0;
    --blue2:   #0D47A1;
    --sky:     #E3F2FD;
    --gold:    #F9A825;
    --green:   #2E7D32;
    --red:     #C62828;
    --orange:  #E65100;
    --bg:      #F0F4F8;
    --card:    #FFFFFF;
    --text:    #0D1B2A;
    --muted:   #546E7A;
    --border:  #CFD8DC;
}

* { font-family: 'Outfit', sans-serif; color: var(--text); }

.stApp {
    background: var(--bg);
    background-image: radial-gradient(circle at 15% 15%, rgba(21,101,192,0.06) 0%, transparent 60%),
                      radial-gradient(circle at 85% 85%, rgba(249,168,37,0.05) 0%, transparent 60%);
}

/* ── Header ── */
.app-header {
    background: linear-gradient(135deg, var(--blue2) 0%, var(--blue) 60%, #1976D2 100%);
    border-radius: 20px;
    padding: 30px 36px 26px;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 8px 32px rgba(13,71,161,0.28);
}
.app-header::before {
    content: '🎓';
    position: absolute; right: 32px; top: 50%;
    transform: translateY(-50%);
    font-size: 90px; opacity: 0.10; line-height: 1;
}
.app-header::after {
    content: '';
    position: absolute; bottom: 0; left: 0; right: 0; height: 4px;
    background: linear-gradient(90deg, var(--gold), transparent);
}
.app-tag   { font-size:12px; font-weight:700; letter-spacing:4px; color:rgba(255,255,255,0.55); text-transform:uppercase; margin-bottom:6px; }
.app-title { font-size:32px; font-weight:900; color:#fff; line-height:1.1; margin:0; }
.app-sub   { font-size:14px; color:rgba(255,255,255,0.6); margin-top:6px; }

/* ── Cards ── */
.card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 18px;
    box-shadow: 0 2px 12px rgba(13,71,161,0.06);
}
.card-title {
    font-size: 16px; font-weight: 800; color: var(--blue2);
    border-bottom: 2px solid var(--sky);
    padding-bottom: 10px; margin-bottom: 18px;
    letter-spacing: 0.3px;
}

/* ── Stat boxes ── */
.stat-grid { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:18px; }
.stat-box {
    background: var(--sky); border-radius:12px; padding:14px 20px;
    border-left: 4px solid var(--blue); flex:1; min-width:120px;
}
.stat-box.gold  { background:#FFF8E1; border-color:var(--gold); }
.stat-box.green { background:#E8F5E9; border-color:var(--green); }
.stat-box.red   { background:#FFEBEE; border-color:var(--red); }
.stat-label { font-size:11px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:var(--muted); margin-bottom:4px; }
.stat-value { font-size:26px; font-weight:900; color:var(--blue2); line-height:1; }
.stat-box.gold  .stat-value { color:#7B5800; }
.stat-box.green .stat-value { color:var(--green); }
.stat-box.red   .stat-value { color:var(--red); }

/* ── Mention badge ── */
.mention {
    display:inline-block; padding:3px 12px; border-radius:20px;
    font-size:12px; font-weight:700; letter-spacing:0.5px;
}
.mention-excellent { background:#E8F5E9; color:#1B5E20; border:1px solid #A5D6A7; }
.mention-bien      { background:#E3F2FD; color:#0D47A1; border:1px solid #90CAF9; }
.mention-ab        { background:#FFF8E1; color:#7B5800; border:1px solid #FFE082; }
.mention-passable  { background:#FFF3E0; color:#E65100; border:1px solid #FFCC80; }
.mention-insuffisant{ background:#FFEBEE; color:#B71C1C; border:1px solid #EF9A9A; }

/* ── Progress bar ── */
.prog-bar-wrap { background:#E0E7EF; border-radius:999px; height:8px; margin:4px 0; overflow:hidden; }
.prog-bar { height:8px; border-radius:999px; }
.prog-up   { background:linear-gradient(90deg,#43A047,#66BB6A); }
.prog-down { background:linear-gradient(90deg,#EF5350,#E57373); }
.prog-flat { background:linear-gradient(90deg,#FFA726,#FFB74D); }

/* ── Table ── */
.stDataFrame { border-radius:12px !important; overflow:hidden; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--sky); border-radius:12px; padding:4px; gap:4px; border:none; margin-bottom:18px;
}
.stTabs [data-baseweb="tab"] {
    border-radius:8px; font-weight:700; font-size:13px; color:var(--muted);
    padding:9px 18px; border:none; background:transparent;
}
.stTabs [aria-selected="true"] { background:var(--blue) !important; color:white !important; }
.stTabs [data-baseweb="tab-border"] { display:none; }

/* ── Buttons ── */
.stButton > button {
    background:linear-gradient(135deg,var(--blue),var(--blue2)) !important;
    color:white !important; border:none !important; border-radius:10px !important;
    font-weight:700 !important; font-size:14px !important; padding:11px 28px !important;
    box-shadow:0 4px 14px rgba(13,71,161,0.25) !important; width:100%;
}
.stDownloadButton > button {
    background:white !important; border-radius:10px !important;
    font-weight:700 !important; font-size:13px !important;
    padding:9px 20px !important; width:100%; margin-top:6px;
}
.dl-excel .stDownloadButton > button { color:var(--green) !important; border:2px solid var(--green) !important; }
.dl-pdf   .stDownloadButton > button { color:var(--red) !important;   border:2px solid var(--red)   !important; }

/* ── Inputs ── */
.stSelectbox label, .stFileUploader label { font-weight:700 !important; color:var(--text) !important; }

/* ── Metrics ── */
[data-testid="metric-container"] {
    background:var(--sky); border-radius:10px; padding:12px 16px !important; border:1px solid var(--border);
}
[data-testid="stMetricLabel"] p  { color:#333 !important; font-weight:700 !important; font-size:12px !important; }
[data-testid="stMetricValue"] div{ color:var(--blue2) !important; font-weight:900 !important; font-size:22px !important; }

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
    background:var(--sky) !important; border:2px dashed var(--blue) !important;
    border-radius:12px !important; padding:12px !important;
}

/* ── Footer ── */
.app-footer { text-align:center; padding:18px; color:var(--muted); font-size:12px; margin-top:24px; border-top:1px solid var(--border); }
.badge { display:inline-block; background:var(--sky); color:var(--blue2); font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; border:1px solid var(--border); }

#MainMenu, footer, header { visibility:hidden; }
.block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_mention(moyenne):
    if moyenne >= 16:   return "Excellent",   "mention-excellent"
    elif moyenne >= 14: return "Bien",         "mention-bien"
    elif moyenne >= 12: return "Assez Bien",   "mention-ab"
    elif moyenne >= 10: return "Passable",     "mention-passable"
    else:               return "Insuffisant",  "mention-insuffisant"

def get_trend(d1, d2, d3):
    """Retourne tendance globale entre D1 et D3"""
    diff = d3 - d1
    if diff > 1:    return "📈 En progrès",   "green"
    elif diff < -1: return "📉 En baisse",     "red"
    else:           return "➡️ Stable",        "orange"

def progress_bar(val, max_val=20, trend="flat"):
    pct = min(100, (val / max_val) * 100)
    cls = "prog-up" if trend=="green" else ("prog-down" if trend=="red" else "prog-flat")
    return f'<div class="prog-bar-wrap"><div class="prog-bar {cls}" style="width:{pct}%"></div></div>'

def make_excel_bulletin(df_result):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bulletin de Notes"

    blue_fill  = PatternFill("solid", fgColor="1565C0")
    blue2_fill = PatternFill("solid", fgColor="0D47A1")
    sky_fill   = PatternFill("solid", fgColor="E3F2FD")
    gold_fill  = PatternFill("solid", fgColor="FFF8E1")
    green_fill = PatternFill("solid", fgColor="E8F5E9")
    red_fill   = PatternFill("solid", fgColor="FFEBEE")
    thin = Side(style="thin", color="CFD8DC")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def wc(r, c, v, fill=None, font=None, align="center"):
        x = ws.cell(row=r, column=c, value=v)
        if fill: x.fill = fill
        if font: x.font = font
        x.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        x.border = bdr
        return x

    # Titre
    ws.merge_cells("A1:H1")
    wc(1,1, "🎓 BULLETIN DE NOTES — SUIVI DE PROGRESSION",
       fill=blue_fill, font=Font(bold=True, color="FFFFFF", size=14))
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:H2")
    wc(2,1, f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
       font=Font(italic=True, color="546E7A", size=10))
    ws.row_dimensions[2].height = 18

    # En-têtes
    headers = ["N°", "Nom Étudiant", "Devoir 1", "Devoir 2", "Devoir 3", "Moyenne", "Mention", "Tendance"]
    for c, h in enumerate(headers, 1):
        wc(3, c, h, fill=blue2_fill, font=Font(bold=True, color="FFFFFF", size=11))
    ws.row_dimensions[3].height = 22

    # Données
    for i, row in enumerate(df_result.itertuples(), 1):
        r = i + 3
        mention, _ = get_mention(row.Moyenne)
        trend_txt, trend_color = get_trend(row.Devoir1, row.Devoir2, row.Devoir3)

        # Couleur ligne selon mention
        if row.Moyenne >= 14:   row_fill = green_fill
        elif row.Moyenne >= 10: row_fill = sky_fill
        else:                   row_fill = red_fill

        wc(r, 1, i,             fill=row_fill, font=Font(bold=True))
        wc(r, 2, row.Etudiant,  fill=row_fill, font=Font(bold=True), align="left")
        wc(r, 3, row.Devoir1,   fill=row_fill)
        wc(r, 4, row.Devoir2,   fill=row_fill)
        wc(r, 5, row.Devoir3,   fill=row_fill)
        wc(r, 6, round(row.Moyenne, 2), fill=row_fill, font=Font(bold=True))
        wc(r, 7, mention,       fill=row_fill, font=Font(bold=True))
        wc(r, 8, trend_txt,     fill=row_fill)
        ws.row_dimensions[r].height = 18

    # Ligne stats
    last_r = len(df_result) + 4
    ws.merge_cells(f"A{last_r}:B{last_r}")
    wc(last_r, 1, "STATISTIQUES CLASSE", fill=gold_fill, font=Font(bold=True, color="7B5800"))
    wc(last_r, 3, f"Min: {df_result['Devoir1'].min()}", fill=gold_fill)
    wc(last_r, 4, f"Min: {df_result['Devoir2'].min()}", fill=gold_fill)
    wc(last_r, 5, f"Min: {df_result['Devoir3'].min()}", fill=gold_fill)
    wc(last_r, 6, round(df_result['Moyenne'].mean(), 2), fill=gold_fill, font=Font(bold=True, color="7B5800"))

    # Largeurs
    for col, width in zip("ABCDEFGH", [6, 28, 12, 12, 12, 12, 16, 18]):
        ws.column_dimensions[col].width = width

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

def make_pdf_bulletin(df_result):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    BLUE  = colors.HexColor("#1565C0")
    BLUE2 = colors.HexColor("#0D47A1")
    SKY   = colors.HexColor("#E3F2FD")
    GOLD  = colors.HexColor("#FFF8E1")
    GREEN = colors.HexColor("#E8F5E9")
    RED   = colors.HexColor("#FFEBEE")

    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t", fontSize=15, fontName="Helvetica-Bold", textColor=colors.white)
    sub_s   = ParagraphStyle("s", fontSize=9,  fontName="Helvetica",      textColor=colors.HexColor("#B0BEC5"))
    stat_s  = ParagraphStyle("st",fontSize=10, fontName="Helvetica-Bold", textColor=BLUE2, spaceAfter=4)

    story = []

    # Header
    ht = Table([[
        Paragraph("🎓 Bulletin de Notes — Suivi de Progression", title_s),
        Paragraph(f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}", sub_s)
    ]], colWidths=[18*cm, 9*cm])
    ht.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1), BLUE),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),16),
        ("TOPPADDING",(0,0),(-1,-1),14),
        ("BOTTOMPADDING",(0,0),(-1,-1),14),
    ]))
    story.append(ht)
    story.append(Spacer(1, 0.4*cm))

    # Stats rapides
    moy_classe = df_result['Moyenne'].mean()
    nb_reussi  = (df_result['Moyenne'] >= 10).sum()
    nb_echec   = (df_result['Moyenne'] < 10).sum()
    story.append(Paragraph(
        f"<b>Classe :</b> {len(df_result)} étudiants &nbsp;|&nbsp; "
        f"<b>Moyenne classe :</b> {moy_classe:.2f}/20 &nbsp;|&nbsp; "
        f"<b>Réussite :</b> {nb_reussi} ({nb_reussi/len(df_result)*100:.0f}%) &nbsp;|&nbsp; "
        f"<b>Échec :</b> {nb_echec}", stat_s))
    story.append(Spacer(1, 0.3*cm))

    # Tableau principal
    headers = ["N°", "Nom Étudiant", "Devoir 1\n/20", "Devoir 2\n/20", "Devoir 3\n/20",
               "Moyenne\n/20", "Mention", "Tendance", "Progression\nD1→D3"]
    tdata = [headers]

    for i, row in enumerate(df_result.sort_values("Moyenne", ascending=False).itertuples(), 1):
        mention, _ = get_mention(row.Moyenne)
        trend_txt, _ = get_trend(row.Devoir1, row.Devoir2, row.Devoir3)
        prog = row.Devoir3 - row.Devoir1
        prog_str = f"+{prog:.1f}" if prog > 0 else f"{prog:.1f}"
        tdata.append([
            str(i), row.Etudiant,
            str(row.Devoir1), str(row.Devoir2), str(row.Devoir3),
            f"{row.Moyenne:.2f}", mention, trend_txt, prog_str
        ])

    # Ligne moyenne classe
    tdata.append([
        "—", "MOYENNE CLASSE", "—", "—", "—",
        f"{moy_classe:.2f}", "—", "—", "—"
    ])

    col_w = [1*cm, 6*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.5*cm, 3*cm, 3*cm, 2.5*cm]
    t = Table(tdata, colWidths=col_w, repeatRows=1)

    ts_list = [
        ("BACKGROUND",    (0,0),  (-1,0),  BLUE2),
        ("TEXTCOLOR",     (0,0),  (-1,0),  colors.white),
        ("FONTNAME",      (0,0),  (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),  (-1,-1), 8),
        ("ALIGN",         (0,0),  (-1,-1), "CENTER"),
        ("ALIGN",         (0,1),  (1,-1),  "LEFT"),
        ("FONTNAME",      (0,1),  (1,-1),  "Helvetica-Bold"),
        ("GRID",          (0,0),  (-1,-1), 0.4, colors.HexColor("#CFD8DC")),
        ("TOPPADDING",    (0,0),  (-1,-1), 5),
        ("BOTTOMPADDING", (0,0),  (-1,-1), 5),
        ("LEFTPADDING",   (0,0),  (-1,-1), 6),
        # Dernière ligne = stats
        ("BACKGROUND",    (0,-1), (-1,-1), GOLD),
        ("FONTNAME",      (0,-1), (-1,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0,-1), (-1,-1), colors.HexColor("#7B5800")),
    ]

    # Colorier lignes selon moyenne
    for i, row in enumerate(df_result.sort_values("Moyenne", ascending=False).itertuples(), 1):
        if row.Moyenne >= 14:   fill = GREEN
        elif row.Moyenne >= 10: fill = SKY
        else:                   fill = RED
        ts_list.append(("BACKGROUND", (0,i), (-1,i), fill))

    t.setStyle(TableStyle(ts_list))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(
        "<font color='#546E7A' size='7'>Application Suivi des Notes v1.0 — Généré automatiquement</font>",
        styles["Normal"]))

    doc.build(story)
    buf.seek(0)
    return buf

# ── Parsers fichiers ─────────────────────────────────────────────────────────
def _parse_excel(uploaded_file):
    """Lit un fichier Excel et détecte automatiquement le format (standard ou MASSAR)."""
    import openpyxl
    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)

    # Chercher feuille NotesCC (format MASSAR) ou première feuille
    sheet_name = "NotesCC" if "NotesCC" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    rows = list(ws.iter_rows(values_only=True))

    # ── Détection format MASSAR (colonnes en arabe) ──
    # Chercher ligne contenant "الفرض الأول" (Fard 1)
    header_row_idx = None
    for i, row in enumerate(rows):
        row_str = " ".join([str(v) for v in row if v])
        if "الفرض الأول" in row_str or "إسم التلميذ" in row_str:
            header_row_idx = i
            break

    if header_row_idx is not None:
        # Format MASSAR : col 3=Nom, col 6=D1, col 8=D2, col 10=D3, col 12=Activités
        data = []
        for row in rows[header_row_idx + 2:]:  # skip header + sub-header
            if row[3] and isinstance(row[3], str) and len(str(row[3]).strip()) > 1:
                nom = str(row[3]).strip()
                d1  = row[6]  if len(row) > 6  else None
                d2  = row[8]  if len(row) > 8  else None
                d3  = row[10] if len(row) > 10 else None
                act = row[12] if len(row) > 12 else None
                if any(v is not None for v in [d1, d2, d3]):
                    data.append({"Etudiant": nom, "Devoir1": d1, "Devoir2": d2,
                                 "Devoir3": d3, "Activites": act})
        return pd.DataFrame(data)

    # ── Format standard (colonnes en français/anglais) ──
    # Chercher ligne d'en-tête
    for i, row in enumerate(rows[:20]):
        row_lower = [str(v).lower() if v else "" for v in row]
        if any(x in " ".join(row_lower) for x in ["etudiant","nom","élève","student","name"]):
            df = pd.DataFrame(rows[i+1:], columns=rows[i])
            df.columns = [str(c).strip() if c else f"col_{j}" for j,c in enumerate(df.columns)]
            col_map = {}
            for c in df.columns:
                cl = str(c).lower()
                if any(x in cl for x in ["nom","etudiant","élève","eleve","student","name"]): col_map[c] = "Etudiant"
                elif any(x in cl for x in ["d1","devoir1","devoir 1","note1","ds1","fard1","fard 1"]): col_map[c] = "Devoir1"
                elif any(x in cl for x in ["d2","devoir2","devoir 2","note2","ds2","fard2","fard 2"]): col_map[c] = "Devoir2"
                elif any(x in cl for x in ["d3","devoir3","devoir 3","note3","ds3","fard3","fard 3"]): col_map[c] = "Devoir3"
            return df.rename(columns=col_map)

    # ── Fallback : colonnes positionnelles ──
    df = pd.DataFrame(rows[1:])
    if df.shape[1] >= 4:
        df = df.iloc[:, :4]
        df.columns = ["Etudiant","Devoir1","Devoir2","Devoir3"]
    return df

def _parse_standard(df_raw):
    """Parse CSV standard."""
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
        if any(x in row_str for x in ["etudiant","nom","eleve","student"]):
            df = df_raw.iloc[i+1:].copy()
            df.columns = range(len(df.columns))
            return pd.DataFrame({
                "Etudiant": df.iloc[:,0], "Devoir1": df.iloc[:,1],
                "Devoir2":  df.iloc[:,2], "Devoir3": df.iloc[:,3]
            })
    # Fallback positionnelle
    df = df_raw.copy()
    df.columns = ["Etudiant","Devoir1","Devoir2","Devoir3"] + list(range(max(0, len(df.columns)-4)))
    return df

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-tag">Outil Pédagogique</div>
    <div class="app-title">🎓 Suivi des Notes Étudiants</div>
    <div class="app-sub">Analyse de progression — Import Excel / CSV</div>
</div>
""", unsafe_allow_html=True)

# ── IMPORT FICHIER ─────────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">📂 Import du fichier de notes</div>', unsafe_allow_html=True)

col_up, col_fmt = st.columns([2, 1])
with col_up:
    uploaded = st.file_uploader("Choisir un fichier Excel ou CSV",
                                type=["xlsx", "xls", "csv"],
                                help="Le fichier doit contenir : Nom étudiant, Note D1, Note D2, Note D3")
with col_fmt:
    st.markdown("""
    <div style="background:#E3F2FD;border-radius:10px;padding:14px;border-left:4px solid #1565C0;margin-top:28px">
    <b style="color:#0D47A1">Format attendu :</b><br>
    <code style="font-size:12px;color:#1565C0">
    Etudiant | Devoir1 | Devoir2 | Devoir3
    </code><br>
    <span style="font-size:12px;color:#546E7A">Notes sur 20</span>
    </div>
    """, unsafe_allow_html=True)

# Fichier exemple téléchargeable
sample_data = pd.DataFrame({
    "Etudiant": ["Ali Bennani", "Fatima Zahra", "Youssef Alami", "Sara Idrissi", "Omar Chakir",
                 "Nadia Fassi", "Karim Berrada", "Leila Ouali", "Hamid Tazi", "Amina Sefrioui"],
    "Devoir1":  [12.5, 15.0, 8.0, 11.0, 14.5, 9.5, 16.0, 13.0, 7.5, 18.0],
    "Devoir2":  [13.0, 14.5, 9.5, 12.5, 13.0, 11.0, 15.5, 14.0, 9.0, 17.5],
    "Devoir3":  [14.5, 16.0, 11.0, 11.5, 12.0, 13.5, 17.0, 15.5, 10.5, 19.0],
})
# Exemple en CSV (pas besoin d'openpyxl)
sample_csv = sample_data.to_csv(index=False).encode("utf-8")
st.download_button("📥 Télécharger fichier exemple (CSV)", data=sample_csv,
                   file_name="exemple_notes.csv",
                   mime="text/csv")

st.markdown('</div>', unsafe_allow_html=True)

# ── TRAITEMENT ─────────────────────────────────────────────────────────────────
if uploaded:
    try:
        if uploaded.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded, header=None)
            # Format CSV standard
            df_raw.columns = range(len(df_raw.columns))
            # Chercher ligne d'en-tête
            df = _parse_standard(df_raw)
        else:
            df = _parse_excel(uploaded)

        required = ["Etudiant","Devoir1","Devoir2","Devoir3"]
        missing  = [c for c in required if c not in df.columns]
        if missing:
            st.error(f"❌ Colonnes introuvables. Vérifiez le format du fichier.")
            st.info("Format attendu : Etudiant | Devoir1 | Devoir2 | Devoir3")
        else:
            # Ajouter Activites si présente dans le fichier
            if "Activites" not in df.columns:
                df["Activites"] = None

            cols_base = ["Etudiant","Devoir1","Devoir2","Devoir3","Activites"]
            df = df[[c for c in cols_base if c in df.columns]].copy()
            df = df.dropna(subset=["Etudiant"])
            df["Devoir1"]   = pd.to_numeric(df["Devoir1"],   errors="coerce")
            df["Devoir2"]   = pd.to_numeric(df["Devoir2"],   errors="coerce")
            df["Devoir3"]   = pd.to_numeric(df["Devoir3"],   errors="coerce")
            df["Activites"] = pd.to_numeric(df["Activites"], errors="coerce")
            df = df.dropna(subset=["Devoir1","Devoir2","Devoir3"])
            df["Etudiant"] = df["Etudiant"].astype(str).str.strip()
            df = df[df["Etudiant"].str.len() > 1]
            # Moyenne : 4 notes si Activites présente, sinon 3 notes
            has_act = df["Activites"].notna().any()
            if has_act:
                df["Moyenne"] = df.apply(
                    lambda r: (r["Devoir1"] + r["Devoir2"] + r["Devoir3"] + r["Activites"]) / 4
                    if pd.notna(r["Activites"]) else
                    (r["Devoir1"] + r["Devoir2"] + r["Devoir3"]) / 3, axis=1)
            else:
                df["Moyenne"] = (df["Devoir1"] + df["Devoir2"] + df["Devoir3"]) / 3
            # Progression = D3 - D1 (uniquement sur les 3 devoirs)
            df["Progression"] = df["Devoir3"] - df["Devoir1"]

            # ── STATS GLOBALES ──────────────────────────────────────────────
            moy_classe  = df["Moyenne"].mean()
            nb_reussi   = (df["Moyenne"] >= 10).sum()
            nb_echec    = (df["Moyenne"] < 10).sum()
            meilleur    = df.loc[df["Moyenne"].idxmax(), "Etudiant"]
            plus_progres= df.loc[df["Progression"].idxmax(), "Etudiant"]

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">📊 Statistiques Générales</div>', unsafe_allow_html=True)

            st.markdown(f"""
            <div class="stat-grid">
                <div class="stat-box">
                    <div class="stat-label">Étudiants</div>
                    <div class="stat-value">{len(df)}</div>
                </div>
                <div class="stat-box gold">
                    <div class="stat-label">Moyenne classe</div>
                    <div class="stat-value">{moy_classe:.2f}/20</div>
                </div>
                <div class="stat-box green">
                    <div class="stat-label">Réussite ≥10</div>
                    <div class="stat-value">{nb_reussi} ({nb_reussi/len(df)*100:.0f}%)</div>
                </div>
                <div class="stat-box red">
                    <div class="stat-label">Échec &lt;10</div>
                    <div class="stat-value">{nb_echec} ({nb_echec/len(df)*100:.0f}%)</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            c1.metric("🏆 Meilleur étudiant", meilleur, f"{df['Moyenne'].max():.2f}/20")
            c2.metric("📈 Plus grande progression", plus_progres,
                      f"+{df['Progression'].max():.2f} pts")
            st.markdown('</div>', unsafe_allow_html=True)

            # ── TABS ───────────────────────────────────────────────────────
            tab1, tab2, tab3, tab4 = st.tabs([
                "📋 Tableau de bord", "📈 Graphiques", "🏆 Classement", "📄 Export"
            ])

            # ── TAB 1 : Tableau de bord ─────────────────────────────────────
            with tab1:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📋 Détail par étudiant</div>', unsafe_allow_html=True)

                for _, row in df.sort_values("Moyenne", ascending=False).iterrows():
                    mention, mention_cls = get_mention(row["Moyenne"])
                    trend_txt, trend_color = get_trend(row["Devoir1"], row["Devoir2"], row["Devoir3"])
                    prog = row["Devoir3"] - row["Devoir1"]
                    prog_str = f"+{prog:.1f}" if prog > 0 else f"{prog:.1f}"

                    act_val  = row.get("Activites", None) if isinstance(row, dict) else row["Activites"] if "Activites" in row.index else None
                    act_html = f"&nbsp;&nbsp; Act: <b style='color:#7B5800'>{act_val}</b>" if pd.notna(act_val) else ""
                    st.markdown(f"""
                    <div style="background:#F8FAFC;border:1px solid #E0E7EF;border-radius:12px;
                                padding:14px 18px;margin-bottom:10px;display:flex;
                                justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
                        <div style="flex:1">
                            <div style="font-weight:800;font-size:15px;color:#0D1B2A">{row['Etudiant']}</div>
                            <div style="font-size:12px;color:#546E7A;margin-top:3px">
                                D1: <b>{row['Devoir1']}</b> &nbsp;→&nbsp;
                                D2: <b>{row['Devoir2']}</b> &nbsp;→&nbsp;
                                D3: <b>{row['Devoir3']}</b>
                                {act_html}
                                &nbsp;&nbsp; Progression D1→D3 : <b style="color:{'#2E7D32' if prog>0 else '#C62828'}">{prog_str} pts</b>
                            </div>
                            {progress_bar(row['Moyenne'], 20, trend_color)}
                        </div>
                        <div style="text-align:right">
                            <div style="font-size:24px;font-weight:900;color:#0D47A1">{row['Moyenne']:.2f}<span style="font-size:12px;color:#546E7A">/20</span></div>
                            <span class="mention {mention_cls}">{mention}</span>
                            <div style="font-size:12px;margin-top:4px">{trend_txt}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

            # ── TAB 2 : Graphiques ──────────────────────────────────────────
            with tab2:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📈 Évolution des notes par étudiant</div>', unsafe_allow_html=True)

                # Graphique évolution
                df_plot = df.set_index("Etudiant")[["Devoir1","Devoir2","Devoir3"]].T
                df_plot.index = ["Devoir 1", "Devoir 2", "Devoir 3"]
                st.line_chart(df_plot, height=350)
                st.markdown('</div>', unsafe_allow_html=True)

                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">📊 Moyenne par étudiant</div>', unsafe_allow_html=True)
                    df_bar = df.set_index("Etudiant")["Moyenne"].sort_values(ascending=False)
                    st.bar_chart(df_bar, height=300)
                    st.markdown('</div>', unsafe_allow_html=True)

                with col_g2:
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    st.markdown('<div class="card-title">📊 Progression D1 → D3</div>', unsafe_allow_html=True)
                    df_prog = df.set_index("Etudiant")["Progression"].sort_values(ascending=False)
                    st.bar_chart(df_prog, height=300)
                    st.markdown('</div>', unsafe_allow_html=True)

                # Répartition mentions
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">🏅 Répartition des mentions</div>', unsafe_allow_html=True)
                mentions_count = {}
                for _, row in df.iterrows():
                    m, _ = get_mention(row["Moyenne"])
                    mentions_count[m] = mentions_count.get(m, 0) + 1
                df_mentions = pd.DataFrame(list(mentions_count.items()), columns=["Mention","Nombre"])
                st.bar_chart(df_mentions.set_index("Mention"), height=250)
                st.markdown('</div>', unsafe_allow_html=True)

            # ── TAB 3 : Classement ──────────────────────────────────────────
            with tab3:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">🏆 Classement général</div>', unsafe_allow_html=True)

                df_rank = df.sort_values("Moyenne", ascending=False).reset_index(drop=True)
                df_rank["Rang"]      = df_rank.index + 1
                df_rank["Mention"]   = df_rank["Moyenne"].apply(lambda x: get_mention(x)[0])
                # Progression = D3 - D1 uniquement (3 premiers devoirs)
                df_rank["Prog D1→D3"] = (df_rank["Devoir3"] - df_rank["Devoir1"]).apply(
                    lambda x: f"+{x:.1f}" if x > 0 else f"{x:.1f}")
                df_rank["Tendance"]  = df_rank.apply(
                    lambda r: get_trend(r["Devoir1"], r["Devoir2"], r["Devoir3"])[0], axis=1)
                df_rank["Moy."]      = df_rank["Moyenne"].round(2)
                has_act = "Activites" in df_rank.columns and df_rank["Activites"].notna().any()

                # ── Tableau HTML complet sans scrolling ──
                act_header = "<th>Activités</th>" if has_act else ""
                rows_html  = ""
                for _, r in df_rank.iterrows():
                    rang = int(r["Rang"])
                    if rang == 1:   medal = "🥇"
                    elif rang == 2: medal = "🥈"
                    elif rang == 3: medal = "🥉"
                    else:           medal = str(rang)

                    moy = r["Moy."]
                    if moy >= 14:   bg = "#E8F5E9"
                    elif moy >= 10: bg = "#E3F2FD"
                    else:           bg = "#FFEBEE"

                    prog_val = r["Prog D1→D3"]
                    prog_color = "#2E7D32" if "+" in str(prog_val) else "#C62828"
                    mention, m_cls = get_mention(moy)

                    act_cell = f"<td style='text-align:center;font-weight:700;color:#7B5800'>{r['Activites']}</td>" if has_act else ""

                    rows_html += f"""
                    <tr style="background:{bg}">
                        <td style="text-align:center;font-weight:800;font-size:16px">{medal}</td>
                        <td style="font-weight:700;padding-left:8px">{r["Etudiant"]}</td>
                        <td style="text-align:center">{r["Devoir1"]}</td>
                        <td style="text-align:center">{r["Devoir2"]}</td>
                        <td style="text-align:center">{r["Devoir3"]}</td>
                        {act_cell}
                        <td style="text-align:center;font-weight:900;color:#0D47A1;font-size:16px">{moy}</td>
                        <td style="text-align:center"><span style="background:white;border-radius:12px;padding:2px 10px;font-size:11px;font-weight:700">{mention}</span></td>
                        <td style="text-align:center;font-weight:700;color:{prog_color}">{prog_val}</td>
                        <td style="text-align:center">{r["Tendance"]}</td>
                    </tr>"""

                st.markdown(f"""
                <table style="width:100%;border-collapse:collapse;font-size:13px;font-family:Outfit,sans-serif">
                    <thead>
                        <tr style="background:#0D47A1;color:white">
                            <th style="padding:10px 6px">Rang</th>
                            <th style="padding:10px 8px;text-align:left">Étudiant</th>
                            <th style="padding:10px 6px">D1</th>
                            <th style="padding:10px 6px">D2</th>
                            <th style="padding:10px 6px">D3</th>
                            {act_header}
                            <th style="padding:10px 6px">Moy.</th>
                            <th style="padding:10px 6px">Mention</th>
                            <th style="padding:10px 6px">Prog D1→D3</th>
                            <th style="padding:10px 6px">Tendance</th>
                        </tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
                """, unsafe_allow_html=True)

                # Top 3
                st.markdown("---")
                st.markdown("**🥇 Top 3 étudiants**")
                top3_cols = st.columns(3)
                medals_top = ["🥇","🥈","🥉"]
                for idx, (medal, col) in enumerate(zip(medals_top, top3_cols)):
                    if idx < len(df_rank):
                        r = df_rank.iloc[idx]
                        mention, _ = get_mention(float(r["Moy."]))
                        col.metric(f"{medal} {r['Etudiant']}", f"{r['Moy.']}/20", mention)

                st.markdown('</div>', unsafe_allow_html=True)

            # ── TAB 4 : Export ──────────────────────────────────────────────
            with tab4:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📄 Exporter le bulletin complet</div>', unsafe_allow_html=True)

                st.info("Les exports incluent : toutes les notes, moyennes, mentions, classement et progression.")

                dl1, dl2 = st.columns(2)
                with dl1:
                    st.markdown('<div class="dl-excel">', unsafe_allow_html=True)
                    try:
                        xlsx = make_excel_bulletin(df)
                        st.download_button("📊 Télécharger Excel",
                            data=xlsx,
                            file_name=f"Bulletin_Notes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_excel_notes")
                    except ImportError:
                        st.info("Ajoutez `openpyxl` à requirements.txt")
                    st.markdown('</div>', unsafe_allow_html=True)

                with dl2:
                    st.markdown('<div class="dl-pdf">', unsafe_allow_html=True)
                    try:
                        pdf = make_pdf_bulletin(df)
                        st.download_button("📄 Télécharger PDF",
                            data=pdf,
                            file_name=f"Bulletin_Notes_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf",
                            key="dl_pdf_notes")
                    except ImportError:
                        st.info("Ajoutez `reportlab` à requirements.txt")
                    st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier : {e}")
        st.info("Vérifiez que le fichier contient bien les colonnes : Etudiant, Devoir1, Devoir2, Devoir3")

else:
    st.markdown("""
    <div style="text-align:center;padding:48px 20px;color:#546E7A">
        <div style="font-size:56px;margin-bottom:16px">📂</div>
        <div style="font-size:18px;font-weight:700;color:#0D47A1;margin-bottom:8px">
            Importez votre fichier de notes
        </div>
        <div style="font-size:14px">
            Formats acceptés : <b>Excel (.xlsx, .xls)</b> ou <b>CSV (.csv)</b><br>
            Téléchargez le fichier exemple ci-dessus pour voir le format attendu
        </div>
    </div>
    """, unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="app-footer">
    <span class="badge">Suivi Notes v1.0</span> &nbsp;|&nbsp;
    Analyse de progression étudiante &nbsp;|&nbsp;
    {datetime.now().strftime('%Y')}
</div>
""", unsafe_allow_html=True)
