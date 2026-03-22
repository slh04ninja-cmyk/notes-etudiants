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
/* Browse files button → blanc */
[data-testid="stFileUploader"] button {
    background: var(--blue) !important;
    color: white !important;
    border: none !important;
    font-weight: 700 !important;
}
[data-testid="stFileUploader"] button:hover {
    background: var(--blue2) !important;
}
/* Nom du fichier uploadé → noir */
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] small,
[data-testid="uploadedFileData"] span {
    color: #1A1A1A !important;
    font-weight: 600 !important;
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

def _ar(text):
    """Convertit texte arabe pour affichage correct dans PDF (RTL + reshaping)."""
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

def _clean(text):
    """Supprime les emojis et caractères non-latin pour PDF."""
    import re
    # Remplacer emojis tendance par texte
    replacements = {
        "📈": "(+)", "📉": "(-)", "➡️": "(=)",
        "🥇": "1er", "🥈": "2eme", "🥉": "3eme",
        "🎓": "", "📊": "", "📄": "",
    }
    for emoji, txt in replacements.items():
        text = text.replace(emoji, txt)
    # Supprimer emojis restants
    text = re.sub(r'[^-éàèùâêîôûäëïöüçœæ؀-ۿ -~+\-./()%:,0-9 ]', '', text)
    return text.strip()

def make_pdf_bulletin(df_result):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import os, urllib.request

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)

    # ── Police Arabic ──
    arabic_font = "Helvetica"
    arabic_font_bold = "Helvetica-Bold"
    font_dir = "/tmp/fonts"
    os.makedirs(font_dir, exist_ok=True)
    amiri_path      = f"{font_dir}/Amiri-Regular.ttf"
    amiri_bold_path = f"{font_dir}/Amiri-Bold.ttf"

    try:
        if not os.path.exists(amiri_path):
            urllib.request.urlretrieve(
                "https://github.com/alif-type/amiri/raw/main/Amiri-Regular.ttf",
                amiri_path)
        if not os.path.exists(amiri_bold_path):
            urllib.request.urlretrieve(
                "https://github.com/alif-type/amiri/raw/main/Amiri-Bold.ttf",
                amiri_bold_path)
        pdfmetrics.registerFont(TTFont("Amiri",     amiri_path))
        pdfmetrics.registerFont(TTFont("AmiriBold", amiri_bold_path))
        arabic_font      = "Amiri"
        arabic_font_bold = "AmiriBold"
    except Exception:
        pass  # Fallback Helvetica si téléchargement échoue

    BLUE  = colors.HexColor("#1565C0")
    BLUE2 = colors.HexColor("#0D47A1")
    SKY   = colors.HexColor("#E3F2FD")
    GOLD  = colors.HexColor("#FFF8E1")
    GREEN = colors.HexColor("#E8F5E9")
    RED   = colors.HexColor("#FFEBEE")

    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t",  fontSize=14, fontName="Helvetica-Bold", textColor=colors.white)
    sub_s   = ParagraphStyle("s",  fontSize=9,  fontName="Helvetica",      textColor=colors.HexColor("#B0BEC5"))
    stat_s  = ParagraphStyle("st", fontSize=10, fontName="Helvetica-Bold", textColor=BLUE2, spaceAfter=4)
    ar_s    = ParagraphStyle("ar", fontSize=9,  fontName=arabic_font,      alignment=2)  # RTL

    story = []

    # ── Header ──
    ht = Table([[
        Paragraph("Bulletin de Notes — Suivi de Progression", title_s),
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

    # ── Stats ──
    moy_classe = df_result["Moyenne"].mean()
    nb_reussi  = (df_result["Moyenne"] >= 10).sum()
    nb_echec   = (df_result["Moyenne"] < 10).sum()
    story.append(Paragraph(
        f"<b>Classe :</b> {len(df_result)} etudiants  |  "
        f"<b>Moyenne classe :</b> {moy_classe:.2f}/20  |  "
        f"<b>Reussite :</b> {nb_reussi} ({nb_reussi/len(df_result)*100:.0f}%)  |  "
        f"<b>Echec :</b> {nb_echec}", stat_s))
    story.append(Spacer(1, 0.3*cm))

    # ── Détection Activités ──
    has_act = "Activites" in df_result.columns and df_result["Activites"].notna().any()

    # ── En-têtes tableau ──
    if has_act:
        headers = ["N°", "Nom Etudiant", "D1/20", "D2/20", "D3/20",
                   "Act./20", "Moy./20", "Mention", "Tendance", "Prog D1-D3"]
        col_w = [0.8*cm, 5.5*cm, 1.8*cm, 1.8*cm, 1.8*cm, 1.8*cm, 2*cm, 2.8*cm, 2.5*cm, 2*cm]
    else:
        headers = ["N°", "Nom Etudiant", "D1/20", "D2/20", "D3/20",
                   "Moy./20", "Mention", "Tendance", "Prog D1-D3"]
        col_w = [0.8*cm, 6*cm, 2*cm, 2*cm, 2*cm, 2.2*cm, 3*cm, 3*cm, 2.5*cm]

    tdata = [headers]

    df_sorted = df_result.sort_values("Moyenne", ascending=False).reset_index(drop=True)
    for i, row in enumerate(df_sorted.itertuples(), 1):
        mention, _ = get_mention(row.Moyenne)
        prog       = row.Devoir3 - row.Devoir1
        prog_str   = f"+{prog:.1f}" if prog > 0 else f"{prog:.1f}"

        # Tendance sans emoji
        diff = row.Devoir3 - row.Devoir1
        if diff > 1:    trend = "En progres"
        elif diff < -1: trend = "En baisse"
        else:           trend = "Stable"

        # Nom arabe correctement formé
        nom_ar = _ar(row.Etudiant)

        if has_act:
            act_val = f"{row.Activites:.2f}" if pd.notna(row.Activites) else "-"
            tdata.append([str(i), nom_ar,
                str(row.Devoir1), str(row.Devoir2), str(row.Devoir3),
                act_val, f"{row.Moyenne:.2f}", mention, trend, prog_str])
        else:
            tdata.append([str(i), nom_ar,
                str(row.Devoir1), str(row.Devoir2), str(row.Devoir3),
                f"{row.Moyenne:.2f}", mention, trend, prog_str])

    # Ligne moyenne
    empty = ["—"] * (len(headers) - 3)
    tdata.append(["—", "MOYENNE CLASSE", "—", "—", "—"] + empty + [f"{moy_classe:.2f}", "—", "—", "—"]
                 if not has_act else
                 ["—", "MOYENNE CLASSE", "—", "—", "—", "—", f"{moy_classe:.2f}", "—", "—", "—"])

    t = Table(tdata, colWidths=col_w, repeatRows=1)

    ts_list = [
        ("BACKGROUND",    (0,0),  (-1,0),  BLUE2),
        ("TEXTCOLOR",     (0,0),  (-1,0),  colors.white),
        ("FONTNAME",      (0,0),  (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),  (-1,-1), 8),
        ("ALIGN",         (0,0),  (-1,-1), "CENTER"),
        ("GRID",          (0,0),  (-1,-1), 0.4, colors.HexColor("#CFD8DC")),
        ("TOPPADDING",    (0,0),  (-1,-1), 5),
        ("BOTTOMPADDING", (0,0),  (-1,-1), 5),
        ("LEFTPADDING",   (0,0),  (-1,-1), 4),
        ("BACKGROUND",    (0,-1), (-1,-1), GOLD),
        ("FONTNAME",      (0,-1), (-1,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0,-1), (-1,-1), colors.HexColor("#7B5800")),
        # Colonne noms : police arabe + alignement droite
        ("FONTNAME",      (1,1),  (1,-2),  arabic_font),
        ("ALIGN",         (1,1),  (1,-2),  "RIGHT"),
    ]

    for i, row in enumerate(df_sorted.itertuples(), 1):
        if row.Moyenne >= 14:   fill = GREEN
        elif row.Moyenne >= 10: fill = SKY
        else:                   fill = RED
        ts_list.append(("BACKGROUND", (0,i), (-1,i), fill))

    t.setStyle(TableStyle(ts_list))
    story.append(t)
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(
        "<font color='#546E7A' size='7'>Application Suivi des Notes v1.0 — Genere automatiquement</font>",
        styles["Normal"]))

    doc.build(story)
    buf.seek(0)
    return buf

# ── Parsers fichiers ─────────────────────────────────────────────────────────
def _detect_format(df_raw):
    """Détecte le format (MASSAR ou standard) depuis un DataFrame pandas brut."""
    # Convertir tout en string pour chercher les marqueurs arabes
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row if pd.notna(v)])
        if "الفرض الأول" in row_str or "إسم التلميذ" in row_str:
            # Format MASSAR tr
