import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
import json
from datetime import datetime
from PIL import Image
import requests

import folium
from streamlit_folium import st_folium

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ReportLab para geração de PDF ABNT & Profissional
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# Componente de Assinatura
from streamlit_drawable_canvas import st_canvas
from streamlit_javascript import st_javascript

# Ocultar elementos padrão do Streamlit
ocultar_elementos = """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppHeader {display: none;}
    </style>
"""
st.markdown(ocultar_elementos, unsafe_allow_html=True)

st.set_page_config(
    page_title="Boletim de Sondagem Mineral",
    page_icon="⛏️",
    layout="wide"
)

# Estilização do Streamlit
st.markdown("""
    <style>
    div[data-testid="stVerticalBlock"] > div[data-testid="stBlock"] {
        background: #FFFFFF;
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        border: 1px solid #E2E8F0;
    }
    h1 {
        color: #0F172A !important;
        background: linear-gradient(135deg, #E0F2FE 0%, #F0F9FF 100%);
        padding: 16px 20px;
        border-radius: 14px;
        border-left: 6px solid #0284C7;
        font-weight: 800 !important;
    }
    h2, h3 {
        color: #0369A1 !important;
        font-weight: 700 !important;
    }
    .stButton > button {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 10px 24px !important;
        font-weight: 700 !important;
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# DICIONÁRIO LITOLÓGICO
DADOS_LITOLOGIA = {
    'Solo / Cobertura':       {'cor': '#E5D3B3', 'hatch': '....'},
    'Siltito / Argilito':     {'cor': '#D2B48C', 'hatch': '----'},
    'Quartzito':              {'cor': '#FFF8DC', 'hatch': '////'},
    'Schisto / Filito':       {'cor': '#94A3B8', 'hatch': '\\\\\\\\'},
    'Gnaisse / Granito':      {'cor': '#E2E8F0', 'hatch': '++++'},
    'Basalto / Diabásio':     {'cor': '#475569', 'hatch': 'xxxx'},
    'Minério de Ferro / BIF': {'cor': '#991B1B', 'hatch': '||||'},
    'Calcário / Dolomito':    {'cor': '#BAE6FD', 'hatch': 'OOOO'},
    'Outro':                  {'cor': '#CBD5E1', 'hatch': ''}
}

if 'manobras' not in st.session_state:
    st.session_state['manobras'] = []

st.title("📋 Boletim Digital de Sondagem Mineral")
st.markdown("---")

# 1. DADOS DE GESTÃO E LOCALIZAÇÃO
st.header("1. Cabeçalho do Projeto & Equipe Técnica")

col_g1, col_g2, col_g3, col_g4 = st.columns(4)
with col_g1:
    empresa = st.text_input("Empresa / Mineradora", value="Mineração Picuí S.A.")
    projeto = st.text_input("Nome do Projeto", value="Projeto Picuí")
with col_g2:
    coordenador = st.text_input("Coordenador do Projeto", value="Eng. Carlos Andrade")
    supervisor = st.text_input("Supervisor de Campo", value="Téc. Roberto Lima")
with col_g3:
    geologo = st.text_input("Geólogo Responsável", value="Geól. Mariana Costa")
    sondador = st.text_input("Sondador / Equipe", value="Natanael & Equipe")
with col_g4:
    furo_id = st.text_input("ID do Furo", value="F-001")
    diametro = st.selectbox("Diâmetro", ['HQ (63.5mm)', 'NQ (47.6mm)', 'BQ (36.5mm)', 'RC (Circ. Reversa)', 'Outro'])

with st.expander("🌐 Coordenadas GPS e Mapa do Furo", expanded=True):
    col_geo1, col_geo2, col_geo3, col_geo4 = st.columns(4)
    with col_geo1:
        utm_e = st.number_input("Coordenada UTM (E)", value=250100.0, format="%.2f")
        utm_n = st.number_input("Coordenada UTM (N)", value=9245000.0, format="%.2f")
    with col_geo2:
        cota_z = st.number_input("Cota Z (m)", value=480.5, format="%.2f")
        datum = st.text_input("Datum", value="SIRGAS 2000")
    with col_geo3:
        inclinacao = st.number_input("Inclinação (°)", value=-90.0, format="%.1f")
        azimute = st.number_input("Azimute (°)", value=0.0, format="%.1f")
    with col_geo4:
        data_inicio = st.date_input("Data de Início", value=datetime.now())
        data_fim = st.date_input("Data de Término", value=datetime.now())

    st.markdown("---")
    col_gps1, col_gps2 = st.columns(2)
    with col_gps1:
        lat_furo = st.number_input("Latitude", value=-6.512345, format="%.6f")
    with col_gps2:
        lon_furo = st.number_input("Longitude", value=-36.512345, format="%.6f")
    
    m = folium.Map(location=[lat_furo, lon_furo], zoom_start=16, tiles=None)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri World Imagery', name='Satélite (Esri)', overlay=False
    ).add_to(m)
    folium.Marker([lat_furo, lon_furo], popup=f"Furo: {furo_id}", icon=folium.Icon(color='red')).add_to(m)
    st_folium(m, width="100%", height=300)

st.markdown("---")

# 2. REGISTRO DE MANOBRAS E FOTOS
st.header("2. Registro de Manobras e Fotos do Testemunho")

prox_de = st.session_state['manobras'][-1]['Para (m)'] if st.session_state['manobras'] else 0.0
prox_para = round(prox_de + 1.5, 2)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    de = st.number_input("De (m)", value=float(prox_de), step=0.5, format="%.2f")
with col_m2:
    para = st.number_input("Para (m)", value=float(prox_para), step=0.5, format="%.2f")
with col_m3:
    rec = st.number_input("Rec. (m)", value=round(para - de, 2), step=0.1, format="%.2f")
with col_m4:
    rqd = st.number_input("RQD (m)", value=round((para - de) * 0.8, 2), step=0.1, format="%.2f")

col_l1, col_l2, col_l3 = st.columns(3)
with col_l1:
    litologia = st.selectbox("Litologia", list(DADOS_LITOLOGIA.keys()))
with col_l2:
    alteracao = st.selectbox("Alteração", ['Solo / Inconsol.', 'Completamente Alterada', 'Muito Alterada', 'Moderadamente Alterada', 'Pouco Alterada', 'Rocha Sã'])
with col_l3:
    obs = st.text_input("Observações Geotécnicas", placeholder="Ex: RPT, Fraturado, veios de quartzo...")

st.subheader("📷 Registro Fotográfico da Amostra / Caixa")
aba_cam, aba_up = st.tabs(["📸 Tirar Foto Agora", "📁 Carregar da Galeria"])

img_capturada = None
with aba_cam:
    foto_cam = st.camera_input("Tirar foto da caixa de testemunho")
    if foto_cam: img_capturada = Image.open(foto_cam)

with aba_up:
    foto_file = st.file_uploader("Selecione uma imagem", type=['jpg', 'jpeg', 'png'])
    if foto_file and not img_capturada: img_capturada = Image.open(foto_file)

col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    btn_adicionar = st.button("➕ Adicionar Manobra", type="primary")
with col_btn2:
    btn_remover = st.button("🗑️ Remover Última")

if btn_adicionar:
    avanco = round(para - de, 2)
    if avanco <= 0:
        st.error("⚠️ O valor 'Para' deve ser maior que 'De'!")
    else:
        pct_rec = min(100.0, round((rec / avanco) * 100, 1)) if avanco > 0 else 0.0
        pct_rqd = min(100.0, round((rqd / avanco) * 100, 1)) if avanco > 0 else 0.0
        
        if pct_rqd < 25: rqd_class = 'Muito Pobre'
        elif pct_rqd < 50: rqd_class = 'Pobre'
        elif pct_rqd < 75: rqd_class = 'Razoável'
        elif pct_rqd < 90: rqd_class = 'Boa'
        else: rqd_class = 'Excelente'

        st.session_state['manobras'].append({
            'Manobra': len(st.session_state['manobras']) + 1,
            'De (m)': de, 'Para (m)': para, 'Avanço (m)': avanco,
            'Rec. (m)': rec, 'Rec (%)': pct_rec, 'RQD (m)': rqd,
            'RQD (%)': pct_rqd, 'Qualidade RQD': rqd_class,
            'Litologia': litologia, 'Alteração': alteracao, 
            'Observações': obs, 'Foto': img_capturada
        })
        st.success("✅ Manobra registrada!")
        st.rerun()

if btn_remover and st.session_state['manobras']:
    st.session_state['manobras'].pop()
    st.warning("🗑️ Última manobra removida.")
    st.rerun()

st.markdown("---")

# 3. PERFIL VISUAL E TABELA
st.header("3. Perfil Litológico e Relatórios")

if st.session_state['manobras']:
    df_manobras = pd.DataFrame(st.session_state['manobras'])

    # GERAÇÃO DO GRÁFICO DO PERFIL
    plt.rcParams['hatch.linewidth'] = 1.2
    plt.rcParams['hatch.color'] = '#333333'
    fig, (ax_lito, ax_rqd, ax_rec) = plt.subplots(1, 3, figsize=(11, 4.5), sharey=True, gridspec_kw={'width_ratios': [1.3, 2, 2]})
    
    prof_max = df_manobras['Para (m)'].max()
    ax_lito.set_ylim(prof_max, 0)
    litos_usadas = set()

    for _, row in df_manobras.iterrows():
        de_m, para_m, lito = row['De (m)'], row['Para (m)'], row['Litologia']
        rqd_val, rec_val = row['RQD (%)'], row['Rec (%)']
        info_lito = DADOS_LITOLOGIA.get(lito, {'cor': '#808080', 'hatch': ''})
        litos_usadas.add(lito)

        rect = mpatches.Rectangle((0, de_m), 1, para_m - de_m, facecolor=info_lito['cor'], hatch=info_lito['hatch'], edgecolor='#1E293B', linewidth=1.2)
        ax_lito.add_patch(rect)
        ax_lito.axhline(para_m, color='#0F172A', linestyle='--', linewidth=0.8)
        ax_lito.text(0.5, (de_m + para_m)/2, f"{lito}\n({de_m:.1f}m - {para_m:.1f}m)", ha='center', va='center', fontsize=8, fontweight='bold', bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', alpha=0.85, edgecolor='#94A3B8'))

        color_rqd = '#EF4444' if rqd_val < 25 else '#F97316' if rqd_val < 50 else '#EAB308' if rqd_val < 75 else '#3B82F6' if rqd_val < 90 else '#22C55E'
        ax_rqd.barh(y=de_m + (para_m - de_m)/2, width=rqd_val, height=(para_m - de_m)*0.8, color=color_rqd, edgecolor='black', linewidth=0.8)
        ax_rec.barh(y=de_m + (para_m - de_m)/2, width=rec_val, height=(para_m - de_m)*0.8, color='#0284C7', edgecolor='black', linewidth=0.8)

    ax_lito.set_xlim(0, 1)
    ax_lito.set_title("Estratigrafia", fontsize=10, fontweight='bold')
    ax_lito.set_ylabel("Profundidade (m)", fontsize=9, fontweight='bold')
    ax_lito.get_xaxis().set_visible(False)
    ax_rqd.set_xlim(0, 105)
    ax_rqd.set_title("RQD (%)", fontsize=10, fontweight='bold')
    ax_rec.set_xlim(0, 105)
    ax_rec.set_title("Recuperação (%)", fontsize=10, fontweight='bold')
    plt.tight_layout()

    st.pyplot(fig)

    st.dataframe(df_manobras.drop(columns=['Foto']), use_container_width=True, hide_index=True)

    st.markdown("### ✍️ Assinatura Digital do Responsável")
    canvas_result = st_canvas(fill_color="rgba(255, 255, 255, 0)", stroke_width=2, stroke_color="#000000", background_color="#F8FAFC", height=120, width=400, drawing_mode="freedraw", key="canvas_assinatura")

    col_exp1, col_exp2 = st.columns(2)

    # --- EXPORTAÇÃO EXCEL ---
    with col_exp1:
        # Mantida a lógica do Excel funcional e estilizada
        buffer_xls = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Boletim ABNT"
        ws.views.sheetView[0].showGridLines = True
        
        ws.merge_cells('A1:L1')
        ws['A1'] = f"RELATÓRIO TÉCNICO DE SONDAGEM - {empresa.upper()}"
        ws['A1'].font = Font(name='Calibri', size=14, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center')

        df_excel = df_manobras.drop(columns=['Foto'])
        for c_idx, col_name in enumerate(df_excel.columns, 1):
            ws.cell(row=3, column=c_idx, value=col_name).font = Font(bold=True)

        for r_idx, row in df_excel.iterrows():
            for c_idx, val in enumerate(row.values, 1):
                ws.cell(row=r_idx + 4, column=c_idx, value=val)

        wb.save(buffer_xls)
        st.download_button("📄 Baixar Planilha (.xlsx)", buffer_xls.getvalue(), f"Boletim_{furo_id}.xlsx", use_container_width=True)

    # --- EXPORTAÇÃO PDF NORMA ABNT ---
    with col_exp2:
        
        # CLASSE CANVAS CUSTOMIZADA PARA PAGINAÇÃO ABNT
        class NumberedCanvas(canvas.Canvas):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self._saved_page_states = []

            def showPage(self):
                self._saved_page_states.append(dict(self.__dict__))
                self._startPage()

            def save(self):
                num_pages = len(self._saved_page_states)
                for state in self._saved_page_states:
                    self.__dict__.update(state)
                    self.draw_page_number(num_pages)
                    super().showPage()
                super().save()

            def draw_page_number(self, page_count):
                # ABNT: Numeração no canto superior direito a partir da página 2
                if self._pageNumber > 1:
                    self.setFont("Times-Roman", 10)
                    text = f"Página {self._pageNumber} de {page_count}"
                    # Margem direita = 21.0cm - 2.0cm = 19.0cm
                    self.drawRightString(19.0 * cm, 28.0 * cm, text)

        pdf_buf = io.BytesIO()
        
        # Margens ABNT NBR 14724: Sup: 3cm, Esq: 3cm, Dir: 2cm, Inf: 2cm
        doc = SimpleDocTemplate(
            pdf_buf, pagesize=portrait(A4),
            leftMargin=3.0*cm, rightMargin=2.0*cm,
            topMargin=3.0*cm, bottomMargin=2.0*cm
        )
        elements = []
        styles = getSampleStyleSheet()

        # ESTILOS CONFORME REGRA ABNT (Fonte Padrão: Times-Roman)
        abnt_titulo_doc = ParagraphStyle('ABNTTituloDoc', parent=styles['Heading1'], fontName='Times-Bold', fontSize=14, leading=16, alignment=1, spaceAfter=15)
        abnt_sec = ParagraphStyle('ABNTSec', parent=styles['Heading2'], fontName='Times-Bold', fontSize=12, leading=14, spaceBefore=12, spaceAfter=6)
        abnt_text = ParagraphStyle('ABNTText', parent=styles['Normal'], fontName='Times-Roman', fontSize=10, leading=13)
        abnt_text_bold = ParagraphStyle('ABNTTextBold', parent=styles['Normal'], fontName='Times-Bold', fontSize=10, leading=13)
        abnt_caption = ParagraphStyle('ABNTCaption', parent=styles['Italic'], fontName='Times-Italic', fontSize=9, leading=11, alignment=1, spaceAfter=4)
        abnt_fonte = ParagraphStyle('ABNTFonte', parent=styles['Italic'], fontName='Times-Roman', fontSize=8, leading=10, alignment=1, spaceBefore=4, spaceAfter=10)

        # 1. CABEÇALHO DO DOCUMENTO TÉCNICO ABNT
        elements.append(Paragraph(f"<b>{empresa.upper()}</b>", abnt_titulo_doc))
        elements.append(Paragraph(f"<b>RELATÓRIO TÉCNICO DE SONDAGEM GEOLÓGICA - FURO {furo_id}</b>", ParagraphStyle('Sub', parent=abnt_titulo_doc, fontSize=12, spaceAfter=20)))

        # 2. SEÇÃO 1: DADOS DE GESTÃO E LOCALIZAÇÃO
        elements.append(Paragraph("<b>1. DADOS DE GESTÃO E LOCALIZAÇÃO DO FURO</b>", abnt_sec))
        
        prof_total_val = float(df_manobras['Para (m)'].max())
        dados_furo_table = [
            [Paragraph("<b>Projeto:</b>", abnt_text), Paragraph(projeto, abnt_text), Paragraph("<b>Coordenador:</b>", abnt_text), Paragraph(coordenador, abnt_text)],
            [Paragraph("<b>ID do Furo:</b>", abnt_text), Paragraph(furo_id, abnt_text_bold), Paragraph("<b>Supervisor:</b>", abnt_text), Paragraph(supervisor, abnt_text)],
            [Paragraph("<b>Diâmetro:</b>", abnt_text), Paragraph(diametro, abnt_text), Paragraph("<b>Geólogo Resp.:</b>", abnt_text), Paragraph(geologo, abnt_text)],
            [Paragraph("<b>Início / Fim:</b>", abnt_text), Paragraph(f"{data_inicio} a {data_fim}", abnt_text), Paragraph("<b>Sondador:</b>", abnt_text), Paragraph(sondador, abnt_text)],
            [Paragraph("<b>UTM (E / N):</b>", abnt_text), Paragraph(f"{utm_e:.2f} / {utm_n:.2f}", abnt_text), Paragraph("<b>Cota Z / Datum:</b>", abnt_text), Paragraph(f"{cota_z:.2f} m ({datum})", abnt_text)],
            [Paragraph("<b>Inclin. / Azimute:</b>", abnt_text), Paragraph(f"{inclinacao}° / {azimute}°", abnt_text), Paragraph("<b>Prof. Total:</b>", abnt_text), Paragraph(f"{prof_total_val:.2f} m", abnt_text_bold)],
        ]
        
        # Largura útil da página = 21.0 cm - 3.0 cm (Esq) - 2.0 cm (Dir) = 16.0 cm
        t_furo = Table(dados_furo_table, colWidths=[3.2*cm, 4.8*cm, 3.2*cm, 4.8*cm])
        t_furo.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#000000')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elements.append(t_furo)
        elements.append(Spacer(1, 10))

        # 3. IMAGEM DE SATÉLITE (LEGENDA E FONTE PADRÃO ABNT)
        elements.append(Paragraph("<b>1.1 Localização Real de Satélite</b>", abnt_sec))
        try:
            url_satelite = f"https://static-maps.yandex.ru/1.x/?lang=pt_BR&ll={lon_furo},{lat_furo}&z=15&size=600,200&l=sat&pt={lon_furo},{lat_furo},pm2rdm"
            res = requests.get(url_satelite, timeout=5)
            if res.status_code == 200:
                img_sat_buf = io.BytesIO(res.content)
                img_sat_pdf = RLImage(img_sat_buf, width=16.0*cm, height=4.5*cm)
                
                elements.append(Paragraph(f"<b>Figura 1</b> – Imagem de satélite da localização do furo {furo_id}.", abnt_caption))
                elements.append(img_sat_pdf)
                elements.append(Paragraph("Fonte: Adaptado de Yandex Maps (2026).", abnt_fonte))
        except Exception:
            elements.append(Paragraph("<i>(Imagem de satélite indisponível no momento da geração)</i>", abnt_text))

        elements.append(Spacer(1, 10))

        # 4. SEÇÃO 2: TABELA DE MANOBRAS E GEOTECNIA (ABNT)
        elements.append(Paragraph("<b>2. REGISTRO DE MANOBRAS E DADOS GEOTÉCNICOS</b>", abnt_sec))
        elements.append(Paragraph("<b>Tabela 1</b> – Dados geotécnicos e litológicos obtidos nas manobras.", abnt_caption))

        table_manobras_data = [
            [Paragraph("<b>Mnb</b>", abnt_text_bold), Paragraph("<b>De (m)</b>", abnt_text_bold), Paragraph("<b>Para (m)</b>", abnt_text_bold), 
             Paragraph("<b>Av. (m)</b>", abnt_text_bold), Paragraph("<b>Rec. (m)</b>", abnt_text_bold), Paragraph("<b>Rec (%)</b>", abnt_text_bold), 
             Paragraph("<b>RQD (%)</b>", abnt_text_bold), Paragraph("<b>Litologia</b>", abnt_text_bold), Paragraph("<b>Observações</b>", abnt_text_bold)]
        ]

        for _, r in df_manobras.iterrows():
            table_manobras_data.append([
                Paragraph(str(int(r['Manobra'])), abnt_text), Paragraph(f"{r['De (m)']:.2f}", abnt_text),
                Paragraph(f"{r['Para (m)']:.2f}", abnt_text), Paragraph(f"{r['Avanço (m)']:.2f}", abnt_text),
                Paragraph(f"{r['Rec. (m)']:.2f}", abnt_text), Paragraph(f"{r['Rec (%)']:.1f}", abnt_text),
                Paragraph(f"{r['RQD (%)']:.1f}", abnt_text), Paragraph(str(r['Litologia']), abnt_text),
                Paragraph(str(r['Observações'] or '-'), abnt_text)
            ])

        # Linha de Totais/Médias
        table_manobras_data.append([
            Paragraph("<b>Total / Média</b>", abnt_text_bold), Paragraph("-", abnt_text), Paragraph("-", abnt_text),
            Paragraph(f"<b>{df_manobras['Avanço (m)'].sum():.2f}</b>", abnt_text_bold),
            Paragraph(f"<b>{df_manobras['Rec. (m)'].sum():.2f}</b>", abnt_text_bold),
            Paragraph(f"<b>{df_manobras['Rec (%)'].mean():.1f}%</b>", abnt_text_bold),
            Paragraph(f"<b>{df_manobras['RQD (%)'].mean():.1f}%</b>", abnt_text_bold),
            Paragraph("-", abnt_text), Paragraph("-", abnt_text)
        ])

        t_manobras = Table(table_manobras_data, colWidths=[1.1*cm, 1.4*cm, 1.4*cm, 1.4*cm, 1.4*cm, 1.5*cm, 1.5*cm, 3.1*cm, 3.2*cm])
        t_manobras.setStyle(TableStyle([
            ('LINEABOVE', (0,0), (-1,0), 1.0, colors.black),
            ('LINEBELOW', (0,0), (-1,0), 1.0, colors.black),
            ('LINEBELOW', (0,-1), (-1,-1), 1.0, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        
        elements.append(t_manobras)
        elements.append(Paragraph("Fonte: Dados do projeto (2026).", abnt_fonte))

        # FORCE PAGE BREAK PARA MANTER ORGANIZAÇÃO TÉCNICA
        elements.append(PageBreak())

        # 5. SEÇÃO 3: ESTRATIGRAFIA E PERFIL GEOLÓGICO
        elements.append(Paragraph("<b>3. PERFIL STRATIGRÁFICO E GRÁFICOS DE RECUPERAÇÃO/RQD</b>", abnt_sec))
        elements.append(Paragraph("<b>Figura 2</b> – Perfil geológico e variação dos parâmetros geotécnicos.", abnt_caption))

        img_plt_buf = io.BytesIO()
        fig.savefig(img_plt_buf, format='png', dpi=200, bbox_inches='tight')
        img_plt_buf.seek(0)
        
        img_perfil_pdf = RLImage(img_plt_buf, width=16.0*cm, height=7.5*cm)
        elements.append(img_perfil_pdf)
        elements.append(Paragraph("Fonte: Elaborado pelos autores a partir de medições de campo.", abnt_fonte))

        # 6. SEÇÃO 4: REGISTRO FOTOGRÁFICO
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("<b>4. REGISTRO FOTOGRÁFICO DOS TESTEMUNHOS DE SONDAGEM</b>", abnt_sec))
        
        fotos_list = [m['Foto'] for m in st.session_state['manobras'] if m['Foto'] is not None]
        if fotos_list:
            elements.append(Paragraph("<b>Figura 3</b> – Amostras e caixas de testemunho registradas.", abnt_caption))
            fotos_grid = []
            row_temp = []
            for idx, img_p in enumerate(fotos_list):
                img_b = io.BytesIO()
                img_p.save(img_b, format='PNG')
                img_b.seek(0)
                rl_img = RLImage(img_b, width=3.6*cm, height=2.6*cm)
                row_temp.append(rl_img)
                if len(row_temp) == 4:
                    fotos_grid.append(row_temp)
                    row_temp = []
            if row_temp:
                while len(row_temp) < 4:
                    row_temp.append(Paragraph("", abnt_text))
                fotos_grid.append(row_temp)

            t_fotos = Table(fotos_grid, colWidths=[4.0*cm, 4.0*cm, 4.0*cm, 4.0*cm])
            t_fotos.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            elements.append(t_fotos)
            elements.append(Paragraph("Fonte: Acervo fotográfico da amostragem (2026).", abnt_fonte))
        else:
            elements.append(Paragraph("<i>Nenhum registro fotográfico anexado.</i>", abnt_text))

        # 7. SEÇÃO 5: ASSINATURA TÉCNICA E RESPONSABILIDADE
        elements.append(Spacer(1, 15))
        
        elements_assinatura = []
        elements_assinatura.append(Paragraph("<b>5. ENCERRAMENTO E ASSINATURA TÉCNICA</b>", abnt_sec))
        
        if canvas_result.image_data is not None:
            img_ass = Image.fromarray(canvas_result.image_data.astype('uint8'))
            ass_buf = io.BytesIO()
            img_ass.save(ass_buf, format='PNG')
            ass_buf.seek(0)
            rl_ass = RLImage(ass_buf, width=5.5*cm, height=1.8*cm)
        else:
            rl_ass = Paragraph("<br/><br/>___________________________________", abnt_text)

        sig_table_data = [
            [rl_ass, Paragraph(f"<br/><br/>___________________________________<br/><b>{geologo}</b><br/>Geólogo Responsável", abnt_text)],
        ]
        t_sig = Table(sig_table_data, colWidths=[8.0*cm, 8.0*cm])
        t_sig.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elements_assinatura.append(t_sig)
        
        elements.append(KeepTogether(elements_assinatura))

        # CONSTRUÇÃO DO PDF USANDO O CANVAS ABNT CUSTOMIZADO
        doc.build(elements, canvasmaker=NumberedCanvas)
        
        st.download_button(
            label="📄 Baixar Boletim ABNT (.pdf)",
            data=pdf_buf.getvalue(),
            file_name=f"Boletim_ABNT_{furo_id}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
