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
from openpyxl.drawing.image import Image as OpenpyxlImage

# ReportLab para geração de PDF ABNT & Profissional
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# Componente de Assinatura
from streamlit_drawable_canvas import st_canvas

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
    'Solo / Cobertura':        {'cor': '#E5D3B3', 'hatch': '....'},
    'Siltito / Argilito':      {'cor': '#D2B48C', 'hatch': '----'},
    'Quartzito':               {'cor': '#FFF8DC', 'hatch': '////'},
    'Schisto / Filito':        {'cor': '#94A3B8', 'hatch': '\\\\\\\\'},
    'Gnaisse / Granito':       {'cor': '#E2E8F0', 'hatch': '++++'},
    'Basalto / Diabásio':      {'cor': '#475569', 'hatch': 'xxxx'},
    'Minério de Ferro / BIF': {'cor': '#991B1B', 'hatch': '||||'},
    'Calcário / Dolomito':     {'cor': '#BAE6FD', 'hatch': 'OOOO'},
    'Outro':                   {'cor': '#CBD5E1', 'hatch': ''}
}

if 'manobras' not in st.session_state:
    st.session_state['manobras'] = []

st.title("📋 Boletim Digital de Sondagem Mineral")
st.markdown("---")

# 1. DADOS DE GESTÃO E LOCALIZAÇÃO
st.header("1. Cabeçalho do Projeto & Equipe Técnica")

col_logo, col_gest = st.columns([1, 3])
with col_logo:
    st.subheader("🖼️ Logomarca da Empresa")
    logo_file = st.file_uploader("Carregar Logo (PNG/JPG)", type=['png', 'jpg', 'jpeg'])
    img_logo_pil = Image.open(logo_file) if logo_file else None
    if img_logo_pil:
        st.image(img_logo_pil, width=180)

with col_gest:
    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        empresa = st.text_input("Empresa / Mineradora", value="Mineração Picuí S.A.")
        projeto = st.text_input("Nome do Projeto", value="Projeto Picuí")
    with col_g2:
        coordenador = st.text_input("Coordenador do Projeto", value="Eng. Carlos Andrade")
        supervisor = st.text_input("Supervisor de Campo", value="Téc. Roberto Lima")
    with col_g3:
        geologo = st.text_input("Geólogo Responsável", value="Geól. Mariana Costa")
        sondador = st.text_input("Sondador / Equipe", value="Natanael & Equipe")

col_furo1, col_furo2 = st.columns(2)
with col_furo1:
    furo_id = st.text_input("ID do Furo", value="F-001")
with col_furo2:
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

    # --- EXPORTAÇÃO EXCEL COM LOGO E ASSINATURA ---
    with col_exp1:
        buffer_xls = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Boletim de Sondagem"
        ws.views.sheetView[0].showGridLines = True

        # Estilos Profissionais
        font_titulo = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        font_sub = Font(name='Calibri', size=10, italic=True, color='FFFFFF')
        font_sec = Font(name='Calibri', size=11, bold=True, color='0F172A')
        font_header = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        font_body = Font(name='Calibri', size=10)
        font_total = Font(name='Calibri', size=10, bold=True)

        fill_banner = PatternFill(start_color='0284C7', end_color='0284C7', fill_type='solid')
        fill_sec = PatternFill(start_color='E0F2FE', end_color='E0F2FE', fill_type='solid')
        fill_header = PatternFill(start_color='0369A1', end_color='0369A1', fill_type='solid')
        fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
        fill_total = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')

        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        double_bottom = Border(
            top=Side(style='thin', color='0F172A'),
            bottom=Side(style='double', color='0F172A')
        )

        align_center = Alignment(horizontal='center', vertical='center')
        align_left = Alignment(horizontal='left', vertical='center')
        align_right = Alignment(horizontal='right', vertical='center')

        # Insert Logo on Excel Top (If present)
        if img_logo_pil:
            img_logo_excel_buf = io.BytesIO()
            img_logo_pil.save(img_logo_excel_buf, format='PNG')
            img_logo_excel_buf.seek(0)
            xl_logo = OpenpyxlImage(img_logo_excel_buf)
            xl_logo.width = 110
            xl_logo.height = 40
            ws.add_image(xl_logo, 'A1')

        # 1. Banners de Título
        ws.merge_cells('C1:L1')
        ws['C1'] = empresa.upper()
        ws['C1'].font = font_titulo
        ws['C1'].fill = fill_banner
        ws['C1'].alignment = align_center

        ws.merge_cells('A2:L2')
        ws['A2'] = f"BOLETIM TÉCNICO DE SONDAGEM GEOLÓGICA - FURO {furo_id}"
        ws['A2'].font = font_sub
        ws['A2'].fill = fill_banner
        ws['A2'].alignment = align_center

        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 18

        # 2. Bloco de Cabeçalho do Projeto
        ws.merge_cells('A4:L4')
        ws['A4'] = "1. DADOS DE GESTÃO E LOCALIZAÇÃO"
        ws['A4'].font = font_sec
        ws['A4'].fill = fill_sec
        ws['A4'].alignment = align_left

        dados_header = [
            [("Projeto:", projeto), ("Coordenador:", coordenador), ("UTM (E):", utm_e), ("Latitude:", lat_furo)],
            [("ID Furo:", furo_id), ("Supervisor:", supervisor), ("UTM (N):", utm_n), ("Longitude:", lon_furo)],
            [("Diâmetro:", diametro), ("Geólogo Resp.:", geologo), ("Cota Z (m):", cota_z), ("Início:", str(data_inicio))],
            [("Inclin./Az.:", f"{inclinacao}° / {azimute}°"), ("Sondador:", sondador), ("Datum:", datum), ("Término:", str(data_fim))]
        ]

        curr_row = 5
        for row in dados_header:
            col_pairs = [(1,2,3), (4,5,6), (7,8,9), (10,11,12)]
            for idx, (lbl, val) in enumerate(row):
                c_lbl, c_val_start, c_val_end = col_pairs[idx]
                ws.cell(row=curr_row, column=c_lbl, value=lbl).font = Font(name='Calibri', size=10, bold=True)
                ws.cell(row=curr_row, column=c_lbl).alignment = align_left
                
                if c_val_start != c_val_end:
                    ws.merge_cells(start_row=curr_row, start_column=c_val_start, end_row=curr_row, end_column=c_val_end)
                cell_v = ws.cell(row=curr_row, column=c_val_start, value=val)
                cell_v.font = font_body
                cell_v.alignment = align_left
            curr_row += 1

        # 3. Tabela de Manobras
        curr_row += 1
        ws.merge_cells(f'A{curr_row}:L{curr_row}')
        ws[f'A{curr_row}'] = "2. REGISTRO DE MANOBRAS E PARÂMETROS GEOTÉCNICOS"
        ws[f'A{curr_row}'].font = font_sec
        ws[f'A{curr_row}'].fill = fill_sec
        ws[f'A{curr_row}'].alignment = align_left

        curr_row += 1
        df_excel = df_manobras.drop(columns=['Foto'])
        
        # Cabeçalhos
        for c_idx, col_name in enumerate(df_excel.columns, 1):
            cell = ws.cell(row=curr_row, column=c_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border
        ws.row_dimensions[curr_row].height = 22

        # Linhas de Dados
        header_row_idx = curr_row
        curr_row += 1
        for r_idx, row in df_excel.iterrows():
            row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
            for c_idx, val in enumerate(row.values, 1):
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                cell.fill = row_fill
                
                if isinstance(val, (int, float)):
                    cell.alignment = align_right
                    if "Rec (%)" in df_excel.columns[c_idx-1] or "RQD (%)" in df_excel.columns[c_idx-1]:
                        cell.number_format = '0.0'
                    elif "m" in df_excel.columns[c_idx-1]:
                        cell.number_format = '0.00'
                else:
                    cell.alignment = align_center if c_idx == 1 else align_left
            curr_row += 1

        # Totais
        ws.cell(row=curr_row, column=1, value="Total / Média").font = font_total
        ws.cell(row=curr_row, column=1).alignment = align_center
        ws.cell(row=curr_row, column=1).fill = fill_total
        ws.cell(row=curr_row, column=1).border = double_bottom

        for c_idx in range(2, len(df_excel.columns) + 1):
            col_name = df_excel.columns[c_idx-1]
            cell = ws.cell(row=curr_row, column=c_idx)
            cell.font = font_total
            cell.fill = fill_total
            cell.border = double_bottom
            cell.alignment = align_right

            start_letter = get_column_letter(c_idx)
            start_cell = f"{start_letter}{header_row_idx + 1}"
            end_cell = f"{start_letter}{curr_row - 1}"

            if col_name in ['Avanço (m)', 'Rec. (m)', 'RQD (m)']:
                cell.value = f"=SUM({start_cell}:{end_cell})"
                cell.number_format = '0.00'
            elif col_name in ['Rec (%)', 'RQD (%)']:
                cell.value = f"=AVERAGE({start_cell}:{end_cell})"
                cell.number_format = '0.0%'
            else:
                cell.value = "-"
                cell.alignment = align_center

        # 4. Campo de Assinatura no Excel
        curr_row += 3
        ws.merge_cells(f'A{curr_row}:E{curr_row}')
        ws[f'A{curr_row}'] = "3. VALIDAÇÃO E ASSINATURA TÉCNICA"
        ws[f'A{curr_row}'].font = font_sec
        ws[f'A{curr_row}'].fill = fill_sec

        curr_row += 2
        if canvas_result.image_data is not None:
            img_ass_pil = Image.fromarray(canvas_result.image_data.astype('uint8'))
            ass_excel_buf = io.BytesIO()
            img_ass_pil.save(ass_excel_buf, format='PNG')
            ass_excel_buf.seek(0)
            
            xl_ass = OpenpyxlImage(ass_excel_buf)
            xl_ass.width = 180
            xl_ass.height = 55
            ws.add_image(xl_ass, f'A{curr_row}')

        ws.cell(row=curr_row+3, column=1, value="_________________________________________").font = font_body
        ws.cell(row=curr_row+4, column=1, value=f"{geologo} - Geólogo Responsável").font = Font(name='Calibri', size=10, bold=True)

        # Ajuste de largura das colunas
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.coordinate in ws.merged_cells:
                    continue
                if cell.value:
                    val_str = str(cell.value)
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        wb.save(buffer_xls)
        st.download_button(
            label="📊 Baixar Planilha Excel (.xlsx)",
            data=buffer_xls.getvalue(),
            file_name=f"Boletim_{furo_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # --- EXPORTAÇÃO PDF ABNT ---
    with col_exp2:
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
                if self._pageNumber > 1:
                    self.setFont("Times-Roman", 9)
                    text = f"Página {self._pageNumber} de {page_count}"
                    self.drawRightString(19.0 * cm, 28.0 * cm, text)

        pdf_buf = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buf, pagesize=portrait(A4),
            leftMargin=3.0*cm, rightMargin=2.0*cm,
            topMargin=3.0*cm, bottomMargin=2.0*cm
        )
        elements = []
        styles = getSampleStyleSheet()

        # Estilos ABNT
        abnt_titulo_doc = ParagraphStyle('ABNTTituloDoc', parent=styles['Heading1'], fontName='Times-Bold', fontSize=13, leading=15, alignment=1, spaceAfter=4)
        abnt_sub_doc = ParagraphStyle('ABNTSubDoc', parent=styles['Normal'], fontName='Times-Roman', fontSize=10, leading=12, alignment=1, spaceAfter=15)
        abnt_sec = ParagraphStyle('ABNTSec', parent=styles['Heading2'], fontName='Times-Bold', fontSize=11, leading=13, spaceBefore=10, spaceAfter=6)
        abnt_text = ParagraphStyle('ABNTText', parent=styles['Normal'], fontName='Times-Roman', fontSize=8.5, leading=11)
        abnt_text_bold = ParagraphStyle('ABNTTextBold', parent=styles['Normal'], fontName='Times-Bold', fontSize=8.5, leading=11)
        abnt_th = ParagraphStyle('ABNTTH', parent=styles['Normal'], fontName='Times-Bold', fontSize=8, leading=9, alignment=1)
        abnt_td = ParagraphStyle('ABNTTD', parent=styles['Normal'], fontName='Times-Roman', fontSize=8, leading=10, alignment=1)
        abnt_caption = ParagraphStyle('ABNTCaption', parent=styles['Italic'], fontName='Times-Italic', fontSize=8.5, leading=10, alignment=0, spaceAfter=4)
        abnt_fonte = ParagraphStyle('ABNTFonte', parent=styles['Italic'], fontName='Times-Roman', fontSize=7.5, leading=9, alignment=0, spaceBefore=3, spaceAfter=8)

        # 1. Cabeçalho com Logo
        if img_logo_pil:
            img_logo_pdf_buf = io.BytesIO()
            img_logo_pil.save(img_logo_pdf_buf, format='PNG')
            img_logo_pdf_buf.seek(0)
            rl_logo = RLImage(img_logo_pdf_buf, width=4.0*cm, height=1.5*cm)
            
            header_table = Table([[rl_logo, Paragraph(f"<b>{empresa.upper()}</b><br/>RELATÓRIO TÉCNICO DE SONDAGEM GEOLÓGICA", abnt_titulo_doc)]], colWidths=[4.5*cm, 11.5*cm])
            header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,0), 'LEFT')]))
            elements.append(header_table)
            elements.append(Spacer(1, 10))
        else:
            elements.append(Paragraph(f"<b>{empresa.upper()}</b>", abnt_titulo_doc))
            elements.append(Paragraph(f"RELATÓRIO TÉCNICO DE SONDAGEM GEOLÓGICA — FURO <b>{furo_id}</b>", abnt_sub_doc))

        # 2. Dados Furo
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
        
        t_furo = Table(dados_furo_table, colWidths=[3.0*cm, 5.0*cm, 3.0*cm, 5.0*cm])
        t_furo.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#475569')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC'))
        ]))
        elements.append(t_furo)
        elements.append(Spacer(1, 6))

        # 3. Satélite
        elements.append(Paragraph("<b>1.1 Localização de Satélite do Furo</b>", abnt_sec))
        try:
            url_satelite = f"https://staticmap.openstreetmap.de/staticmap.php?center={lat_furo},{lon_furo}&zoom=15&size=600x180&maptype=mapnik&markers={lat_furo},{lon_furo},red-pushpin"
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            res = requests.get(url_satelite, headers=headers, timeout=5)
            if res.status_code == 200:
                img_sat_buf = io.BytesIO(res.content)
                img_sat_pdf = RLImage(img_sat_buf, width=16.0*cm, height=4.2*cm)
                elements.append(Paragraph(f"<b>Figura 1</b> – Imagem de satélite e localização geográfica do Furo {furo_id}.", abnt_caption))
                elements.append(img_sat_pdf)
                elements.append(Paragraph("Fonte: Adaptado de OpenStreetMap (2026).", abnt_fonte))
            else:
                elements.append(Paragraph("<i>(Imagem de satélite indisponível no momento)</i>", abnt_text))
        except Exception:
            elements.append(Paragraph("<i>(Imagem de satélite indisponível no momento)</i>", abnt_text))

        elements.append(Spacer(1, 4))

        # 4. Tabela de Manobras
        elements.append(Paragraph("<b>2. REGISTRO DE MANOBRAS E DADOS GEOTÉCNICOS</b>", abnt_sec))
        elements.append(Paragraph("<b>Tabela 1</b> – Parâmetros geotécnicos e descrição litológica por manobra.", abnt_caption))

        table_manobras_data = [
            [Paragraph("<b>Mnb.</b>", abnt_th), Paragraph("<b>De (m)</b>", abnt_th), Paragraph("<b>Para (m)</b>", abnt_th), 
             Paragraph("<b>Avanço (m)</b>", abnt_th), Paragraph("<b>Rec. (m)</b>", abnt_th), Paragraph("<b>Rec. (%)</b>", abnt_th), 
             Paragraph("<b>RQD (%)</b>", abnt_th), Paragraph("<b>Litologia</b>", abnt_th), Paragraph("<b>Observações</b>", abnt_th)]
        ]

        for _, r in df_manobras.iterrows():
            table_manobras_data.append([
                Paragraph(str(int(r['Manobra'])), abnt_td), Paragraph(f"{r['De (m)']:.2f}", abnt_td),
                Paragraph(f"{r['Para (m)']:.2f}", abnt_td), Paragraph(f"{r['Avanço (m)']:.2f}", abnt_td),
                Paragraph(f"{r['Rec. (m)']:.2f}", abnt_td), Paragraph(f"{r['Rec (%)']:.1f}", abnt_td),
                Paragraph(f"{r['RQD (%)']:.1f}", abnt_td), Paragraph(str(r['Litologia']), abnt_text),
                Paragraph(str(r['Observações'] or '-'), abnt_text)
            ])

        table_manobras_data.append([
            Paragraph("<b>Total / Média</b>", abnt_th), Paragraph("-", abnt_td), Paragraph("-", abnt_td),
            Paragraph(f"<b>{df_manobras['Avanço (m)'].sum():.2f}</b>", abnt_th),
            Paragraph(f"<b>{df_manobras['Rec. (m)'].sum():.2f}</b>", abnt_th),
            Paragraph(f"<b>{df_manobras['Rec (%)'].mean():.1f}%</b>", abnt_th),
            Paragraph(f"<b>{df_manobras['RQD (%)'].mean():.1f}%</b>", abnt_th),
            Paragraph("-", abnt_td), Paragraph("-", abnt_td)
        ])

        t_manobras = Table(table_manobras_data, colWidths=[1.2*cm, 1.4*cm, 1.4*cm, 1.6*cm, 1.4*cm, 1.5*cm, 1.5*cm, 3.0*cm, 3.0*cm])
        t_manobras.setStyle(TableStyle([
            ('LINEABOVE', (0,0), (-1,0), 1.0, colors.black),
            ('LINEBELOW', (0,0), (-1,0), 1.0, colors.black),
            ('LINEBELOW', (0,-1), (-1,-1), 1.0, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#F1F5F9')),
        ]))
        
        elements.append(t_manobras)
        elements.append(Paragraph("Fonte: Dados de campo do projeto (2026).", abnt_fonte))

        elements.append(PageBreak())

        # 5. Perfil
        elements.append(Paragraph("<b>3. PERFIL STRATIGRÁFICO E CURVAS DE RECUPERAÇÃO/RQD</b>", abnt_sec))
        elements.append(Paragraph("<b>Figura 2</b> – Perfil geológico e variação dos parâmetros geotécnicos em profundidade.", abnt_caption))

        img_plt_buf = io.BytesIO()
        fig.savefig(img_plt_buf, format='png', dpi=200, bbox_inches='tight')
        img_plt_buf.seek(0)
        
        img_perfil_pdf = RLImage(img_plt_buf, width=16.0*cm, height=7.5*cm)
        elements.append(img_perfil_pdf)
        elements.append(Paragraph("Fonte: Elaborado pelos autores a partir de medições de campo.", abnt_fonte))

        # 6. Galeria de Fotos
        elements.append(Spacer(1, 6))
        elements.append(Paragraph("<b>4. REGISTRO FOTOGRÁFICO DOS TESTEMUNHOS DE SONDAGEM</b>", abnt_sec))
        
        fotos_list = [m['Foto'] for m in st.session_state['manobras'] if m['Foto'] is not None]
        if fotos_list:
            elements.append(Paragraph("<b>Figura 3</b> – Registros fotográficos das amostras obtidas.", abnt_caption))
            fotos_grid = []
            row_temp = []
            for idx, img_p in enumerate(fotos_list):
                img_b = io.BytesIO()
                img_copy = img_p.copy()
                img_copy.thumbnail((400, 300))
                img_copy.save(img_b, format='PNG')
                img_b.seek(0)
                
                rl_img = RLImage(img_b, width=3.8*cm, height=2.8*cm)
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
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            elements.append(t_fotos)
            elements.append(Paragraph("Fonte: Acervo fotográfico da amostragem (2026).", abnt_fonte))
        else:
            elements.append(Paragraph("<i>Nenhum registro fotográfico anexado a este furo.</i>", abnt_text))

        # 7. Assinatura PDF
        elements.append(Spacer(1, 10))
        elements_assinatura = [Paragraph("<b>5. ENCERRAMENTO E ASSINATURA TÉCNICA</b>", abnt_sec)]
        
        if canvas_result.image_data is not None:
            img_ass = Image.fromarray(canvas_result.image_data.astype('uint8'))
            ass_buf = io.BytesIO()
            img_ass.save(ass_buf, format='PNG')
            ass_buf.seek(0)
            rl_ass = RLImage(ass_buf, width=5.5*cm, height=1.6*cm)
        else:
            rl_ass = Paragraph("<br/><br/>___________________________________", abnt_text)

        sig_table_data = [[rl_ass, Paragraph(f"<br/><br/>___________________________________<br/><b>{geologo}</b><br/>Geólogo Responsável", abnt_text)]]
        t_sig = Table(sig_table_data, colWidths=[8.0*cm, 8.0*cm])
        t_sig.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        elements_assinatura.append(t_sig)
        elements.append(KeepTogether(elements_assinatura))

        # Gerar PDF
        doc.build(elements, canvasmaker=NumberedCanvas)
        
        st.download_button(
            label="📄 Baixar Boletim ABNT (.pdf)",
            data=pdf_buf.getvalue(),
            file_name=f"Boletim_ABNT_{furo_id}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
