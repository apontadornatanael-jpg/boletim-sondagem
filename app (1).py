import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
from datetime import datetime
from PIL import Image

import folium
from streamlit_folium import st_folium

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ReportLab para geração de PDF
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

# Componente de Assinatura
from streamlit_drawable_canvas import st_canvas

# Configuração da página
st.set_page_config(
    page_title="Boletim de Sondagem Mineral",
    page_icon="⛏️",
    layout="wide"
)

# Estilização CSS
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

CORES_LITOLOGIA = {
    'Solo / Cobertura': '#D2B48C',
    'Siltito / Argilito': '#A0522D',
    'Quartzito': '#FFF8DC',
    'Schisto / Filito': '#708090',
    'Gnaisse / Granito': '#E6E6FA',
    'Basalto / Diabásio': '#2F4F4F',
    'Minério de Ferro / BIF': '#8B0000',
    'Calcário / Dolomito': '#B0C4DE',
    'Outro': '#808080'
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
    st_folium(m, width="100%", height=350)

st.markdown("---")

# 2. REGISTRO DE MANOBRAS E FOTOS DO TESTEMUNHO
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
    litologia = st.selectbox("Litologia", list(CORES_LITOLOGIA.keys()))
with col_l2:
    alteracao = st.selectbox("Alteração", ['Solo / Inconsol.', 'Completamente Alterada', 'Muito Alterada', 'Moderadamente Alterada', 'Pouco Alterada', 'Rocha Sã'])
with col_l3:
    obs = st.text_input("Observações Geotécnicas", placeholder="Ex: RPT, Fraturado, veios de quartzo...")

# ANEXO DE FOTO / CÂMERA DO CELULAR
st.subheader("📷 Registro Fotográfico da Amostra / Caixa")
aba_cam, aba_up = st.tabs(["📸 Tirar Foto Agora (Câmera)", "📁 Carregar da Galeria"])

img_capturada = None
with aba_cam:
    foto_cam = st.camera_input("Tirar foto da caixa de testemunho")
    if foto_cam:
        img_capturada = Image.open(foto_cam)

with aba_up:
    foto_file = st.file_uploader("Selecione uma imagem", type=['jpg', 'jpeg', 'png'])
    if foto_file and not img_capturada:
        img_capturada = Image.open(foto_file)

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
        st.success("✅ Manobra e Foto registradas com sucesso!")
        st.rerun()

if btn_remover and st.session_state['manobras']:
    st.session_state['manobras'].pop()
    st.warning("🗑️ Última manobra removida.")
    st.rerun()

# EXIBIÇÃO DA GALERIA DAS FOTOS REGISTRADAS
if st.session_state['manobras']:
    st.markdown("#### 🖼️ Galeria de Fotos das Amostras Registradas")
    cols_galeria = st.columns(4)
    idx_col = 0
    for m_item in st.session_state['manobras']:
        if m_item['Foto'] is not None:
            with cols_galeria[idx_col % 4]:
                st.image(m_item['Foto'], caption=f"Furo {furo_id} | {m_item['De (m)']}m - {m_item['Para (m)']}m", use_container_width=True)
            idx_col += 1

st.markdown("---")

# 3. TABELA CONSOLIDADA E GERADOR DE RELATÓRIOS (EXCEL FORMATEDO E PDF)
st.header("3. Relatórios e Assinatura Digital")

if st.session_state['manobras']:
    df_manobras = pd.DataFrame(st.session_state['manobras'])
    
    st.dataframe(
        df_manobras.drop(columns=['Foto']).style.format({
            'De (m)': '{:.2f}', 'Para (m)': '{:.2f}', 'Avanço (m)': '{:.2f}',
            'Rec. (m)': '{:.2f}', 'Rec (%)': '{:.1f}%', 'RQD (m)': '{:.2f}', 'RQD (%)': '{:.1f}%'
        }),
        use_container_width=True, hide_index=True
    )

    # GERADOR DO GRÁFICO PARA O PDF
    def gerar_grafico_perfil():
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(8, 4), sharey=True)
        prof_max = df_manobras['Para (m)'].max()
        ax1.set_ylim(prof_max, 0)

        for _, row in df_manobras.iterrows():
            de_m, para_m, lito = row['De (m)'], row['Para (m)'], row['Litologia']
            cor = CORES_LITOLOGIA.get(lito, '#808080')
            ax1.add_patch(mpatches.Rectangle((0, de_m), 1, para_m - de_m, facecolor=cor, edgecolor='black'))

        ax1.set_xlim(0, 1)
        ax1.set_title("Litologia")
        ax1.set_ylabel("Profundidade (m)")
        ax1.get_xaxis().set_visible(False)

        for _, row in df_manobras.iterrows():
            ax2.barh(y=row['De (m)'] + row['Avanço (m)']/2, width=row['RQD (%)'], height=row['Avanço (m)'], color='#2ecc71', edgecolor='black')
        ax2.set_xlim(0, 105)
        ax2.set_title("RQD (%)")
        
        plt.tight_layout()
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', dpi=150)
        plt.close(fig)
        img_buf.seek(0)
        return img_buf

    st.markdown("### ✍️ Assinatura Digital do Responsável")
    st.write("Assine no quadro abaixo usando a tela do celular ou mouse:")
    
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#F8FAFC",
        height=120,
        width=400,
        drawing_mode="freedraw",
        key="canvas_assinatura",
    )

    col_exp1, col_exp2 = st.columns(2)

    # EXPORTAÇÃO EXCEL FORMATEDO PROFISSIONALMENTE
    with col_exp1:
        buffer_xls = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Boletim de Sondagem"
        ws.views.sheetView[0].showGridLines = True

        font_titulo = Font(name='Calibri', size=16, bold=True, color='1F497D')
        font_subtitulo = Font(name='Calibri', size=11, bold=True, color='595959')
        font_secao = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        font_cabecalho_tab = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        font_dados = Font(name='Calibri', size=10)
        font_bold = Font(name='Calibri', size=10, bold=True)
        
        fill_azul_escuro = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        fill_cinza_secao = PatternFill(start_color='595959', end_color='595959', fill_type='solid')
        fill_zebrado = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
        fill_logo_box = PatternFill(start_color='E9EEF4', end_color='E9EEF4', fill_type='solid')
        
        border_fina = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                             top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        align_center = Alignment(horizontal='center', vertical='center')
        align_left = Alignment(horizontal='left', vertical='center')
        align_right = Alignment(horizontal='right', vertical='center')

        ws.merge_cells('A1:C3')
        ws['A1'] = " Espaço Reservado\n para LOGO MARCA\n da Empresa"
        ws['A1'].font = Font(name='Calibri', size=9, italic=True, color='595959')
        ws['A1'].fill = fill_logo_box
        ws['A1'].alignment = align_center

        ws.merge_cells('D1:L2')
        ws['D1'] = "BOLETIM TÉCNICO DE SONDAGEM GEOLÓGICA"
        ws['D1'].font = font_titulo
        ws['D1'].alignment = align_center

        ws.merge_cells('D3:L3')
        ws['D3'] = f"EMPRESA: {empresa.upper()} | PROJETO: {projeto.upper()}"
        ws['D3'].font = font_subtitulo
        ws['D3'].alignment = align_center

        ws.merge_cells('A5:L5')
        ws['A5'] = " 1. DADOS DE GESTÃO, EQUIPE E LOCALIZAÇÃO DO FURO"
        ws['A5'].font = font_secao
        ws['A5'].fill = fill_cinza_secao
        ws['A5'].alignment = align_left

        painel_dados = [
            [("ID do Furo:", furo_id), ("Coordenador:", coordenador), ("UTM (E):", utm_e), ("Inclinação:", f"{inclinacao}°")],
            [("Diâmetro:", diametro), ("Supervisor:", supervisor), ("UTM (N):", utm_n), ("Azimute:", f"{azimute}°")],
            [("Data Início:", str(data_inicio)), ("Geólogo Resp.:", geologo), ("Cota Z (m):", cota_z), ("Datum:", datum)],
            [("Data Fim:", str(data_fim)), ("Sondador:", sondador), ("Prof. Total (m):", df_manobras['Para (m)'].max()), ("GPS (Lat/Lon):", f"{lat_furo:.5f}, {lon_furo:.5f}")]
        ]

        for row_offset, linha in enumerate(painel_dados, start=6):
            cols_pos = [(1, 2, 3), (4, 5, 6), (7, 8, 9), (10, 11, 12)]
            for idx, (label, val) in enumerate(linha):
                c_lbl, c_val1, c_val2 = cols_pos[idx]
                ws.cell(row=row_offset, column=c_lbl, value=label).font = font_bold
                ws.cell(row=row_offset, column=c_lbl).alignment = align_right
                
                ws.merge_cells(start_row=row_offset, start_column=c_val1, end_row=row_offset, end_column=c_val2)
                cell_val = ws.cell(row=row_offset, column=c_val1, value=val)
                cell_val.font = font_dados
                cell_val.alignment = align_left
                cell_val.border = border_fina

        ws.merge_cells('A11:L11')
        ws['A11'] = " 2. REGISTRO DE MANOBRAS E GEOTECNIA"
        ws['A11'].font = font_secao
        ws['A11'].fill = fill_azul_escuro
        ws['A11'].alignment = align_left

        df_excel = df_manobras.drop(columns=['Foto'])
        headers = list(df_excel.columns)
        for col_idx, col_name in enumerate(headers, 1):
            cell = ws.cell(row=12, column=col_idx, value=col_name)
            cell.font = font_cabecalho_tab
            cell.fill = fill_azul_escuro
            cell.alignment = align_center

        start_row = 13
        for r_idx, row in df_excel.iterrows():
            curr_row = start_row + r_idx
            ws.append(list(row.values))
            fill_curr = fill_zebrado if curr_row % 2 == 0 else PatternFill(fill_type=None)
            
            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=curr_row, column=c_idx)
                cell.font = font_dados
                cell.border = border_fina
                cell.fill = fill_curr
                cell.alignment = align_center
                
                if headers[c_idx-1] in ['De (m)', 'Para (m)', 'Avanço (m)', 'Rec. (m)', 'RQD (m)']:
                    cell.number_format = '0.00'
                elif headers[c_idx-1] in ['Rec (%)', 'RQD (%)']:
                    cell.number_format = '0.0"%"'

        tot_row = start_row + len(df_excel)
        ws.cell(row=tot_row + 1, column=1, value="TOTAL / MÉDIA").font = font_bold
        ws.cell(row=tot_row + 1, column=4, value=df_excel['Avanço (m)'].sum()).font = font_bold
        ws.cell(row=tot_row + 1, column=4).number_format = '0.00'
        ws.cell(row=tot_row + 1, column=5, value=df_excel['Rec. (m)'].sum()).font = font_bold
        ws.cell(row=tot_row + 1, column=5).number_format = '0.00'
        ws.cell(row=tot_row + 1, column=6, value=df_excel['Rec (%)'].mean()).font = font_bold
        ws.cell(row=tot_row + 1, column=6).number_format = '0.0"%"'
        ws.cell(row=tot_row + 1, column=8, value=df_excel['RQD (%)'].mean()).font = font_bold
        ws.cell(row=tot_row + 1, column=8).number_format = '0.0"%"'

        ass_row = tot_row + 5
        ws.merge_cells(start_row=ass_row, start_column=2, end_row=ass_row, end_column=5)
        ws.cell(row=ass_row, column=2, value="________________________________________").alignment = align_center
        ws.merge_cells(start_row=ass_row+1, start_column=2, end_row=ass_row+1, end_column=5)
        ws.cell(row=ass_row+1, column=2, value=f"Geólogo Resp.: {geologo}").font = font_bold
        ws.cell(row=ass_row+1, column=2).alignment = align_center

        ws.merge_cells(start_row=ass_row, start_column=8, end_row=ass_row, end_column=11)
        ws.cell(row=ass_row, column=8, value="________________________________________").alignment = align_center
        ws.merge_cells(start_row=ass_row+1, start_column=8, end_row=ass_row+1, end_column=11)
        ws.cell(row=ass_row+1, column=8, value=f"Supervisor/Coordenador: {supervisor}").font = font_bold
        ws.cell(row=ass_row+1, column=8).alignment = align_center

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

        wb.save(buffer_xls)
        
        st.download_button(
            label="📄 Baixar Boletim Oficial (.xlsx)",
            data=buffer_xls.getvalue(),
            file_name=f"Boletim_Oficial_{furo_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # EXPORTAÇÃO PDF OFICIAL
    with col_exp2:
        if st.button("📄 Gerar Relatório PDF com Fotos e Assinatura", use_container_width=True):
            pdf_buf = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=portrait(A4), rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
            elements = []
            styles = getSampleStyleSheet()

            # Título PDF
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0284C7'), alignment=1)
            elements.append(Paragraph(f"<b>BOLETIM DE SONDAGEM MINERAL - {furo_id}</b>", title_style))
            elements.append(Spacer(1, 10))

            # Tabela de Dados Principais
            dados_cab = [
                [f"Empresa: {empresa}", f"Projeto: {projeto}"],
                [f"Geólogo: {geologo}", f"Sondador: {sondador}"],
                [f"UTM (E/N): {utm_e} / {utm_n}", f"Cota Z: {cota_z}m | Incl: {inclinacao}°"]
            ]
            t_cab = Table(dados_cab, colWidths=[9*cm, 9*cm])
            t_cab.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('FONTSIZE', (0,0), (-1,-1), 9),
            ]))
            elements.append(t_cab)
            elements.append(Spacer(1, 15))

            # Tabela Geotécnica
            head_man = [["De", "Para", "Av.", "Rec%", "RQD%", "Litologia"]]
            for _, r in df_manobras.iterrows():
                head_man.append([f"{r['De (m)']:.2f}", f"{r['Para (m)']:.2f}", f"{r['Avanço (m)']:.2f}", f"{r['Rec (%)']:.1f}%", f"{r['RQD (%)']:.1f}%", str(r['Litologia'])])
            
            t_man = Table(head_man, colWidths=[2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 5.5*cm])
            t_man.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284C7')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94A3B8')),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
            ]))
            elements.append(t_man)
            elements.append(Spacer(1, 15))

            # Gráfico do Perfil
            img_grafico = gerar_grafico_perfil()
            elements.append(RLImage(img_grafico, width=16*cm, height=8*cm))
            elements.append(Spacer(1, 15))

            # Anexo de Fotos se existirem
            fotos_para_pdf = [m for m in st.session_state['manobras'] if m['Foto'] is not None]
            if fotos_para_pdf:
                elements.append(PageBreak())
                elements.append(Paragraph("<b>REGISTRO FOTOGRÁFICO DAS AMOSTRAS</b>", title_style))
                elements.append(Spacer(1, 10))
                
                foto_rows = []
                temp_row = []
                for item in fotos_para_pdf:
                    im = item['Foto']
                    im_bytes = io.BytesIO()
                    im.save(im_bytes, format='JPEG')
                    im_bytes.seek(0)
                    
                    cell_content = [
                        RLImage(im_bytes, width=7.5*cm, height=5*cm),
                        Paragraph(f"<font size=8>Manobra: {item['De (m)']}m a {item['Para (m)']}m</font>", styles['Normal'])
                    ]
                    temp_row.append(cell_content)
                    
                    if len(temp_row) == 2:
                        foto_rows.append(temp_row)
                        temp_row = []
                if temp_row:
                    foto_rows.append(temp_row)
                
                t_fotos = Table(foto_rows, colWidths=[9*cm, 9*cm])
                elements.append(t_fotos)
                elements.append(Spacer(1, 15))

            # Assinatura Digital no PDF
            if canvas_result.image_data is not None:
                img_ass = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                ass_bytes = io.BytesIO()
                img_ass.save(ass_bytes, format='PNG')
                ass_bytes.seek(0)
                
                elements.append(Spacer(1, 10))
                elements.append(Paragraph(f"<b>Responsável Técnico: {geologo}</b>", styles['Normal']))
                elements.append(RLImage(ass_bytes, width=6*cm, height=2*cm))

            doc.build(elements)
            
            st.download_button(
                label="📥 Baixar PDF Concluído",
                data=pdf_buf.getvalue(),
                file_name=f"Relatorio_Sondagem_{furo_id}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
else:
    st.info("Nenhuma manobra cadastrada até o momento.")
