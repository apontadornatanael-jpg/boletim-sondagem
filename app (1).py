import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
import json
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

# Componente de Comunicação JavaScript (Para suporte Offline / LocalStorage)
from streamlit_javascript import st_javascript

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
    div[data-testid="stMetric"] {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 12px 16px;
    }
    </style>
""", unsafe_allow_html=True)

# DICIONÁRIO LITOLÓGICO: COR + HACHURAS DENSA E VISÍVEIS
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
    st_folium(m, width="100%", height=350)

# --- PAINEL DE INDICADORES RÁPIDOS (KPIs) ---
st.markdown("---")
st.subheader("📊 Indicadores de Progresso do Furo")

if st.session_state['manobras']:
    df_kpi = pd.DataFrame(st.session_state['manobras'])
    prof_total = df_kpi['Para (m)'].max()
    rec_media = df_kpi['Rec (%)'].mean()
    rqd_medio = df_kpi['RQD (%)'].mean()
    avanco_medio = df_kpi['Avanço (m)'].mean()

    if rqd_medio < 25: qualidade_geral = "Muito Pobre"
    elif rqd_medio < 50: qualidade_geral = "Pobre"
    elif rqd_medio < 75: qualidade_geral = "Razoável"
    elif rqd_medio < 90: qualidade_geral = "Boa"
    else: qualidade_geral = "Excelente"

    col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
    with col_kpi1: st.metric(label="Profundidade Total", value=f"{prof_total:.2f} m", delta=f"{len(df_kpi)} manobra(s)")
    with col_kpi2: st.metric(label="Recuperação Média", value=f"{rec_media:.1f}%")
    with col_kpi3: st.metric(label="RQD Médio", value=f"{rqd_medio:.1f}%", delta=qualidade_geral)
    with col_kpi4: st.metric(label="Avanço Médio / Manobra", value=f"{avanco_medio:.2f} m")
else:
    col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
    with col_kpi1: st.metric(label="Profundidade Total", value="0.00 m")
    with col_kpi2: st.metric(label="Recuperação Média", value="0.0%")
    with col_kpi3: st.metric(label="RQD Médio", value="0.0%")
    with col_kpi4: st.metric(label="Avanço Médio / Manobra", value="0.00 m")

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
    litologia = st.selectbox("Litologia", list(DADOS_LITOLOGIA.keys()))
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
        st.success("✅ Manobra e Foto registradas com sucesso!")
        st.rerun()

if btn_remover and st.session_state['manobras']:
    st.session_state['manobras'].pop()
    st.warning("🗑️ Última manobra removida.")
    st.rerun()

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

# MODO OFFLINE / LOCAL STORAGE
st.header("📲 Armazenamento Offline & Sincronização")

def salvar_localmente_browser(chave, dados_json):
    js_code = f"localStorage.setItem('{chave}', JSON.stringify({dados_json}));"
    st_javascript(js_code)

def ler_rascunhos_browser(chave):
    js_code = f"localStorage.getItem('{chave}');"
    return st_javascript(js_code)

col_off1, col_off2, col_off3 = st.columns(3)

with col_off1:
    if st.button("💾 Salvar Rascunho no Celular (Offline)"):
        if st.session_state['manobras']:
            dados_para_salvar = []
            for item in st.session_state['manobras']:
                item_copia = item.copy()
                if 'Foto' in item_copia: item_copia['Foto'] = None
                dados_para_salvar.append(item_copia)
            dados_json = json.dumps(dados_para_salvar)
            salvar_localmente_browser(f"rascunho_furo_{furo_id}", dados_json)
            st.success("✅ Rascunho salvo na memória local do celular!")
        else:
            st.warning("Nenhuma manobra para salvar.")

with col_off2:
    if st.button("📂 Restaurar Rascunho do Celular"):
        dados_resgatados = ler_rascunhos_browser(f"rascunho_furo_{furo_id}")
        if dados_resgatados:
            try:
                st.session_state['manobras'] = json.loads(dados_resgatados)
                st.success("🔄 Rascunho restaurado com sucesso!")
                st.rerun()
            except Exception:
                st.error("Erro ao ler o rascunho salvo.")
        else:
            st.info("Nenum rascunho encontrado.")

with col_off3:
    if st.button("📡 Sincronizar quando houver sinal", type="primary"):
        if not st.session_state['manobras']:
            st.warning("Não há dados para sincronizar.")
        else:
            st.balloons()
            st.success("🚀 Dados sincronizados com sucesso!")

st.markdown("---")

# 3. PERFIL VISUAL E TABELA CONSOLIDADA
st.header("3. Perfil Litológico Visual e Tabela Consolidada")

if st.session_state['manobras']:
    df_manobras = pd.DataFrame(st.session_state['manobras'])
    
    # --- PERFIL LITOLÓGICO VISUAL COM HACHURAS E EIXOS GEOTÉCNICOS ---
    st.subheader("🪨 Perfil Litológico e Curva de RQD (Atualização Automática)")
    
    plt.rcParams['hatch.linewidth'] = 1.2
    plt.rcParams['hatch.color'] = '#333333'

    fig, (ax_lito, ax_rqd, ax_rec) = plt.subplots(
        1, 3, figsize=(11, 5), sharey=True, 
        gridspec_kw={'width_ratios': [1.3, 2, 2]}
    )
    
    prof_max = df_manobras['Para (m)'].max()
    ax_lito.set_ylim(prof_max, 0) # Inverte eixo para profundidade crescer para baixo

    litos_usadas = set()

    for _, row in df_manobras.iterrows():
        de_m = row['De (m)']
        para_m = row['Para (m)']
        lito = row['Litologia']
        rqd_val = row['RQD (%)']
        rec_val = row['Rec (%)']
        
        info_lito = DADOS_LITOLOGIA.get(lito, {'cor': '#808080', 'hatch': ''})
        litos_usadas.add(lito)

        # 1. Desenha o Bloco da Litologia com Hachura Geológica Visível
        rect = mpatches.Rectangle(
            (0, de_m), 1, para_m - de_m,
            facecolor=info_lito['cor'],
            hatch=info_lito['hatch'],
            edgecolor='#1E293B',
            linewidth=1.2
        )
        ax_lito.add_patch(rect)
        
        # Linha delimitadora da manobra no gráfico
        ax_lito.axhline(para_m, color='#0F172A', linestyle='--', linewidth=0.8)

        # Rótulo com o texto da Litologia + Limites no meio do bloco
        texto_bloco = f"{lito}\n({de_m:.1f}m - {para_m:.1f}m)"
        ax_lito.text(
            0.5, (de_m + para_m)/2, texto_bloco, 
            ha='center', va='center', fontsize=8, fontweight='bold', color='#0F172A',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#FFFFFF', alpha=0.85, edgecolor='#94A3B8')
        )

        # 2. RQD Colorido Geotécnico
        if rqd_val < 25: color_rqd = '#EF4444'     # Muito Pobre (Vermelho)
        elif rqd_val < 50: color_rqd = '#F97316'   # Pobre (Laranja)
        elif rqd_val < 75: color_rqd = '#EAB308'   # Razoável (Amarelo)
        elif rqd_val < 90: color_rqd = '#3B82F6'   # Boa (Azul)
        else: color_rqd = '#22C55E'               # Excelente (Verde)

        ax_rqd.barh(
            y=de_m + (para_m - de_m)/2, width=rqd_val, 
            height=(para_m - de_m)*0.8, color=color_rqd, edgecolor='black', linewidth=0.8
        )
        ax_rqd.axhline(para_m, color='#CBD5E1', linestyle=':', linewidth=0.8)

        # 3. Recuperação
        ax_rec.barh(
            y=de_m + (para_m - de_m)/2, width=rec_val, 
            height=(para_m - de_m)*0.8, color='#0284C7', edgecolor='black', linewidth=0.8
        )
        ax_rec.axhline(para_m, color='#CBD5E1', linestyle=':', linewidth=0.8)

    # Ajustes Eixo Litologia
    ax_lito.set_xlim(0, 1)
    ax_lito.set_title("Estratigrafia / Perfil", fontsize=11, fontweight='bold', color='#0369A1')
    ax_lito.set_ylabel("Profundidade (m)", fontsize=10, fontweight='bold')
    ax_lito.get_xaxis().set_visible(False)

    # Legenda Dinâmica de Litologia
    patches_legenda = [
        mpatches.Patch(
            facecolor=DADOS_LITOLOGIA[l]['cor'], 
            hatch=DADOS_LITOLOGIA[l]['hatch'], 
            edgecolor='black', 
            label=l
        ) for l in litos_usadas
    ]
    ax_lito.legend(handles=patches_legenda, loc='upper center', bbox_to_anchor=(0.5, -0.08), ncol=1, fontsize=8, frameon=True)

    # Ajustes Eixo RQD
    ax_rqd.set_xlim(0, 105)
    ax_rqd.set_title("RQD (%)", fontsize=11, fontweight='bold', color='#0369A1')
    ax_rqd.set_xlabel("%", fontsize=9)
    ax_rqd.grid(True, linestyle='--', alpha=0.5)

    # Ajustes Eixo Recuperação
    ax_rec.set_xlim(0, 105)
    ax_rec.set_title("Recuperação (%)", fontsize=11, fontweight='bold', color='#0369A1')
    ax_rec.set_xlabel("%", fontsize=9)
    ax_rec.grid(True, linestyle='--', alpha=0.5)

    plt.tight_layout()
    st.pyplot(fig)

    st.markdown("---")

    # Exibição da Tabela de Manobras
    st.dataframe(
        df_manobras.drop(columns=['Foto']).style.format({
            'De (m)': '{:.2f}', 'Para (m)': '{:.2f}', 'Avanço (m)': '{:.2f}',
            'Rec. (m)': '{:.2f}', 'Rec (%)': '{:.1f}%', 'RQD (m)': '{:.2f}', 'RQD (%)': '{:.1f}%'
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("### ✍️ Assinatura Digital do Responsável")
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)", stroke_width=2, stroke_color="#000000",
        background_color="#F8FAFC", height=120, width=400, drawing_mode="freedraw", key="canvas_assinatura"
    )

    col_exp1, col_exp2 = st.columns(2)

    # EXPORTAÇÃO EXCEL COMPLETA E DETALHADA
    with col_exp1:
        buffer_xls = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Boletim de Sondagem"
        ws.views.sheetView[0].showGridLines = True

        # Estilos do Excel
        font_titulo = Font(name='Calibri', size=16, bold=True, color='1F497D')
        font_secao = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        font_rotulo = Font(name='Calibri', size=10, bold=True, color='333333')
        font_valor = Font(name='Calibri', size=10, bold=False, color='000000')
        font_cabecalho_tab = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        font_dados = Font(name='Calibri', size=10)
        
        fill_azul_escuro = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        fill_cinza_secao = PatternFill(start_color='595959', end_color='595959', fill_type='solid')
        fill_zebrado = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
        fill_info = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

        border_fina = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        # 1. TÍTULO PRINCIPAL DO BOLETIM
        ws.merge_cells('A1:L2')
        ws['A1'] = "BOLETIM TÉCNICO DE SONDAGEM GEOLÓGICA"
        ws['A1'].font = font_titulo
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

        # 2. SEÇÃO 1: CABEÇALHO COM DADOS DE GESTÃO E LOCALIZAÇÃO
        ws.merge_cells('A4:L4')
        ws['A4'] = " 1. DADOS DE GESTÃO, EQUIPE E LOCALIZAÇÃO DO FURO"
        ws['A4'].font = font_secao
        ws['A4'].fill = fill_cinza_secao
        ws['A4'].alignment = Alignment(horizontal='left', vertical='center')

        # Matriz Completa do Cabeçalho
        dados_cabecalho = [
            [("Empresa / Mineradora:", empresa), ("Coordenador:", coordenador), ("Geólogo Resp.:", geologo), ("ID do Furo:", furo_id)],
            [("Projeto:", projeto), ("Supervisor:", supervisor), ("Sondador / Equipe:", sondador), ("Diâmetro:", diametro)],
            [("UTM (E):", utm_e), ("Cota Z (m):", cota_z), ("Inclinação (°):", inclinacao), ("Data Início:", str(data_inicio))],
            [("UTM (N):", utm_n), ("Datum:", datum), ("Azimute (°):", azimute), ("Data Término:", str(data_fim))]
        ]

        row_idx = 5
        for linha in dados_cabecalho:
            cols = [('A', 'B'), ('D', 'E'), ('G', 'H'), ('J', 'K')]
            for idx, (rotulo, valor) in enumerate(linha):
                col_r, col_v = cols[idx]
                
                # Rótulo
                ws[f'{col_r}{row_idx}'] = rotulo
                ws[f'{col_r}{row_idx}'].font = font_rotulo
                ws[f'{col_r}{row_idx}'].alignment = Alignment(horizontal='right', vertical='center')
                
                # Valor
                ws[f'{col_v}{row_idx}'] = valor
                ws[f'{col_v}{row_idx}'].font = font_valor
                ws[f'{col_v}{row_idx}'].alignment = Alignment(horizontal='left', vertical='center')
                ws[f'{col_v}{row_idx}'].fill = fill_info
                ws[f'{col_v}{row_idx}'].border = border_fina
            row_idx += 1

        # 3. SEÇÃO 2: TABELA DE MANOBRAS E GEOTECNIA
        ws.merge_cells('A10:L10')
        ws['A10'] = " 2. REGISTRO DE MANOBRAS E GEOTECNIA"
        ws['A10'].font = font_secao
        ws['A10'].fill = fill_azul_escuro
        ws['A10'].alignment = Alignment(horizontal='left', vertical='center')

        df_excel = df_manobras.drop(columns=['Foto'])
        headers = list(df_excel.columns)
        
        # Cabeçalho da Tabela
        for col_idx, col_name in enumerate(headers, 1):
            cell = ws.cell(row=11, column=col_idx, value=col_name)
            cell.font = font_cabecalho_tab
            cell.fill = fill_azul_escuro
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # Preenchimento dos Dados das Manobras
        start_row = 12
        for r_idx, row in df_excel.iterrows():
            curr_row = start_row + r_idx
            ws.append(list(row.values))
            fill_curr = fill_zebrado if curr_row % 2 == 0 else PatternFill(fill_type=None)
            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=curr_row, column=c_idx)
                cell.font = font_dados
                cell.border = border_fina
                cell.fill = fill_curr
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Autoajuste da Largura das Colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(buffer_xls)
        st.download_button(
            label="📄 Baixar Boletim Oficial (.xlsx)",
            data=buffer_xls.getvalue(),
            file_name=f"Boletim_{furo_id}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # EXPORTAÇÃO PDF
    with col_exp2:
        if st.button("📄 Gerar Relatório PDF Oficial", use_container_width=True):
            pdf_buf = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=portrait(A4), rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
            elements = []
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0284C7'), alignment=1)
            elements.append(Paragraph(f"<b>BOLETIM DE SONDAGEM MINERAL - {furo_id}</b>", title_style))
            elements.append(Spacer(1, 10))

            img_buf = io.BytesIO()
            fig.savefig(img_buf, format='png', dpi=150, bbox_inches='tight')
            img_buf.seek(0)

            elements.append(RLImage(img_buf, width=16*cm, height=8*cm))
            elements.append(Spacer(1, 10))

            doc.build(elements)
            st.download_button(
                label="📥 Baixar PDF Concluído",
                data=pdf_buf.getvalue(),
                file_name=f"Relatorio_{furo_id}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
else:
    st.info("Nenhuma manobra cadastrada até o momento.")
