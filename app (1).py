import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Boletim de Sondagem Mineral",
    page_icon="⛏️",
    layout="wide"
)

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

# 1. DADOS DE GESTÃO, EQUIPE E LOCALIZAÇÃO
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

with st.expander("🌐 Dados Geográficos e Parâmetros do Furo (Opcional)", expanded=False):
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

# 2. INSERIR NOVA MANOBRA
st.header("2. Registro de Manobras de Sondagem")

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
    alteracao = st.selectbox("Alteração", [
        'Solo / Inconsol.', 'Completamente Alterada', 'Muito Alterada', 
        'Moderadamente Alterada', 'Pouco Alterada', 'Rocha Sã'
    ])
with col_l3:
    obs = st.text_input("Observações Geotécnicas", placeholder="Ex: RPT, Fraturado, veios de quartzo...")

col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    btn_adicionar = st.button("➕ Adicionar Manobra", type="primary", use_container_width=True)
with col_btn2:
    btn_remover = st.button("🗑️ Remover Última", use_container_width=True)

if btn_adicionar:
    avanco = round(para - de, 2)
    if avanco <= 0:
        st.error("⚠️ ERRO: O valor 'Para' deve ser maior que 'De'!")
    elif rec > avanco:
        st.error(f"⚠️ ATENÇÃO: A recuperação ({rec:.2f}m) NÃO pode ser maior que o avanço ({avanco:.2f}m)!")
    elif rqd > rec:
        st.error(f"⚠️ ATENÇÃO: O RQD ({rqd:.2f}m) NÃO pode ser maior que a recuperação ({rec:.2f}m)!")
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
            'Litologia': litologia, 'Alteração': alteracao, 'Observações': obs
        })
        st.success("✅ Manobra adicionada!")
        st.rerun()

if btn_remover and st.session_state['manobras']:
    st.session_state['manobras'].pop()
    st.warning("🗑️ Última manobra removida.")
    st.rerun()

st.markdown("---")

# 3. REGISTRO CONSOLIDADO & EXPORTAÇÃO EXCEL AMPLIADA
st.header("3. Tabela Consolidada e Exportação Oficial")

if st.session_state['manobras']:
    df_manobras = pd.DataFrame(st.session_state['manobras'])
    
    st.dataframe(
        df_manobras.style.format({
            'De (m)': '{:.2f}', 'Para (m)': '{:.2f}', 'Avanço (m)': '{:.2f}',
            'Rec. (m)': '{:.2f}', 'Rec (%)': '{:.1f}%', 'RQD (m)': '{:.2f}', 'RQD (%)': '{:.1f}%'
        }),
        use_container_width=True, hide_index=True
    )
    
    buffer = io.BytesIO()
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
        [("Data Fim:", str(data_fim)), ("Sondador:", sondador), ("Prof. Total (m):", df_manobras['Para (m)'].max()), ("-", "-")]
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

    headers = list(df_manobras.columns)
    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col_idx, value=col_name)
        cell.font = font_cabecalho_tab
        cell.fill = fill_azul_escuro
        cell.alignment = align_center

    start_row = 13
    for r_idx, row in df_manobras.iterrows():
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

    tot_row = start_row + len(df_manobras) + 1
    ws.cell(row=tot_row, column=1, value="TOTAL / MÉDIA").font = font_bold
    ws.cell(row=tot_row, column=4, value=df_manobras['Avanço (m)'].sum()).font = font_bold
    ws.cell(row=tot_row, column=4).number_format = '0.00'
    ws.cell(row=tot_row, column=5, value=df_manobras['Rec. (m)'].sum()).font = font_bold
    ws.cell(row=tot_row, column=5).number_format = '0.00'
    ws.cell(row=tot_row, column=6, value=df_manobras['Rec (%)'].mean()).font = font_bold
    ws.cell(row=tot_row, column=6).number_format = '0.0"%"'
    ws.cell(row=tot_row, column=8, value=df_manobras['RQD (%)'].mean()).font = font_bold
    ws.cell(row=tot_row, column=8).number_format = '0.0"%"'

    ass_row = tot_row + 4
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

    wb.save(buffer)
    
    st.download_button(
        label="📄 Baixar Boletim de Sondagem Oficial (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Boletim_Oficial_{furo_id}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Nenhuma manobra cadastrada até o momento.")

st.markdown("---")

# 4. PERFIL VISUAL DE SONDAGEM
st.header("4. Perfil Visual de Sondagem")

if st.session_state['manobras']:
    df = pd.DataFrame(st.session_state['manobras'])
    
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(11, 6), sharey=True)
    fig.suptitle(f"Perfil de Sondagem - Furo: {furo_id} | Projeto: {projeto}", fontsize=13, fontweight='bold')
    
    prof_max = df['Para (m)'].max()
    ax1.set_ylim(prof_max, 0)

    for _, row in df.iterrows():
        de_m, para_m, lito = row['De (m)'], row['Para (m)'], row['Litologia']
        cor = CORES_LITOLOGIA.get(lito, '#808080')
        ax1.add_patch(mpatches.Rectangle((0, de_m), 1, para_m - de_m, facecolor=cor, edgecolor='black'))
        if (para_m - de_m) >= 0.3:
            ax1.text(0.5, de_m + (para_m - de_m)/2, lito, ha='center', va='center', fontsize=8, color='black', weight='bold')

    ax1.set_xlim(0, 1)
    ax1.set_title("Litologia", fontsize=11, fontweight='bold')
    ax1.set_ylabel("Profundidade (m)", fontsize=11)
    ax1.get_xaxis().set_visible(False)

    for _, row in df.iterrows():
        ax2.barh(y=row['De (m)'] + row['Avanço (m)']/2, width=row['Rec (%)'], height=row['Avanço (m)'], color='#3498db', edgecolor='black', alpha=0.8)
    ax2.set_xlim(0, 105)
    ax2.axvline(100, color='red', linestyle='--', linewidth=1)
    ax2.set_title("Recuperação (%)", fontsize=11, fontweight='bold')
    ax2.grid(True, linestyle=':', alpha=0.6)

    for _, row in df.iterrows():
        rqd_val = row['RQD (%)']
        cor_rqd = '#e74c3c' if rqd_val < 50 else ('#f1c40f' if rqd_val < 75 else '#2ecc71')
        ax3.barh(y=row['De (m)'] + row['Avanço (m)']/2, width=rqd_val, height=row['Avanço (m)'], color=cor_rqd, edgecolor='black', alpha=0.8)
    ax3.set_xlim(0, 105)
    ax3.axvline(50, color='orange', linestyle='--', linewidth=1)
    ax3.set_title("Qualidade RQD (%)", fontsize=11, fontweight='bold')
    ax3.grid(True, linestyle=':', alpha=0.6)

    plt.tight_layout()
    st.pyplot(fig)
