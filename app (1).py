import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import base64

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Boletim Diário de Campo - Sondagem / Obra",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Boletim Diário de Campo")
st.markdown("Preencha os dados operacionais abaixo para gerar a planilha automatizada em conformidade com os padrões técnicos.")

# ==========================================
# FORMULÁRIO DE ENTRADA DE DADOS
# ==========================================
with st.form("form_boletim"):
    st.subheader("📌 1. Identificação do Projeto e Sondagem")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        obra_nome = st.text_input("Nome da Obra / Cliente", value="Obra Residencial - Bloco A")
        furo_id = st.text_input("Identificação do Furo/Ponto", value="SP-01")
    with col2:
        operador_nome = st.text_input("Responsável Técnico / Operador", value="Natanael Souza")
        data_boletim = st.date_input("Data do Registro", datetime.today())
    with col3:
        localizacao = st.text_input("Localização / Estaca", value="Estaca 12 + 5,00m")
        condicao_tempo = st.selectbox("Condição Climática", ["BOM", "NUBLADO", "CHUVA LEVE", "CHUVA FORTE"])

    st.markdown("---")
    st.subheader("⏱️ 2. Horário de Trabalho e Jornais Operacionais")
    col_h1, col_h2, col_h3 = st.columns(3)
    
    with col_h1:
        hora_inicio = st.time_input("Horário de Início", value=datetime.strptime("07:00", "%H:%M").time())
    with col_h2:
        hora_fim = st.time_input("Horário de Término", value=datetime.strptime("17:00", "%H:%M").time())
    with col_h3:
        intervalo_min = st.number_input("Intervalo de Almoço/Descanso (minutos)", min_value=0, value=60, step=15)

    # Cálculo prévio no Python para exibição em tempo real na interface
    dt_inicio = datetime.combine(datetime.today(), hora_inicio)
    dt_fim = datetime.combine(datetime.today(), hora_fim)
    if dt_fim < dt_inicio:
        dt_fim += timedelta(days=1)
    
    minutos_totais = max(0, int((dt_fim - dt_inicio).total_seconds() / 60) - intervalo_min)
    h_formatada = f"{minutos_totais // 60:02d}:{minutos_totais % 60:02d}"
    
    st.info(f"⏱️ **Cálculo de Jornada Efetiva:** {h_formatada} hrs ({minutos_totais / 60:.2f} horas decimais)")

    st.markdown("---")
    st.subheader("📦 3. Consumo de Insumos e Materiais")
    
    # Tabela dinâmica de insumos dentro do formulário
    insumos_default = pd.DataFrame([
        {"Descrição do Insumo": "Óleo Diesel", "Quantidade": 15.0, "Unidade": "Litros"},
        {"Descrição do Insumo": "Cimento CP II", "Quantidade": 2.0, "Unidade": "Sacos"},
        {"Descrição do Insumo": "Água para Perfuração", "Quantidade": 500.0, "Unidade": "Litros"},
    ])
    
    df_insumos_input = st.data_editor(
        insumos_default,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Descrição do Insumo": st.column_config.TextColumn("Descrição do Insumo", required=True),
            "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=0.0, format="%.2f"),
            "Unidade": st.column_config.SelectboxColumn("Unidade", options=["Litros", "Sacos", "Kg", "Unid", "m³", "Metros"])
        }
    )

    st.markdown("---")
    st.subheader("📝 4. Observações Técnicas")
    observacoes = st.text_area("Ocorrências, paradas ou observações de campo:", value="Execução realizada sem interrupções técnicas. Nível d'água encontrado a -2,50m.")

    btn_submeter = st.form_submit_button("⚙️ Gerar Planilha ABNT Automática", use_container_width=True)

# ==========================================
# PROCESSAMENTO E GERAÇÃO DA PLANILHA EXCEL
# ==========================================
if btn_submeter:
    wb = Workbook()
    
    # --------------------------------------
    # ABA 1: BOLETIM TÉCNICO (ABNT)
    # --------------------------------------
    ws = wb.active
    ws.title = "Boletim Diário"
    ws.views.sheetView[0].showGridLines = True

    # DEFINIÇÃO DE ESTILOS DE ACORDO COM ABNT
    font_titulo = Font(name="Arial", size=14, bold=True, color="1E293B")
    font_subtitulo = Font(name="Arial", size=10, italic=True, color="475569")
    font_secao = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Arial", size=10, bold=True, color="0F172A")
    font_dados = Font(name="Arial", size=10, bold=False, color="000000")
    font_bold = Font(name="Arial", size=10, bold=True, color="000000")

    fill_secao = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Azul Escuro Corporativo
    fill_header_tabela = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_destaque = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Amarelo Suave

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    border_celula = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_top_thick = Border(top=Side(border_style="medium", color="0F172A"))

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # CABEÇALHO ABNT
    ws.merge_cells("A1:E1")
    ws["A1"] = "RELATÓRIO DIÁRIO DE CAMPO - BOLETIM TÉCNICO"
    ws["A1"].font = font_titulo
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:E2")
    ws["A2"] = "Padronização ABNT NBR 14724 / Registro Operacional"
    ws["A2"].font = font_subtitulo
    ws["A2"].alignment = align_center

    # BLANCO SEPARADOR
    ws.row_dimensions[3].height = 10

    # SEÇÃO 1: IDENTIFICAÇÃO DO PROJETO
    ws.merge_cells("A4:E4")
    ws["A4"] = "1. IDENTIFICAÇÃO E DADOS GERAIS"
    ws["A4"].font = font_secao
    ws["A4"].fill = fill_secao
    ws["A4"].alignment = align_left

    dados_identificacao = [
        ("Obra / Cliente:", obra_nome, "Data do Registro:", data_boletim.strftime("%d/%m/%Y")),
        ("Ponto / Furo ID:", furo_id, "Clima:", condicao_tempo),
        ("Responsável Técnico:", operador_nome, "Localização / Estaca:", localizacao)
    ]

    curr_row = 5
    for row_data in dados_identificacao:
        ws.cell(row=curr_row, column=1, value=row_data[0]).font = font_header
        ws.cell(row=curr_row, column=2, value=row_data[1]).font = font_dados
        ws.cell(row=curr_row, column=4, value=row_data[2]).font = font_header
        ws.cell(row=curr_row, column=5, value=row_data[3]).font = font_dados
        curr_row += 1

    # SEÇÃO 2: HORÁRIO DE TRABALHO E CÁLCULO DE HORAS
    ws.merge_cells(f"A{curr_row}:E{curr_row}")
    ws[f"A{curr_row}"] = "2. JORNADA DE TRABALHO E HORAS EFETIVAS"
    ws[f"A{curr_row}"].font = font_secao
    ws[f"A{curr_row}"].fill = fill_secao
    ws[f"A{curr_row}"].alignment = align_left
    curr_row += 1

    # Inserção de Horários como valores datetime.time nativos
    c_inc = ws.cell(row=curr_row, column=1, value="Hora Início:")
    c_inc.font = font_header
    c_val_inc = ws.cell(row=curr_row, column=2, value=hora_inicio)
    c_val_inc.font = font_dados
    c_val_inc.alignment = align_center
    c_val_inc.number_format = 'hh:mm'

    c_fim = ws.cell(row=curr_row, column=4, value="Hora Término:")
    c_fim.font = font_header
    c_val_fim = ws.cell(row=curr_row, column=5, value=hora_fim)
    c_val_fim.font = font_dados
    c_val_fim.alignment = align_center
    c_val_fim.number_format = 'hh:mm'
    curr_row += 1

    ws.cell(row=curr_row, column=1, value="Intervalo (min):").font = font_header
    ws.cell(row=curr_row, column=2, value=intervalo_min).font = font_dados
    ws.cell(row=curr_row, column=2).alignment = align_center

    # FÓRMULA AUTOMÁTICA DO EXCEL PARA HORAS TRABALHADAS
    ws.cell(row=curr_row, column=4, value="Total Horas Trabalhadas:").font = font_header
    
    # Coordenadas das células no Excel para a fórmula
    cell_inc = f"B{curr_row-1}"
    cell_fim = f"E{curr_row-1}"
    cell_int = f"B{curr_row}"
    cell_res = f"E{curr_row}"

    # Fórmula Excel nativa automatizada com suporte à virada de dia
    formula_horas = f'=IF({cell_fim}<{cell_inc}, ({cell_fim}+1)-{cell_inc}-({cell_int}/1440), {cell_fim}-{cell_inc}-({cell_int}/1440))'
    
    ws[cell_res] = formula_horas
    ws[cell_res].font = font_bold
    ws[cell_res].fill = fill_destaque
    ws[cell_res].number_format = '[hh]:mm'
    ws[cell_res].alignment = align_center
    curr_row += 2

    # SEÇÃO 3: CONSUMO DE INSUMOS
    ws.merge_cells(f"A{curr_row}:E{curr_row}")
    ws[f"A{curr_row}"] = "3. REGISTRO DE INSUMOS E MATERIAIS"
    ws[f"A{curr_row}"].font = font_secao
    ws[f"A{curr_row}"].fill = fill_secao
    ws[f"A{curr_row}"].alignment = align_left
    curr_row += 1

    # Cabeçalho da Tabela de Insumos
    headers_tabela = ["Item", "Descrição do Insumo", "Quantidade", "Unidade", "Status/Aplicação"]
    for col_idx, header in enumerate(headers_tabela, 1):
        cell = ws.cell(row=curr_row, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header_tabela
        cell.alignment = align_center
        cell.border = border_celula
    curr_row += 1

    # Inserção das linhas de insumos
    for idx, row in df_insumos_input.iterrows():
        ws.cell(row=curr_row, column=1, value=idx + 1).alignment = align_center
        ws.cell(row=curr_row, column=2, value=row["Descrição do Insumo"]).alignment = align_left
        ws.cell(row=curr_row, column=3, value=row["Quantidade"]).alignment = align_right
        ws.cell(row=curr_row, column=3).number_format = '#,##0.00'
        ws.cell(row=curr_row, column=4, value=row["Unidade"]).alignment = align_center
        ws.cell(row=curr_row, column=5, value="Consumido").alignment = align_center

        for col_idx in range(1, 6):
            ws.cell(row=curr_row, column=col_idx).font = font_dados
            ws.cell(row=curr_row, column=col_idx).border = border_celula
        curr_row += 1

    curr_row += 1

    # SEÇÃO 4: OBSERVAÇÕES E ASSINATURA
    ws.merge_cells(f"A{curr_row}:E{curr_row}")
    ws[f"A{curr_row}"] = "4. OBSERVAÇÕES TÉCNICAS E OCORRÊNCIAS"
    ws[f"A{curr_row}"].font = font_secao
    ws[f"A{curr_row}"].fill = fill_secao
    ws[f"A{curr_row}"].alignment = align_left
    curr_row += 1

    ws.merge_cells(f"A{curr_row}:E{curr_row+2}")
    ws[f"A{curr_row}"] = observacoes
    ws[f"A{curr_row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws[f"A{curr_row}"].font = font_dados
    curr_row += 4

    # Campo de Assinatura Digital / Responsável
    ws.merge_cells(f"B{curr_row}:D{curr_row}")
    ws[f"B{curr_row}"].border = Border(top=Side(border_style="thin", color="000000"))
    ws[f"B{curr_row}"] = f"Assinatura do Responsável: {operador_nome}"
    ws[f"B{curr_row}"].font = font_subtitulo
    ws[f"B{curr_row}"].alignment = align_center

    # AJUSTE AUTOMÁTICO DE LARGURA DE COLUNAS
    merged_coords = [r.coord for r in ws.merged_cells.ranges]
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                if any(cell.coordinate in rng for rng in merged_coords):
                    continue
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    # GRAVAÇÃO EM BUFFER MEMÓRIA
    buffer_xls = io.BytesIO()
    wb.save(buffer_xls)
    buffer_xls.seek(0)

    # ==========================================
    # BOTÃO DE DOWNLOAD COMPATÍVEL COM IOS/ANDROID
    # ==========================================
    b64 = base64.b64encode(buffer_xls.getvalue()).decode()
    nome_arquivo = f"Boletim_Tecnico_{furo_id}_{data_boletim.strftime('%Y%m%d')}.xlsx"

    btn_excel_html = f'''
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
           download="{nome_arquivo}" 
           target="_blank" 
           style="text-decoration: none;">
            <button style="
                width: 100%;
                background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%);
                color: white;
                border: none;
                border-radius: 10px;
                padding: 14px 24px;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
                margin-top: 15px;
                margin-bottom: 10px;">
                📊 Baixar Planilha Excel ABNT (.xlsx)
            </button>
        </a>
    '''
    
    st.success("✅ Boletim Diário gerado com sucesso!")
    st.markdown(btn_excel_html, unsafe_allow_html=True)
    st.caption("📱 *Dica para iPhone/iOS: Após o download, toque em '‹ Safari' no canto superior esquerdo para retornar.*")
