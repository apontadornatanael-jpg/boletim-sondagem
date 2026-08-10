import io
from datetime import datetime

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Boletim de Sondagem Mineral",
    page_icon="⛏️",
    layout="wide",
)

CORES_LITOLOGIA = {
    "Solo / Cobertura": "#D2B48C",
    "Siltito / Argilito": "#A0522D",
    "Quartzito": "#FFF8DC",
    "Schisto / Filito": "#708090",
    "Gnaisse / Granito": "#E6E6FA",
    "Basalto / Diabásio": "#2F4F4F",
    "Minério de Ferro / BIF": "#8B0000",
    "Calcário / Dolomito": "#B0C4DE",
    "Outro": "#808080",
}

OPCOES_DIAMETRO = [
    "HQ (63.5mm)",
    "NQ (47.6mm)",
    "BQ (36.5mm)",
    "RC (Circ. Reversa)",
    "Outro",
]

OPCOES_ALTERACAO = [
    "Solo / Inconsol.",
    "Completamente Alterada",
    "Muito Alterada",
    "Moderadamente Alterada",
    "Pouco Alterada",
    "Rocha Sã",
]


# ============================================================
# ESTADO DA APLICAÇÃO
# ============================================================

def inicializar_estado():
    """Inicializa o estado da sessão.

    Nesta etapa ainda usamos session_state.
    Na próxima etapa ele poderá ser substituído por banco de dados.
    """
    if "manobras" not in st.session_state:
        st.session_state["manobras"] = []


# ============================================================
# FUNÇÕES DE CÁLCULO
# ============================================================

def classificar_rqd(percentual):
    """Classifica o RQD conforme a regra usada no app original."""
    if percentual < 25:
        return "Muito Pobre"
    if percentual < 50:
        return "Pobre"
    if percentual < 75:
        return "Razoável"
    if percentual < 90:
        return "Boa"
    return "Excelente"


def calcular_manobra(de, para, rec, rqd):
    """Valida e calcula os indicadores de uma manobra."""
    avanco = round(para - de, 2)

    if avanco <= 0:
        return None, "⚠️ ERRO: O valor 'Para' deve ser maior que 'De'!"

    if rec > avanco:
        return None, (
            f"⚠️ ATENÇÃO: A recuperação ({rec:.2f}m) "
            f"NÃO pode ser maior que o avanço ({avanco:.2f}m)!"
        )

    if rqd > rec:
        return None, (
            f"⚠️ ATENÇÃO: O RQD ({rqd:.2f}m) "
            f"NÃO pode ser maior que a recuperação ({rec:.2f}m)!"
        )

    pct_rec = min(100.0, round((rec / avanco) * 100, 1))
    pct_rqd = min(100.0, round((rqd / avanco) * 100, 1))

    return {
        "Avanço (m)": avanco,
        "Rec (%)": pct_rec,
        "RQD (%)": pct_rqd,
        "Qualidade RQD": classificar_rqd(pct_rqd),
    }, None


def proxima_profundidade():
    """Calcula automaticamente o próximo intervalo de manobra."""
    if not st.session_state["manobras"]:
        return 0.0, 1.5

    prox_de = st.session_state["manobras"][-1]["Para (m)"]
    prox_para = round(prox_de + 1.5, 2)
    return float(prox_de), float(prox_para)


# ============================================================
# INTERFACE - CABEÇALHO
# ============================================================

def renderizar_cabecalho():
    st.title("📋 Boletim Digital de Sondagem Mineral")
    st.markdown("---")


# ============================================================
# INTERFACE - DADOS DO PROJETO
# ============================================================

def renderizar_dados_projeto():
    st.header("1. Cabeçalho do Projeto & Equipe Técnica")

    col_g1, col_g2, col_g3, col_g4 = st.columns(4)

    with col_g1:
        empresa = st.text_input(
            "Empresa / Mineradora",
            value="Mineração Picuí S.A.",
        )
        projeto = st.text_input(
            "Nome do Projeto",
            value="Projeto Picuí",
        )

    with col_g2:
        coordenador = st.text_input(
            "Coordenador do Projeto",
            value="Eng. Carlos Andrade",
        )
        supervisor = st.text_input(
            "Supervisor de Campo",
            value="Téc. Roberto Lima",
        )

    with col_g3:
        geologo = st.text_input(
            "Geólogo Responsável",
            value="Geól. Mariana Costa",
        )
        sondador = st.text_input(
            "Sondador / Equipe",
            value="Natanael & Equipe",
        )

    with col_g4:
        furo_id = st.text_input(
            "ID do Furo",
            value="F-001",
        )
        diametro = st.selectbox(
            "Diâmetro",
            OPCOES_DIAMETRO,
        )

    with st.expander(
        "🌐 Dados Geográficos e Parâmetros do Furo (Opcional)",
        expanded=False,
    ):
        col_geo1, col_geo2, col_geo3, col_geo4 = st.columns(4)

        with col_geo1:
            utm_e = st.number_input(
                "Coordenada UTM (E)",
                value=250100.0,
                format="%.2f",
            )
            utm_n = st.number_input(
                "Coordenada UTM (N)",
                value=9245000.0,
                format="%.2f",
            )

        with col_geo2:
            cota_z = st.number_input(
                "Cota Z (m)",
                value=480.5,
                format="%.2f",
            )
            datum = st.text_input(
                "Datum",
                value="SIRGAS 2000",
            )

        with col_geo3:
            inclinacao = st.number_input(
                "Inclinação (°)",
                value=-90.0,
                format="%.1f",
            )
            azimute = st.number_input(
                "Azimute (°)",
                value=0.0,
                format="%.1f",
            )

        with col_geo4:
            data_inicio = st.date_input(
                "Data de Início",
                value=datetime.now(),
            )
            data_fim = st.date_input(
                "Data de Término",
                value=datetime.now(),
            )

    return {
        "empresa": empresa,
        "projeto": projeto,
        "coordenador": coordenador,
        "supervisor": supervisor,
        "geologo": geologo,
        "sondador": sondador,
        "furo_id": furo_id,
        "diametro": diametro,
        "utm_e": utm_e,
        "utm_n": utm_n,
        "cota_z": cota_z,
        "datum": datum,
        "inclinacao": inclinacao,
        "azimute": azimute,
        "data_inicio": data_inicio,
        "data_fim": data_fim,
    }


# ============================================================
# INTERFACE - MANOBRAS
# ============================================================

def adicionar_manobra(dados):
    resultado, erro = calcular_manobra(
        dados["de"],
        dados["para"],
        dados["rec"],
        dados["rqd"],
    )

    if erro:
        st.error(erro)
        return

    registro = {
        "Manobra": len(st.session_state["manobras"]) + 1,
        "De (m)": dados["de"],
        "Para (m)": dados["para"],
        "Avanço (m)": resultado["Avanço (m)"],
        "Rec. (m)": dados["rec"],
        "Rec (%)": resultado["Rec (%)"],
        "RQD (m)": dados["rqd"],
        "RQD (%)": resultado["RQD (%)"],
        "Qualidade RQD": resultado["Qualidade RQD"],
        "Litologia": dados["litologia"],
        "Alteração": dados["alteracao"],
        "Observações": dados["obs"],
    }

    st.session_state["manobras"].append(registro)
    st.success("✅ Manobra adicionada!")
    st.rerun()


def remover_ultima_manobra():
    if st.session_state["manobras"]:
        st.session_state["manobras"].pop()
        st.warning("🗑️ Última manobra removida.")
        st.rerun()


def renderizar_manobras():
    st.header("2. Registro de Manobras de Sondagem")

    prox_de, prox_para = proxima_profundidade()

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)

    with col_m1:
        de = st.number_input(
            "De (m)",
            value=prox_de,
            step=0.5,
            format="%.2f",
        )

    with col_m2:
        para = st.number_input(
            "Para (m)",
            value=prox_para,
            step=0.5,
            format="%.2f",
        )

    with col_m3:
        rec = st.number_input(
            "Rec. (m)",
            value=round(para - de, 2),
            step=0.1,
            format="%.2f",
        )

    with col_m4:
        rqd = st.number_input(
            "RQD (m)",
            value=round(max(para - de, 0) * 0.8, 2),
            step=0.1,
            format="%.2f",
        )

    col_l1, col_l2, col_l3 = st.columns(3)

    with col_l1:
        litologia = st.selectbox(
            "Litologia",
            list(CORES_LITOLOGIA.keys()),
        )

    with col_l2:
        alteracao = st.selectbox(
            "Alteração",
            OPCOES_ALTERACAO,
        )

    with col_l3:
        obs = st.text_input(
            "Observações Geotécnicas",
            placeholder="Ex: RPT, Fraturado, veios de quartzo...",
        )

    col_btn1, col_btn2 = st.columns([1, 4])

    with col_btn1:
        if st.button(
            "➕ Adicionar Manobra",
            type="primary",
            use_container_width=True,
        ):
            adicionar_manobra(
                {
                    "de": de,
                    "para": para,
                    "rec": rec,
                    "rqd": rqd,
                    "litologia": litologia,
                    "alteracao": alteracao,
                    "obs": obs,
                }
            )

    with col_btn2:
        if st.button(
            "🗑️ Remover Última",
            use_container_width=True,
        ):
            remover_ultima_manobra()


# ============================================================
# DADOS CONSOLIDADOS
# ============================================================

def obter_dataframe_manobras():
    if not st.session_state["manobras"]:
        return None
    return pd.DataFrame(st.session_state["manobras"])


def renderizar_tabela_consolidada(df):
    st.header("3. Tabela Consolidada e Exportação Oficial")

    st.dataframe(
        df.style.format(
            {
                "De (m)": "{:.2f}",
                "Para (m)": "{:.2f}",
                "Avanço (m)": "{:.2f}",
                "Rec. (m)": "{:.2f}",
                "Rec (%)": "{:.1f}%",
                "RQD (m)": "{:.2f}",
                "RQD (%)": "{:.1f}%",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# EXPORTAÇÃO EXCEL
# ============================================================

def configurar_estilos_excel():
    return {
        "font_titulo": Font(
            name="Calibri",
            size=16,
            bold=True,
            color="1F497D",
        ),
        "font_subtitulo": Font(
            name="Calibri",
            size=11,
            bold=True,
            color="595959",
        ),
        "font_secao": Font(
            name="Calibri",
            size=11,
            bold=True,
            color="FFFFFF",
        ),
        "font_cabecalho_tab": Font(
            name="Calibri",
            size=10,
            bold=True,
            color="FFFFFF",
        ),
        "font_dados": Font(
            name="Calibri",
            size=10,
        ),
        "font_bold": Font(
            name="Calibri",
            size=10,
            bold=True,
        ),
        "fill_azul_escuro": PatternFill(
            start_color="1F497D",
            end_color="1F497D",
            fill_type="solid",
        ),
        "fill_cinza_secao": PatternFill(
            start_color="595959",
            end_color="595959",
            fill_type="solid",
        ),
        "fill_zebrado": PatternFill(
            start_color="F2F5F9",
            end_color="F2F5F9",
            fill_type="solid",
        ),
        "fill_logo_box": PatternFill(
            start_color="E9EEF4",
            end_color="E9EEF4",
            fill_type="solid",
        ),
        "border_fina": Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        ),
        "align_center": Alignment(
            horizontal="center",
            vertical="center",
        ),
        "align_left": Alignment(
            horizontal="left",
            vertical="center",
        ),
        "align_right": Alignment(
            horizontal="right",
            vertical="center",
        ),
    }


def criar_excel_boletim(dados_projeto, df):
    """Cria o arquivo Excel do boletim.

    A função não depende da interface Streamlit, facilitando a futura
    reutilização em uma API ou aplicativo mobile.
    """
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Boletim de Sondagem"
    ws.views.sheetView[0].showGridLines = True

    estilos = configurar_estilos_excel()

    # Impressão
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    # Cabeçalho
    ws.merge_cells("A1:C3")
    ws["A1"] = " Espaço Reservado\n para LOGO MARCA\n da Empresa"
    ws["A1"].font = Font(
        name="Calibri",
        size=9,
        italic=True,
        color="595959",
    )
    ws["A1"].fill = estilos["fill_logo_box"]
    ws["A1"].alignment = estilos["align_center"]

    ws.merge_cells("D1:L2")
    ws["D1"] = "BOLETIM TÉCNICO DE SONDAGEM GEOLÓGICA"
    ws["D1"].font = estilos["font_titulo"]
    ws["D1"].alignment = estilos["align_center"]

    ws.merge_cells("D3:L3")
    ws["D3"] = (
        f"EMPRESA: {dados_projeto['empresa'].upper()} | "
        f"PROJETO: {dados_projeto['projeto'].upper()}"
    )
    ws["D3"].font = estilos["font_subtitulo"]
    ws["D3"].alignment = estilos["align_center"]

    # Seção 1
    ws.merge_cells("A5:L5")
    ws["A5"] = " 1. DADOS DE GESTÃO, EQUIPE E LOCALIZAÇÃO DO FURO"
    ws["A5"].font = estilos["font_secao"]
    ws["A5"].fill = estilos["fill_cinza_secao"]
    ws["A5"].alignment = estilos["align_left"]

    painel_dados = [
        [
            ("ID do Furo:", dados_projeto["furo_id"]),
            ("Coordenador:", dados_projeto["coordenador"]),
            ("UTM (E):", dados_projeto["utm_e"]),
            ("Inclinação:", f"{dados_projeto['inclinacao']}°"),
        ],
        [
            ("Diâmetro:", dados_projeto["diametro"]),
            ("Supervisor:", dados_projeto["supervisor"]),
            ("UTM (N):", dados_projeto["utm_n"]),
            ("Azimute:", f"{dados_projeto['azimute']}°"),
        ],
        [
            ("Data Início:", str(dados_projeto["data_inicio"])),
            ("Geólogo Resp.:", dados_projeto["geologo"]),
            ("Cota Z (m):", dados_projeto["cota_z"]),
            ("Datum:", dados_projeto["datum"]),
        ],
        [
            ("Data Fim:", str(dados_projeto["data_fim"])),
            ("Sondador:", dados_projeto["sondador"]),
            ("Prof. Total (m):", df["Para (m)"].max()),
            ("-", "-"),
        ],
    ]

    for row_offset, linha in enumerate(painel_dados, start=6):
        cols_pos = [
            (1, 2, 3),
            (4, 5, 6),
            (7, 8, 9),
            (10, 11, 12),
        ]

        for idx, (label, val) in enumerate(linha):
            c_lbl, c_val1, c_val2 = cols_pos[idx]

            ws.cell(
                row=row_offset,
                column=c_lbl,
                value=label,
            ).font = estilos["font_bold"]

            ws.cell(
                row=row_offset,
                column=c_lbl,
            ).alignment = estilos["align_right"]

            ws.merge_cells(
                start_row=row_offset,
                start_column=c_val1,
                end_row=row_offset,
                end_column=c_val2,
            )

            cell_val = ws.cell(
                row=row_offset,
                column=c_val1,
                value=val,
            )
            cell_val.font = estilos["font_dados"]
            cell_val.alignment = estilos["align_left"]
            cell_val.border = estilos["border_fina"]

    # Seção 2
    ws.merge_cells("A11:L11")
    ws["A11"] = " 2. REGISTRO DE MANOBRAS E GEOTECNIA"
    ws["A11"].font = estilos["font_secao"]
    ws["A11"].fill = estilos["fill_azul_escuro"]
    ws["A11"].alignment = estilos["align_left"]

    headers = list(df.columns)

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(
            row=12,
            column=col_idx,
            value=col_name,
        )
        cell.font = estilos["font_cabecalho_tab"]
        cell.fill = estilos["fill_azul_escuro"]
        cell.alignment = estilos["align_center"]

    start_row = 13

    for r_idx, row in df.iterrows():
        curr_row = start_row + r_idx

        for c_idx, value in enumerate(row.values, 1):
            cell = ws.cell(
                row=curr_row,
                column=c_idx,
                value=value,
            )
            cell.font = estilos["font_dados"]
            cell.border = estilos["border_fina"]
            cell.alignment = estilos["align_center"]

            if headers[c_idx - 1] in [
                "De (m)",
                "Para (m)",
                "Avanço (m)",
                "Rec. (m)",
                "RQD (m)",
            ]:
                cell.number_format = "0.00"
            elif headers[c_idx - 1] in [
                "Rec (%)",
                "RQD (%)",
            ]:
                cell.number_format = '0.0"%"'

            if curr_row % 2 == 0:
                cell.fill = estilos["fill_zebrado"]

    # Total
    tot_row = start_row + len(df)

    ws.cell(
        row=tot_row,
        column=1,
        value="TOTAL / MÉDIA",
    ).font = estilos["font_bold"]

    ws.cell(
        row=tot_row,
        column=4,
        value=df["Avanço (m)"].sum(),
    ).font = estilos["font_bold"]
    ws.cell(row=tot_row, column=4).number_format = "0.00"

    ws.cell(
        row=tot_row,
        column=5,
        value=df["Rec. (m)"].sum(),
    ).font = estilos["font_bold"]
    ws.cell(row=tot_row, column=5).number_format = "0.00"

    ws.cell(
        row=tot_row,
        column=6,
        value=df["Rec (%)"].mean(),
    ).font = estilos["font_bold"]
    ws.cell(row=tot_row, column=6).number_format = '0.0"%"'

    ws.cell(
        row=tot_row,
        column=8,
        value=df["RQD (%)"].mean(),
    ).font = estilos["font_bold"]
    ws.cell(row=tot_row, column=8).number_format = '0.0"%"'

    # Assinaturas
    ass_row = tot_row + 3

    ws.merge_cells(
        start_row=ass_row,
        start_column=2,
        end_row=ass_row,
        end_column=5,
    )
    ws.cell(
        row=ass_row,
        column=2,
        value="________________________________________",
    ).alignment = estilos["align_center"]

    ws.merge_cells(
        start_row=ass_row + 1,
        start_column=2,
        end_row=ass_row + 1,
        end_column=5,
    )
    ws.cell(
        row=ass_row + 1,
        column=2,
        value=f"Geólogo Resp.: {dados_projeto['geologo']}",
    ).font = estilos["font_bold"]
    ws.cell(
        row=ass_row + 1,
        column=2,
    ).alignment = estilos["align_center"]

    ws.merge_cells(
        start_row=ass_row,
        start_column=8,
        end_row=ass_row,
        end_column=11,
    )
    ws.cell(
        row=ass_row,
        column=8,
        value="________________________________________",
    ).alignment = estilos["align_center"]

    ws.merge_cells(
        start_row=ass_row + 1,
        start_column=8,
        end_row=ass_row + 1,
        end_column=11,
    )
    ws.cell(
        row=ass_row + 1,
        column=8,
        value=f"Supervisor/Coordenador: {dados_projeto['supervisor']}",
    ).font = estilos["font_bold"]
    ws.cell(
        row=ass_row + 1,
        column=8,
    ).alignment = estilos["align_center"]

    # Largura das colunas
    for col in ws.columns:
        max_len = 0

        for cell in col:
            value = str(cell.value or "")
            if cell.coordinate in ws.merged_cells:
                continue
            max_len = max(max_len, len(value))

        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(buffer)
    return buffer.getvalue()


def renderizar_exportacao(dados_projeto, df):
    arquivo_excel = criar_excel_boletim(dados_projeto, df)

    st.download_button(
        label="📄 Baixar Boletim de Sondagem Oficial (.xlsx)",
        data=arquivo_excel,
        file_name=f"Boletim_Oficial_{dados_projeto['furo_id']}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# ============================================================
# PERFIL VISUAL
# ============================================================

def criar_perfil_visual(df, furo_id, projeto):
    fig, (ax1, ax2, ax3) = plt.subplots(
        1,
        3,
        figsize=(11, 6),
        sharey=True,
    )

    fig.suptitle(
        f"Perfil de Sondagem - Furo: {furo_id} | Projeto: {projeto}",
        fontsize=13,
        fontweight="bold",
    )

    prof_max = df["Para (m)"].max()
    ax1.set_ylim(prof_max, 0)

    for _, row in df.iterrows():
        de_m = row["De (m)"]
        para_m = row["Para (m)"]
        lito = row["Litologia"]
        cor = CORES_LITOLOGIA.get(lito, "#808080")

        ax1.add_patch(
            mpatches.Rectangle(
                (0, de_m),
                1,
                para_m - de_m,
                facecolor=cor,
                edgecolor="black",
            )
        )

        if (para_m - de_m) >= 0.3:
            ax1.text(
                0.5,
                de_m + (para_m - de_m) / 2,
                lito,
                ha="center",
                va="center",
                fontsize=8,
                color="black",
                weight="bold",
            )

    ax1.set_xlim(0, 1)
    ax1.set_title("Litologia", fontsize=11, fontweight="bold")
    ax1.set_ylabel("Profundidade (m)", fontsize=11)
    ax1.get_xaxis().set_visible(False)

    for _, row in df.iterrows():
        ax2.barh(
            y=row["De (m)"] + row["Avanço (m)"] / 2,
            width=row["Rec (%)"],
            height=row["Avanço (m)"],
            color="#3498db",
            edgecolor="black",
            alpha=0.8,
        )

    ax2.set_xlim(0, 105)
    ax2.axvline(
        100,
        color="red",
        linestyle="--",
        linewidth=1,
    )
    ax2.set_title(
        "Recuperação (%)",
        fontsize=11,
        fontweight="bold",
    )
    ax2.grid(True, linestyle=":", alpha=0.6)

    for _, row in df.iterrows():
        rqd_val = row["RQD (%)"]

        if rqd_val < 50:
            cor_rqd = "#e74c3c"
        elif rqd_val < 75:
            cor_rqd = "#f1c40f"
        else:
            cor_rqd = "#2ecc71"

        ax3.barh(
            y=row["De (m)"] + row["Avanço (m)"] / 2,
            width=rqd_val,
            height=row["Avanço (m)"],
            color=cor_rqd,
            edgecolor="black",
            alpha=0.8,
        )

    ax3.set_xlim(0, 105)
    ax3.axvline(
        50,
        color="orange",
        linestyle="--",
        linewidth=1,
    )
    ax3.set_title(
        "Qualidade RQD (%)",
        fontsize=11,
        fontweight="bold",
    )
    ax3.grid(True, linestyle=":", alpha=0.6)

    plt.tight_layout()
    return fig


def renderizar_perfil_visual(df, dados_projeto):
    st.header("4. Perfil Visual de Sondagem")

    fig = criar_perfil_visual(
        df,
        dados_projeto["furo_id"],
        dados_projeto["projeto"],
    )

    st.pyplot(fig)


# ============================================================
# APLICAÇÃO PRINCIPAL
# ============================================================

def main():
    inicializar_estado()

    renderizar_cabecalho()

    dados_projeto = renderizar_dados_projeto()

    st.markdown("---")

    renderizar_manobras()

    st.markdown("---")

    df = obter_dataframe_manobras()

    if df is not None:
        renderizar_tabela_consolidada(df)
        renderizar_exportacao(dados_projeto, df)

        st.markdown("---")

        renderizar_perfil_visual(df, dados_projeto)
    else:
        st.header("3. Tabela Consolidada e Exportação Oficial")
        st.info("Nenhuma manobra cadastrada até o momento.")

        st.markdown("---")

        st.header("4. Perfil Visual de Sondagem")
        st.info("Adicione uma manobra para visualizar o perfil.")


if __name__ == "__main__":
    main()
