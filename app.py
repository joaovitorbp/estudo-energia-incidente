import streamlit as st
import math
import pandas as pd
import numpy as np
import io
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from datetime import datetime

# ==============================================================================
# 1. BACKEND: LÓGICA DE CÁLCULO IEEE 1584-2018 (INTACTO)
# ==============================================================================
TABLE_1 = {'VCB': [-0.04287, 1.035, -0.083, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092], 'VCBB': [-0.017432, 0.98, -0.05, 0, 0, -5.767e-9, 2.524e-6, -0.00034, 0.01187, 1.013], 'HCB': [0.054922, 0.988, -0.11, 0, 0, -5.382e-9, 2.316e-6, -0.000302, 0.0091, 0.9725], 'VOA': [0.043785, 1.04, -0.18, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092], 'HOA': [0.111147, 1.008, -0.24, 0, 0, -3.895e-9, 1.641e-6, -0.000197, 0.002615, 1.1]}
TABLE_2 = {'VCB': [0, -1.4269e-6, 8.3137e-5, -0.0019382, 0.022366, -0.12645, 0.30226], 'VCBB': [1.138e-6, -6.0287e-5, 0.0012758, -0.013778, 0.080217, -0.24066, 0.33524], 'HCB': [0, -3.097e-6, 0.00016405, -0.0033609, 0.033308, -0.16182, 0.34627], 'VOA': [9.5606e-7, -5.1543e-5, 0.0011161, -0.01242, 0.075125, -0.23584, 0.33696], 'HOA': [0, -3.1555e-6, 0.0001682, -0.0034607, 0.034124, -0.1599, 0.34629]}
TABLE_3 = {'VCB': [0.753364, 0.566, 1.752636, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092, 0, -1.598, 0.957], 'VCBB': [3.068459, 0.26, -0.098107, 0, 0, -5.767e-9, 2.524e-6, -0.00034, 0.01187, 1.013, -0.06, -1.809, 1.19], 'HCB': [4.073745, 0.344, -0.370259, 0, 0, -5.382e-9, 2.316e-6, -0.000302, 0.0091, 0.9725, 0, -2.03, 1.036], 'VOA': [0.679294, 0.746, 1.222636, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092, 0, -1.598, 0.997], 'HOA': [3.470417, 0.465, -0.261863, 0, 0, -3.895e-9, 1.641e-6, -0.000197, 0.002615, 1.1, 0, -1.99, 1.04]}
TABLE_7_TYPICAL = {'VCB': [-0.000302, 0.03441, 0.4325], 'VCBB': [-0.0002976, 0.032, 0.479], 'HCB': [-0.0001923, 0.01935, 0.6899]}
TABLE_7_SHALLOW = {'VCB': [0.002222, -0.02556, 0.6222], 'VCBB': [-0.002778, 0.1194, -0.2778], 'HCB': [-0.0005556, 0.03722, 0.4778]}
CONSTANTS_AB = {'VCB':  {'A': 4,  'B': 20}, 'VCBB': {'A': 10, 'B': 24}, 'HCB':  {'A': 10, 'B': 22}}

# ==============================================================================
# 2. FUNÇÕES MATEMÁTICAS
# ==============================================================================
def log10(x): return math.log10(x) if x > 0 else 0

def obter_categoria_nfpa(e):
    if e <= 1.2: return "Isento (&lt; 1.2)", "#28a745", "N/A (< 1.2 cal/cm²)" 
    if e <= 4.0: return "Categoria 1", "#ffc107", "Min. 4.0 cal/cm²"
    if e <= 8.0: return "Categoria 2", "#fd7e14", "Min. 8.0 cal/cm²"
    if e <= 25.0: return "Categoria 3", "#dc3545", "Min. 25.0 cal/cm²"
    if e <= 40.0: return "Categoria 4", "#6f42c1", "Min. 40.0 cal/cm²"
    return "PERIGO EXTREMO", "#000000", "Não Operar (> 40 cal)"

def calcular_ajuste_linear(d, v, c):
    if c not in CONSTANTS_AB: return d / 25.4
    A, B = CONSTANTS_AB[c]['A'], CONSTANTS_AB[c]['B']
    return (660.4 + (d - 660.4) * ((v + A) / B)) * (1 / 25.4)

def calcular_ees_correto(C, H, W, D, V):
    H_in, W_in = H/25.4, W/25.4
    if V < 0.6 and H < 508 and W < 508 and D <= 203.2: return (H_in+W_in)/2, "Shallow (Raso)", H_in, W_in
    W1 = 20.0 if W < 508 else W_in if W <= 660.4 else calcular_ajuste_linear(W if W <= 1244.6 else 1244.6, V, C)
    H1 = (20.0 if H < 508 else H_in if H <= 1244.6 else 49.0) if C == 'VCB' else (20.0 if H < 508 else H_in if H <= 660.4 else calcular_ajuste_linear(H if H <= 1244.6 else 1244.6, V, C))
    return (H1 + W1)/2, "Typical (Típico)", H1, W1

def calcular_tudo(Voc_V, Ibf, Config, Gap, Dist, T_ms, T_min_ms, H_mm, W_mm, D_mm):
    if any(v is None for v in [Voc_V, Ibf, Config, Gap, Dist, H_mm, W_mm, D_mm]): return None
    if Ibf <= 0 or Gap <= 0 or Dist <= 0: return None

    Voc = Voc_V / 1000.0
    k = TABLE_1[Config]
    term1 = 10 ** (k[0] + k[1]*log10(Ibf) + k[2]*log10(Gap))
    term2 = (k[3]*Ibf**6 + k[4]*Ibf**5 + k[5]*Ibf**4 + k[6]*Ibf**3 + k[7]*Ibf**2 + k[8]*Ibf + k[9])
    Iarc600 = term1 * term2
    term_a = (0.6/Voc)**2
    term_b = (1/Iarc600)**2 - ((0.6**2 - Voc**2)/(0.6**2 * Ibf**2))
    if term_a * term_b <= 0: return None
    Iarc = 1 / math.sqrt(term_a * term_b)
    
    is_open = Config in ['VOA', 'HOA']
    if is_open: CF, box_type, EES, H1, W1 = 1.0, "Open Air (Ar Livre)", 0.0, 0.0, 0.0
    else:
        EES, box_type, H1, W1 = calcular_ees_correto(Config, H_mm, W_mm, D_mm, Voc)
        b = TABLE_7_SHALLOW[Config] if "Shallow" in box_type else TABLE_7_TYPICAL[Config]
        CF = 1/(b[0]*EES**2 + b[1]*EES + b[2]) if "Shallow" in box_type else b[0]*EES**2 + b[1]*EES + b[2]

    vk = TABLE_2[Config]
    VarCf = vk[0]*Voc**6 + vk[1]*Voc**5 + vk[2]*Voc**4 + vk[3]*Voc**3 + vk[4]*Voc**2 + vk[5]*Voc + vk[6]
    Imin = Iarc * (1 - 0.5 * VarCf)
    
    T_calc_nom = T_ms if T_ms is not None else 0
    T_calc_min = T_min_ms if T_min_ms is not None else 0
    tk = TABLE_3[Config]
    
    def get_energy(I_curr, Time):
        C2 = tk[3]*Ibf**7 + tk[4]*Ibf**6 + tk[5]*Ibf**5 + tk[6]*Ibf**4 + tk[7]*Ibf**3 + tk[8]*Ibf**2 + tk[9]*Ibf
        C3 = tk[10]*log10(Ibf) + tk[11]*log10(Dist) + log10(1/CF)
        exponent = tk[0] + tk[1]*log10(Gap) + (tk[2]*Iarc600/C2) + C3 + tk[12]*log10(I_curr)
        E = ((12.552/50)*Time*(10**exponent))/4.184
        AFB = Dist * (1.2/E)**(1/tk[11]) if E > 0 else 0
        return E, AFB

    E_cal, AFB = get_energy(Iarc, T_calc_nom)
    E_min_cal, AFB_min = get_energy(Imin, T_calc_min)
    E_final = max(E_cal, E_min_cal)
    AFB_final = max(AFB, AFB_min)
    
    return {
        "ia_600": Iarc600, "i_arc": Iarc, "i_min": Imin, "var_cf": VarCf,
        "e_nominal": E_cal, "afb_nominal": AFB, "e_min": E_min_cal, "afb_min": AFB_min,
        "e_final": E_final, "afb_final": AFB_final,
        "pior_caso": "Nominal" if E_final == E_cal else "Reduzida"
    }

# ==============================================================================
# 3. FRONTEND: STREAMLIT APP (V24.0 FINAL AJUSTADA)
# ==============================================================================
st.set_page_config(page_title="Calc. Energia Incidente", layout="wide")

# --- GERENCIAMENTO DE ESTADO (CALLBACKS) ---
if 'results' not in st.session_state: st.session_state.results = None
if 'inputs' not in st.session_state: st.session_state.inputs = {}
if 't_nom' not in st.session_state: st.session_state.t_nom = 0.0
if 't_min' not in st.session_state: st.session_state.t_min = 0.0

# Callback 1: Mudança no Sistema (Limpa Result + Zera Tempos)
def on_system_change():
    st.session_state.results = None
    st.session_state.t_nom = 0.0
    st.session_state.t_min = 0.0

# Callback 2: Mudança no Tempo (Limpa apenas Result)
def on_time_change():
    st.session_state.results = None

st.markdown("""
<style>
    .std-card { background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05); height: 100%; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; color: #1f2937; }
    .card-label { font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #6b7280; margin-bottom: 8px; }
    .card-value { font-size: 20px; font-weight: 700; color: #111827; }
    .card-unit { font-size: 14px; font-weight: 400; color: #6b7280; }
    .final-card { background-color: #ffffff; border: 2px solid #e5e7eb; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); height: 100%; display: flex; flex-direction: column; justify-content: center; align-items: center; }
    .final-label { font-size: 14px; font-weight: 700; text-transform: uppercase; color: #6b7280; margin-bottom: 12px; }
    .final-value { font-size: 32px; font-weight: 800; }
    .final-unit { font-size: 18px; font-weight: 500; color: #9ca3af; }
    .summary-footer { margin-top: 15px; padding: 10px 20px; background-color: #f9fafb; border-radius: 8px; text-align: center; color: #4b5563; font-size: 14px; display: flex; justify-content: center; gap: 20px; align-items: center; border: 1px solid #e5e7eb; }
    .summary-item { display: flex; align-items: center; gap: 6px; }
    .summary-label { font-weight: 500; color: #6b7280; }
    .summary-val-bold { font-weight: 700; color: #1f2937; text-transform: uppercase;}
    .stNumberInput input { text-align: center; }
    .vertical-divider { border-right: 1px solid #e5e7eb; height: 100%; width: 1px; margin: 0 auto; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Identificação")
    # Equipamento agora com key no session_state mas sem callback (não reseta cálculo)
    equip_name = st.text_input("TAG do Equipamento", value="", key="equip_tag")
    st.caption("Desenvolvido em Python | IEEE 1584-2018")

st.title("⚡ Calculadora de Energia Incidente")

# --- SEÇÃO 1: DADOS DO SISTEMA ---
st.subheader("1. Dados do Sistema Elétrico")
c1, c2, c3, c4 = st.columns(4)
voltage = c1.selectbox("Tensão de Alimentação (V)", [220, 380, 440, 480], index=None, placeholder="Selecione...", key="v_in", on_change=on_system_change)
config_electrode = c2.selectbox("Configuração", ["VCB", "VCBB", "HCB", "VOA", "HOA"], index=None, placeholder="Selecione...", key="cfg_in", on_change=on_system_change)
ibf_ka = c3.number_input("Corrente de Curto Circuito (kA)", min_value=0.0, value=None, step=1.0, format="%.2f", key="icc_in", on_change=on_system_change)
gap_mm = c4.number_input("Distância entre Condutores (mm)", min_value=0, value=None, step=1, format="%d", key="gap_in", on_change=on_system_change)

c5, c6, c7, c8 = st.columns(4)
dist_mm = c5.number_input("Distância de Trabalho (mm)", min_value=0, value=None, step=1, format="%d", key="dist_in", on_change=on_system_change)
is_open = config_electrode in ['VOA', 'HOA']
h_mm = c6.number_input("Altura do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d", key="h_in", on_change=on_system_change)
w_mm = c7.number_input("Largura do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d", key="w_in", on_change=on_system_change)
d_mm = c8.number_input("Profundidade do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d", key="d_in", on_change=on_system_change)

st.markdown("---")
# Pré-cálculo reativo
pre_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, 0, 0, h_mm, w_mm, d_mm)

# --- SEÇÃO 2: PROTEÇÃO E TEMPOS ---
st.subheader("2. Definição de Tempos de Proteção")
def card(label, value, unit="", color="#0056b3"):
    st.markdown(f"""<div class="std-card"><div class="card-label">{label}</div><div class="card-value" style="color: {color}">{value} <span class="card-unit">{unit}</span></div></div>""", unsafe_allow_html=True)

cp1, cp_sep, cp2 = st.columns([1, 0.1, 1])
with cp1:
    st.markdown("##### Cenário Nominal")
    col_a, col_b = st.columns([1, 1.5])
    val_iarc = f"{pre_res['i_arc']:.3f}" if pre_res else "-"
    with col_a: card("Corrente de Arco", val_iarc, "kA")
    with col_b: time_ms = st.number_input("Tempo de Atuação Cenário Nominal (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_nom", on_change=on_time_change)

with cp_sep: st.markdown('<div class="vertical-divider"></div>', unsafe_allow_html=True)

with cp2:
    st.markdown("##### Cenário Reduzido")
    col_c, col_d = st.columns([1, 1.5])
    val_imin = f"{pre_res['i_min']:.3f}" if pre_res else "-"
    with col_c: card("Corrente de Arco Red.", val_imin, "kA")
    with col_d: time_min_ms = st.number_input("Tempo de Atuação Cenário Reduzido (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_min", on_change=on_time_change)

st.markdown("<br>", unsafe_allow_html=True)

if st.button("CALCULAR ENERGIA FINAL", type="primary", use_container_width=True):
    if not pre_res:
        st.warning("⚠️ Preencha os dados do sistema primeiro.")
    # CORREÇÃO: Ambos devem ser maiores que 0 (Lógica OR para erro)
    elif st.session_state.t_nom <= 0 or st.session_state.t_min <= 0:
        st.warning("⚠️ Preencha os dois tempos de atuação (Nominal e Reduzido).")
    else:
        final_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, st.session_state.t_nom, st.session_state.t_min, h_mm, w_mm, d_mm)
        st.session_state.results = final_res
        st.session_state.inputs = {'voltage': voltage, 'dist_mm': dist_mm} # Equip name pegaremos dinamicamente

if st.session_state.results:
    res = st.session_state.results
    st.markdown("---")
    st.subheader("3. Resultados Intermediários")
    ri1, r_sep, ri2 = st.columns([1, 0.1, 1])
    with ri1:
        st.markdown("##### Cenário Nominal")
        c_nom1, c_nom2 = st.columns(2)
        with c_nom1: card("Energia Incidente", f"{res['e_nominal']:.2f}", "cal/cm²")
        with c_nom2: card("Fronteira de Arco (AFB)", f"{res['afb_nominal']:.0f}", "mm")
    with r_sep: st.markdown('<div class="vertical-divider"></div>', unsafe_allow_html=True)
    with ri2:
        st.markdown("##### Cenário Reduzido")
        c_red1, c_red2 = st.columns(2)
        with c_red1: card("Energia Incidente", f"{res['e_min']:.2f}", "cal/cm²")
        with c_red2: card("Fronteira de Arco (AFB)", f"{res['afb_min']:.0f}", "mm")

    st.markdown("---")
    st.subheader(f"4. Resultados Finais")
    cat_name, color_hex, cat_rate = obter_categoria_nfpa(res['e_final'])
    cf1, cf2, cf3 = st.columns(3)
    def final_card(label, value, unit, color):
        st.markdown(f"""<div class="final-card" style="border-color: {color};"><div class="final-label">{label}</div><div class="final-value" style="color: {color};">{value} <span class="final-unit">{unit}</span></div></div>""", unsafe_allow_html=True)

    with cf1: final_card("Energia Incidente", f"{res['e_final']:.2f}", "cal/cm²", color_hex)
    with cf2: final_card("Fronteira de Arco (AFB)", f"{res['afb_final']:.0f}", "mm", color_hex)
    with cf3: 
        st.markdown(f"""<div class="final-card" style="border-color: {color_hex};"><div class="final-label">Categoria de Risco</div><div style="background-color: {color_hex}; color: white; padding: 5px 15px; border-radius: 20px; font-weight: 700; font-size: 22px; white-space: nowrap;">{cat_name}</div></div>""", unsafe_allow_html=True)

    st.markdown(f"""<div class="summary-footer"><div class="summary-item"><span class="summary-label">Cenário Definidor:</span><span class="summary-val-bold">{res['pior_caso'].upper()}</span></div><div style="width: 1px; height: 15px; background: #d1d5db;"></div><div class="summary-item"><span class="summary-label">EPI Recomendado:</span><span class="summary-val-bold">{cat_rate}</span></div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")
    st.subheader("5. Exportação")

    def preencher_modelo_excel():
        try:
            wb = load_workbook("ADESIVO ENERGIA INCIDENTE - MODELO.xlsx")
            ws = wb.active
        except Exception as e:
            st.error(f"Erro ao carregar o modelo: {e}. Verifique se o arquivo está no GitHub.")
            return None

        # Helper para escrever na célula e lidar com Merge
        def write_cell(ws, r, c, val):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                for range_ in ws.merged_cells.ranges:
                    if cell.coordinate in range_:
                        # Escreve na célula principal do merge (top-left)
                        top_left = ws.cell(row=range_.min_row, column=range_.min_col)
                        top_left.value = val
                        return
            else:
                cell.value = val

        # Helper para buscar e preencher offset
        def fill_label(ws, text_to_find, val_to_write, col_offset):
            for r in ws.iter_rows(min_row=1, max_row=20, min_col=1, max_col=12):
                for cell in r:
                    if cell.value and isinstance(cell.value, str):
                        if text_to_find.lower() in str(cell.value).lower():
                            write_cell(ws, cell.row, cell.column + col_offset, val_to_write)
                            return

        # Dados da Sessão
        inp = st.session_state.inputs
        v_val = inp['voltage']
        
        if v_val < 50: zr, zc, classe = 0, 0, "-"
        elif v_val <= 500: zr, zc, classe = 200, 700, "00 (≤ 500 V)"
        elif v_val <= 1000: zr, zc, classe = 200, 700, "0 (≤ 1000 V)"
        else: zr, zc, classe = 700, 1500, "Consultar (> 1kV)"

        # 1. Preenchimento de Dados Técnicos
        fill_label(ws, "Energia incidente", f"{res['e_final']:.2f} cal/cm²", 2)
        fill_label(ws, "Limite do arco", f"{res['afb_final']:.0f} mm", 2)
        fill_label(ws, "Distância de trabalho", f"{inp['dist_mm']} mm", 2)
        fill_label(ws, "Categoria de risco", cat_name, 2)
        fill_label(ws, "Suportabilidade mínima", cat_rate.replace('Min. ', ''), 2)
        
        fill_label(ws, "Classe:", classe, 1)
        fill_label(ws, "Tensão:", f"{v_val} V", 1)
        fill_label(ws, "Zona controlada:", f"{zc} mm", 1)
        fill_label(ws, "Zona de risco:", f"{zr} mm", 1)
        
        # 2. Preenchimento do Equipamento Dinâmico (Usa o valor atual da widget)
        # Pega o valor atual do input, independente do momento do cálculo
        nome_atual = st.session_state.equip_tag 
        equip_text = f"Equipamento: {nome_atual or 'N/A'}"
        fill_label(ws, "Equipamento:", equip_text, 0)

        # 3. Data
        try:
            write_cell(ws, 13, 9, datetime.now().strftime('%b/%Y').upper())
        except: pass

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    excel_data = preencher_modelo_excel()
    if excel_data:
        # Nome do arquivo também usa o valor atual da tag
        f_tag = st.session_state.equip_tag
        f_name = f"Adesivo_{f_tag}.xlsx" if f_tag else "Adesivo_ArcFlash.xlsx"
        
        st.download_button(
            label="📥 Baixar Adesivo (Modelo Preenchido)",
            data=excel_data,
            file_name=f_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
