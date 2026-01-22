import streamlit as st
import math
import pandas as pd
import numpy as np

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
    # Retorna: (Nome Categoria, Rating EPI, Cor Hex)
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
    if any(v is None for v in [Voc_V, Ibf, Config, Gap, Dist, H_mm, W_mm, D_mm]):
        return None
    if Ibf <= 0 or Gap <= 0 or Dist <= 0:
        return None

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
    if is_open: 
        CF, box_type, EES, H1, W1 = 1.0, "Open Air (Ar Livre)", 0.0, 0.0, 0.0
        b_coeffs = []
    else:
        EES, box_type, H1, W1 = calcular_ees_correto(Config, H_mm, W_mm, D_mm, Voc)
        b = TABLE_7_SHALLOW[Config] if "Shallow" in box_type else TABLE_7_TYPICAL[Config]
        CF = 1/(b[0]*EES**2 + b[1]*EES + b[2]) if "Shallow" in box_type else b[0]*EES**2 + b[1]*EES + b[2]
        b_coeffs = b

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
        if E > 0: AFB = Dist * (1.2/E)**(1/tk[11])
        else: AFB = 0
        return E, AFB, exponent, C2, C3

    E_cal, AFB, exp1, C2_final, C3_final = get_energy(Iarc, T_calc_nom)
    E_min_cal, AFB_min, exp2, _, _ = get_energy(Imin, T_calc_min)
    E_final = max(E_cal, E_min_cal)
    AFB_final = max(AFB, AFB_min)
    
    return {
        "ia_600": Iarc600, "i_arc": Iarc, "i_min": Imin, "var_cf": VarCf,
        "box_type": box_type, "ees": EES, "cf": CF, "h1": H1, "w1": W1,
        "e_nominal": E_cal, "afb_nominal": AFB, "e_min": E_min_cal, "afb_min": AFB_min,
        "e_final": E_final, "afb_final": AFB_final,
        "pior_caso": "Nominal" if E_final == E_cal else "Reduzida",
        "k": k, "tk": tk, "vk": vk, "b": b_coeffs,
        "C2": C2_final, "C3": C3_final,
        "exp1": exp1, "exp2": exp2, "term1": term1, "term2": term2, "voc": Voc, "ibf": Ibf, "gap": Gap, "dist": Dist
    }

# ==============================================================================
# 3. FRONTEND: STREAMLIT APP (V18.0 FINAL AJUSTADA)
# ==============================================================================
st.set_page_config(page_title="Calc. Energia Incidente", layout="wide")

st.markdown("""
<style>
    /* 1. Base do Cartão (Usado em Seção 2 e 3) */
    .std-card {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        color: #1f2937; 
    }
    .card-label { font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #6b7280; margin-bottom: 8px; }
    .card-value { font-size: 20px; font-weight: 700; color: #111827; }
    .card-unit { font-size: 14px; font-weight: 400; color: #6b7280; }

    /* 2. Cartão de Resultado Final (Igual peso para os 3 itens) */
    .final-card {
        background-color: #ffffff;
        border: 2px solid #e5e7eb; /* Será sobrescrito via inline style */
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    
    .final-label {
        font-size: 14px;
        font-weight: 700;
        text-transform: uppercase;
        color: #6b7280;
        margin-bottom: 12px;
    }
    
    .final-value {
        font-size: 32px; 
        font-weight: 800;
    }
    
    .final-unit { font-size: 18px; font-weight: 500; color: #9ca3af; }
    
    /* 3. Rodapé de Resumo Discreto */
    .summary-footer {
        margin-top: 15px;
        padding: 10px 20px;
        background-color: #f9fafb;
        border-radius: 8px;
        text-align: center;
        color: #4b5563;
        font-size: 14px;
        display: flex;
        justify-content: center;
        gap: 20px;
        align-items: center;
        border: 1px solid #e5e7eb;
    }
    .summary-item { display: flex; align-items: center; gap: 6px; }
    .summary-label { font-weight: 500; color: #6b7280; }
    
    /* MODIFICAÇÃO: Cor neutra para o texto do EPI */
    .summary-val-bold { font-weight: 700; color: #1f2937; text-transform: uppercase;}

    /* Utilitários */
    .stNumberInput input { text-align: center; }
    .detail-row { border-bottom: 1px solid #eee; padding: 8px 0; font-family: monospace; font-size: 14px; }
    .detail-label { font-weight: bold; color: #444; }
    .detail-val { color: #007bff; float: right; }
    
    .vertical-divider {
        border-right: 1px solid #e5e7eb;
        height: 100%;
        width: 1px;
        margin: 0 auto;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Identificação")
    equip_name = st.text_input("TAG do Equipamento", value="")
    st.caption("Desenvolvido em Python | IEEE 1584-2018")

st.title("⚡ Calculadora de Energia Incidente")

# --- SEÇÃO 1: DADOS DO SISTEMA ---
st.subheader("1. Dados do Sistema Elétrico")

c1, c2, c3, c4 = st.columns(4)
voltage = c1.selectbox("Tensão de Alimentação (V)", [220, 380, 440, 480], index=None, placeholder="Selecione...")
config_electrode = c2.selectbox("Configuração", ["VCB", "VCBB", "HCB", "VOA", "HOA"], index=None, placeholder="Selecione...")
ibf_ka = c3.number_input("Corrente de Curto Circuito (kA)", min_value=0.0, value=None, step=1.0, format="%.2f")
gap_mm = c4.number_input("Distância entre Condutores (mm)", min_value=0, value=None, step=1, format="%d")

c5, c6, c7, c8 = st.columns(4)
dist_mm = c5.number_input("Distância de Trabalho (mm)", min_value=0, value=None, step=1, format="%d")
is_open = config_electrode in ['VOA', 'HOA']
h_mm = c6.number_input("Altura do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d")
w_mm = c7.number_input("Largura do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d")
d_mm = c8.number_input("Profundidade do Painel (mm)", min_value=0, value=None, step=1, disabled=is_open, format="%d")

st.markdown("---")
pre_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, 0, 0, h_mm, w_mm, d_mm)

# --- SEÇÃO 2: PROTEÇÃO E TEMPOS ---
st.subheader("2. Definição de Tempos de Proteção")

def card(label, value, unit="", color="#0056b3"):
    st.markdown(f"""
    <div class="std-card">
        <div class="card-label">{label}</div>
        <div class="card-value" style="color: {color}">{value} <span class="card-unit">{unit}</span></div>
    </div>
    """, unsafe_allow_html=True)

# Layout de 3 Colunas com Separador Vertical
cp1, cp_sep, cp2 = st.columns([1, 0.1, 1])

with cp1:
    st.markdown("##### Cenário Nominal")
    col_a, col_b = st.columns([1, 1.5])
    val_iarc = f"{pre_res['i_arc']:.3f}" if pre_res else "-"
    with col_a: card("Corrente de Arco", val_iarc, "kA")
    with col_b: time_ms = st.number_input("Tempo de Atuação Cenário Nominal (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_nom")

with cp_sep:
    st.markdown('<div class="vertical-divider"></div>', unsafe_allow_html=True)

with cp2:
    st.markdown("##### Cenário Reduzido")
    col_c, col_d = st.columns([1, 1.5])
    val_imin = f"{pre_res['i_min']:.3f}" if pre_res else "-"
    with col_c: card("Corrente de Arco Red.", val_imin, "kA")
    with col_d: time_min_ms = st.number_input("Tempo de Atuação Cenário Reduzido (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_min")

st.markdown("<br>", unsafe_allow_html=True)
calc_btn = st.button("CALCULAR ENERGIA FINAL", type="primary", use_container_width=True)

# --- RESULTADOS ---
if calc_btn:
    if not pre_res:
        st.warning("⚠️ Preencha os dados do sistema primeiro.")
    elif time_ms is None and time_min_ms is None:
        st.warning("⚠️ Preencha pelo menos um tempo de atuação.")
    else:
        final_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, time_ms, time_min_ms, h_mm, w_mm, d_mm)
        st.markdown("---")
        
        # --- SEÇÃO 3: INTERMEDIÁRIOS ---
        st.subheader("3. Resultados Intermediários")
        
        ri1, r_sep, ri2 = st.columns([1, 0.1, 1])
        with ri1:
            st.markdown("##### Cenário Nominal")
            c_nom1, c_nom2 = st.columns(2)
            with c_nom1: card("Energia Incidente", f"{final_res['e_nominal']:.2f}", "cal/cm²")
            with c_nom2: card("Fronteira de Arco (AFB)", f"{final_res['afb_nominal']:.0f}", "mm")
        
        with r_sep:
            st.markdown('<div class="vertical-divider"></div>', unsafe_allow_html=True)
            
        with ri2:
            st.markdown("##### Cenário Reduzido")
            c_red1, c_red2 = st.columns(2)
            with c_red1: card("Energia Incidente", f"{final_res['e_min']:.2f}", "cal/cm²")
            with c_red2: card("Fronteira de Arco (AFB)", f"{final_res['afb_min']:.0f}", "mm")

        # --- SEÇÃO 4: FINAL ---
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"4. Resultados Finais")
        
        cat_name, color_hex, cat_rate = obter_categoria_nfpa(final_res['e_final'])
        
        cf1, cf2, cf3 = st.columns(3)
        def final_card(label, value, unit, color):
            st.markdown(f"""
            <div class="final-card" style="border-color: {color};">
                <div class="final-label">{label}</div>
                <div class="final-value" style="color: {color};">{value} <span class="final-unit">{unit}</span></div>
            </div>
            """, unsafe_allow_html=True)

        with cf1: final_card("Energia Incidente", f"{final_res['e_final']:.2f}", "cal/cm²", color_hex)
        with cf2: final_card("Fronteira de Arco (AFB)", f"{final_res['afb_final']:.0f}", "mm", color_hex)
        with cf3: 
            st.markdown(f"""
            <div class="final-card" style="border-color: {color_hex};">
                <div class="final-label">Categoria de Risco</div>
                <div style="background-color: {color_hex}; color: white; padding: 5px 15px; border-radius: 20px; font-weight: 700; font-size: 22px; white-space: nowrap;">
                    {cat_name}
                </div>
            </div>
            """, unsafe_allow_html=True)

        # Rodapé Discreto (Texto Neutro - Sem Cor)
        st.markdown(f"""
        <div class="summary-footer">
            <div class="summary-item">
                <span class="summary-label">Cenário Definidor:</span>
                <span class="summary-val-bold">{final_res['pior_caso'].upper()}</span>
            </div>
            <div style="width: 1px; height: 15px; background: #d1d5db;"></div>
            <div class="summary-item">
                <span class="summary-label">EPI Recomendado:</span>
                <span class="summary-val-bold">{cat_rate}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Memória
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📝 Memória de Cálculo Detalhada"):
            d = final_res
            def row(label, val, unit=""):
                st.markdown(f"""<div class="detail-row"><span class="detail-label">{label}</span><span class="detail-val">{val} {unit}</span></div>""", unsafe_allow_html=True)
            
            st.markdown("#### 1. Parâmetros")
            row("Voc", d['voc']*1000, "V"); row("Ibf", d['ibf'], "kA"); row("G", d['gap'], "mm"); row("D", d['dist'], "mm")
            st.markdown("#### 2. Correntes"); row("Iarc", f"{d['i_arc']:.4f}", "kA"); row("Imin", f"{d['i_min']:.4f}", "kA")
            st.markdown("#### 3. Energias"); row("E_Nom", f"{d['e_nominal']:.4f}", "cal"); row("E_Red", f"{d['e_min']:.4f}", "cal")
