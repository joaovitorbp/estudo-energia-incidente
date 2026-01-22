import streamlit as st
import math
import pandas as pd
import numpy as np

# ==============================================================================
# 1. BACKEND: LÓGICA DE CÁLCULO IEEE 1584-2018
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
    if e <= 1.2: return "Isento (< 1.2)", "#28a745"
    if e <= 4.0: return "Categoria 1", "#ffc107"
    if e <= 8.0: return "Categoria 2", "#fd7e14"
    if e <= 25.0: return "Categoria 3", "#dc3545"
    if e <= 40.0: return "Categoria 4", "#6f42c1"
    return "PERIGO EXTREMO", "#000000"

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
    # Verificações de segurança
    if any(v is None for v in [Voc_V, Ibf, Config, Gap, Dist, H_mm, W_mm, D_mm]):
        return None
    if Ibf <= 0 or Gap <= 0 or Dist <= 0:
        return None

    Voc = Voc_V / 1000.0
    
    # 1. Correntes
    k = TABLE_1[Config]
    term1 = 10 ** (k[0] + k[1]*log10(Ibf) + k[2]*log10(Gap))
    term2 = (k[3]*Ibf**6 + k[4]*Ibf**5 + k[5]*Ibf**4 + k[6]*Ibf**3 + k[7]*Ibf**2 + k[8]*Ibf + k[9])
    Iarc600 = term1 * term2
    
    term_a = (0.6/Voc)**2
    term_b = (1/Iarc600)**2 - ((0.6**2 - Voc**2)/(0.6**2 * Ibf**2))
    
    if term_a * term_b <= 0: return None
    
    Iarc = 1 / math.sqrt(term_a * term_b)
    
    # 2. Caixa
    is_open = Config in ['VOA', 'HOA']
    if is_open: 
        CF, box_type, EES, H1, W1 = 1.0, "Open Air (Ar Livre)", 0.0, 0.0, 0.0
        b_coeffs = []
    else:
        EES, box_type, H1, W1 = calcular_ees_correto(Config, H_mm, W_mm, D_mm, Voc)
        b = TABLE_7_SHALLOW[Config] if "Shallow" in box_type else TABLE_7_TYPICAL[Config]
        CF = 1/(b[0]*EES**2 + b[1]*EES + b[2]) if "Shallow" in box_type else b[0]*EES**2 + b[1]*EES + b[2]
        b_coeffs = b

    # 3. Imin
    vk = TABLE_2[Config]
    VarCf = vk[0]*Voc**6 + vk[1]*Voc**5 + vk[2]*Voc**4 + vk[3]*Voc**3 + vk[4]*Voc**2 + vk[5]*Voc + vk[6]
    Imin = Iarc * (1 - 0.5 * VarCf)
    
    # Trata None como 0 para cálculo temporário
    T_calc_nom = T_ms if T_ms is not None else 0
    T_calc_min = T_min_ms if T_min_ms is not None else 0
    
    # 4. Energias
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
        "exp1": exp1, "exp2": exp2,
        "term1": term1, "term2": term2,
        "voc": Voc, "ibf": Ibf, "gap": Gap, "dist": Dist
    }

# ==============================================================================
# 3. FRONTEND: STREAMLIT APP (V6.0)
# ==============================================================================
st.set_page_config(page_title="Calc. Energia Incidente", layout="wide")

st.markdown("""
<style>
    /* Estilo dos Boxes de Corrente (Menos destaque) */
    .info-box { 
        background-color: #f8f9fa; /* Fundo bem claro */
        padding: 10px; 
        border-radius: 6px; 
        text-align: center; 
        border: 1px solid #e0e0e0; 
        height: 100%;
        color: #333 !important; 
    }
    .info-label {
        font-size: 13px;
        color: #666 !important;
        font-weight: 500;
        margin-bottom: 5px;
    }
    .info-value {
        font-size: 18px; /* Fonte reduzida (era 24px) */
        font-weight: bold;
        color: #0056b3 !important; /* Azul mais sóbrio */
    }
    
    /* Estilo do Box de Risco (Resultado Final - Destaque Mantido) */
    .risk-box { color: white; padding: 15px; border-radius: 8px; text-align: center; font-weight: bold; font-size: 20px; }
    
    /* Centralização de inputs */
    .stNumberInput input { text-align: center; }
    
    /* Detalhes */
    .detail-row { border-bottom: 1px solid #eee; padding: 8px 0; font-family: monospace; font-size: 14px; }
    .detail-label { font-weight: bold; color: #444; }
    .detail-val { color: #007bff; float: right; }
    .sub-result { background-color: #fff; padding: 10px; border-radius: 5px; border: 1px solid #ddd; text-align: center; color: #000;}
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
voltage = c1.selectbox("Tensão (V)", [220, 380, 440, 480], index=None, placeholder="Selecione...")
config_electrode = c2.selectbox("Configuração", ["VCB", "VCBB", "HCB", "VOA", "HOA"], index=None, placeholder="Selecione...")

# Icc continua Float
ibf_ka = c3.number_input("Icc - Curto Circuito (kA)", min_value=0.0, value=None, step=1.0, format="%.2f")

# Gap vira Inteiro (step=1, format="%d", min_value=0)
gap_mm = c4.number_input("Gap entre Condutores (mm)", min_value=0, value=None, step=1, format="%d")

c5, c6, c7, c8 = st.columns(4)
# Distância vira Inteiro
dist_mm = c5.number_input("Dist. Trabalho (mm)", min_value=0, value=None, step=1, format="%d")

is_open = config_electrode in ['VOA', 'HOA']
# Dimensões viram Inteiros
h_mm = c6.number_input("Altura Painel (H) [mm]", min_value=0, value=None, step=1, disabled=is_open, format="%d")
w_mm = c7.number_input("Largura Painel (W) [mm]", min_value=0, value=None, step=1, disabled=is_open, format="%d")
d_mm = c8.number_input("Profundidade (D) [mm]", min_value=0, value=None, step=1, disabled=is_open, format="%d")

st.markdown("---")

# --- PRÉ-CÁLCULO AUTOMÁTICO ---
pre_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, 0, 0, h_mm, w_mm, d_mm)

# --- SEÇÃO 2: PROTEÇÃO E TEMPOS ---
st.subheader("2. Definição de Tempos de Proteção")

cp1, cp2 = st.columns(2)

with cp1:
    st.markdown("##### Cenário Nominal")
    col_a, col_b = st.columns([1, 1.5])
    
    val_iarc = f"{pre_res['i_arc']:.3f} kA" if pre_res else "- kA"
    
    with col_a:
        # Box de Corrente com estilo mais discreto (font 18px)
        st.markdown(f"""
        <div class="info-box">
            <div class="info-label">Corrente (Iarc)</div>
            <div class="info-value">{val_iarc}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        time_ms = st.number_input("Tempo de Atuação (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_nom")

with cp2:
    st.markdown("##### Cenário Reduzido")
    col_c, col_d = st.columns([1, 1.5])
    
    val_imin = f"{pre_res['i_min']:.3f} kA" if pre_res else "- kA"
    
    with col_c:
        # Box de Corrente com estilo mais discreto
        st.markdown(f"""
        <div class="info-box">
            <div class="info-label">Corrente (Imin)</div>
            <div class="info-value">{val_imin}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_d:
        time_min_ms = st.number_input("Tempo de Atuação (ms)", min_value=0.0, value=None, step=0.1, format="%.1f", key="t_min")

st.markdown("<br>", unsafe_allow_html=True)

# --- BOTÃO PARA CALCULAR ---
calc_btn = st.button("CALCULAR ENERGIA FINAL", type="primary", use_container_width=True)

# --- SEÇÃO 3: RESULTADOS ---
if calc_btn:
    if not pre_res:
        st.warning("⚠️ Preencha todos os dados do sistema (Seção 1) antes de calcular.")
    elif time_ms is None and time_min_ms is None:
        st.warning("⚠️ Preencha pelo menos um tempo de atuação.")
    else:
        final_res = calcular_tudo(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, time_ms, time_min_ms, h_mm, w_mm, d_mm)
        
        st.markdown("---")
        st.subheader("3. Resultados do Estudo")
        
        # 3.1 Comparativo
        r1, r2 = st.columns(2)
        with r1:
            st.markdown(f"""<div class="sub-result"><strong>Cenário Nominal</strong><br>Energia: <b>{final_res['e_nominal']:.2f} cal/cm²</b><br>AFB: <b>{final_res['afb_nominal']:.0f} mm</b></div>""", unsafe_allow_html=True)
        with r2:
            st.markdown(f"""<div class="sub-result"><strong>Cenário Reduzido</strong><br>Energia: <b>{final_res['e_min']:.2f} cal/cm²</b><br>AFB: <b>{final_res['afb_min']:.0f} mm</b></div>""", unsafe_allow_html=True)

        # 3.2 Resultado Final (Destaque Mantido)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### ✅ Resultado Final (Pior Caso)")
        
        cat_txt, color_hex = obter_categoria_nfpa(final_res['e_final'])
        rf1, rf2, rf3 = st.columns([1, 1, 1.5])
        with rf1:
            st.metric("Energia Incidente Final", f"{final_res['e_final']:.2f} cal/cm²")
            st.caption(f"Definido pelo cenário: {final_res['pior_caso']}")
        with rf2:
            st.metric("Fronteira de Arco (AFB)", f"{final_res['afb_final']:.0f} mm")
        with rf3:
            st.markdown(f"""<div class="risk-box" style="background-color: {color_hex};">{cat_txt}</div>""", unsafe_allow_html=True)

        # 3.3 Memória
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📝 Memória de Cálculo Detalhada (Coeficientes e Variáveis)"):
            d = final_res
            def row(label, val, unit=""):
                st.markdown(f"""<div class="detail-row"><span class="detail-label">{label}</span><span class="detail-val">{val} {unit}</span></div>""", unsafe_allow_html=True)
            
            st.markdown("#### 1. Parâmetros de Entrada")
            row("Tensão (Voc)", d['voc']*1000, "V")
            row("Icc (Ibf)", d['ibf'], "kA")
            row("Gap (G)", d['gap'], "mm")
            row("Distância (D)", d['dist'], "mm")
            
            st.markdown("#### 2. Coeficientes da Tabela 1")
            for i, val in enumerate(d['k']): row(f"k{i+1}", val)
            
            st.markdown("#### 3. Cálculo de Correntes")
            row("Iarc_600 (Base)", f"{d['ia_600']:.4f}", "kA")
            row("Iarc (Nominal)", f"{d['i_arc']:.4f}", "kA")
            row("VarCf (Fator Imin)", f"{d['var_cf']:.4f}")
            row("Imin (Reduzida)", f"{d['i_min']:.4f}", "kA")
            
            st.markdown("#### 4. Ajuste de Invólucro")
            row("Tipo", d['box_type'])
            row("EES", f"{d['ees']:.4f}", "in")
            row("CF", f"{d['cf']:.4f}")
            if d['b']: 
                for i, val in enumerate(d['b']): row(f"b{i+1}", val)
            
            st.markdown("#### 5. Energias")
            row("Expoente (Nominal)", f"{d['exp1']:.4f}")
            row("Expoente (Reduzida)", f"{d['exp2']:.4f}")
            row("Energia (Nominal)", f"{d['e_nominal']:.4f}", "cal/cm²")
            row("Energia (Reduzida)", f"{d['e_min']:.4f}", "cal/cm²")
