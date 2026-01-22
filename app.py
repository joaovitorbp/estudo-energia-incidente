import streamlit as st
import math
import pandas as pd
import numpy as np

# ==============================================================================
# 1. BACKEND: TABELAS E CONSTANTES (IEEE 1584-2018) - INTACTAS
# ==============================================================================
TABLE_1 = {'VCB': [-0.04287, 1.035, -0.083, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092], 'VCBB': [-0.017432, 0.98, -0.05, 0, 0, -5.767e-9, 2.524e-6, -0.00034, 0.01187, 1.013], 'HCB': [0.054922, 0.988, -0.11, 0, 0, -5.382e-9, 2.316e-6, -0.000302, 0.0091, 0.9725], 'VOA': [0.043785, 1.04, -0.18, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092], 'HOA': [0.111147, 1.008, -0.24, 0, 0, -3.895e-9, 1.641e-6, -0.000197, 0.002615, 1.1]}
TABLE_2 = {'VCB': [0, -1.4269e-6, 8.3137e-5, -0.0019382, 0.022366, -0.12645, 0.30226], 'VCBB': [1.138e-6, -6.0287e-5, 0.0012758, -0.013778, 0.080217, -0.24066, 0.33524], 'HCB': [0, -3.097e-6, 0.00016405, -0.0033609, 0.033308, -0.16182, 0.34627], 'VOA': [9.5606e-7, -5.1543e-5, 0.0011161, -0.01242, 0.075125, -0.23584, 0.33696], 'HOA': [0, -3.1555e-6, 0.0001682, -0.0034607, 0.034124, -0.1599, 0.34629]}
TABLE_3 = {'VCB': [0.753364, 0.566, 1.752636, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092, 0, -1.598, 0.957], 'VCBB': [3.068459, 0.26, -0.098107, 0, 0, -5.767e-9, 2.524e-6, -0.00034, 0.01187, 1.013, -0.06, -1.809, 1.19], 'HCB': [4.073745, 0.344, -0.370259, 0, 0, -5.382e-9, 2.316e-6, -0.000302, 0.0091, 0.9725, 0, -2.03, 1.036], 'VOA': [0.679294, 0.746, 1.222636, 0, 0, -4.783e-9, 1.962e-6, -0.000229, 0.003141, 1.092, 0, -1.598, 0.997], 'HOA': [3.470417, 0.465, -0.261863, 0, 0, -3.895e-9, 1.641e-6, -0.000197, 0.002615, 1.1, 0, -1.99, 1.04]}
TABLE_7_TYPICAL = {'VCB': [-0.000302, 0.03441, 0.4325], 'VCBB': [-0.0002976, 0.032, 0.479], 'HCB': [-0.0001923, 0.01935, 0.6899]}
TABLE_7_SHALLOW = {'VCB': [0.002222, -0.02556, 0.6222], 'VCBB': [-0.002778, 0.1194, -0.2778], 'HCB': [-0.0005556, 0.03722, 0.4778]}
CONSTANTS_AB = {'VCB':  {'A': 4,  'B': 20}, 'VCBB': {'A': 10, 'B': 24}, 'HCB':  {'A': 10, 'B': 22}}

# ==============================================================================
# 2. FUNÇÕES AUXILIARES (LÓGICA MATEMÁTICA PURA)
# ==============================================================================
def log10(x):
    if x <= 0: return 0 
    return math.log10(x)

def obter_categoria_nfpa(e):
    if e <= 1.2: return "Isento (< 1.2)", "Não requer AR", "#28a745" # Verde
    if e <= 4.0: return "Categoria 1", "4.0 cal/cm²", "#ffc107" # Amarelo
    if e <= 8.0: return "Categoria 2", "8.0 cal/cm²", "#fd7e14" # Laranja
    if e <= 25.0: return "Categoria 3", "25.0 cal/cm²", "#dc3545" # Vermelho
    if e <= 40.0: return "Categoria 4", "40.0 cal/cm²", "#6f42c1" # Roxo
    return "PERIGO EXTREMO", "> 40 cal/cm²", "#000000" # Preto

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

def core_calculo_ieee1584(Voc_V, Ibf, Config, Gap, Dist, T_ms, T_min_ms, H_mm, W_mm, D_mm):
    Voc = Voc_V / 1000.0
    
    # --- CÁLCULO DAS CORRENTES (NOMINAL) ---
    k = TABLE_1[Config]
    term1 = 10 ** (k[0] + k[1]*log10(Ibf) + k[2]*log10(Gap))
    term2 = (k[3]*Ibf**6 + k[4]*Ibf**5 + k[5]*Ibf**4 + k[6]*Ibf**3 + k[7]*Ibf**2 + k[8]*Ibf + k[9])
    Iarc600 = term1 * term2
    
    term_a = (0.6/Voc)**2
    term_b = (1/Iarc600)**2 - ((0.6**2 - Voc**2)/(0.6**2 * Ibf**2))
    
    if term_a * term_b <= 0: return None # Erro matemático
    
    Iarc = 1 / math.sqrt(term_a * term_b)
    
    # --- CÁLCULO DE INVÓLUCRO E FATOR DE CORREÇÃO (CF) ---
    is_open = Config in ['VOA', 'HOA']
    if is_open: 
        CF, box_type, EES, H1, W1 = 1.0, "Open Air (Ar Livre)", 0.0, 0.0, 0.0
    else:
        EES, box_type, H1, W1 = calcular_ees_correto(Config, H_mm, W_mm, D_mm, Voc)
        b = TABLE_7_SHALLOW[Config] if "Shallow" in box_type else TABLE_7_TYPICAL[Config]
        CF = 1/(b[0]*EES**2 + b[1]*EES + b[2]) if "Shallow" in box_type else b[0]*EES**2 + b[1]*EES + b[2]

    # --- CÁLCULO DA CORRENTE REDUZIDA (IMIN) ---
    vk = TABLE_2[Config]
    VarCf = vk[0]*Voc**6 + vk[1]*Voc**5 + vk[2]*Voc**4 + vk[3]*Voc**3 + vk[4]*Voc**2 + vk[5]*Voc + vk[6]
    Imin = Iarc * (1 - 0.5 * VarCf)
    
    # --- CÁLCULO DE ENERGIA (FUNÇÃO INTERNA) ---
    def calc_energia_individual(I_current, Time_ms):
        tk = TABLE_3[Config]
        C2 = tk[3]*Ibf**7 + tk[4]*Ibf**6 + tk[5]*Ibf**5 + tk[6]*Ibf**4 + tk[7]*Ibf**3 + tk[8]*Ibf**2 + tk[9]*Ibf
        C3 = tk[10]*log10(Ibf) + tk[11]*log10(Dist) + log10(1/CF)
        exponent = tk[0] + tk[1]*log10(Gap) + (tk[2]*Iarc600/C2) + C3 + tk[12]*log10(I_current)
        E_cal = ((12.552/50)*Time_ms*(10**exponent))/4.184
        AFB = Dist * (1.2/E_cal)**(1/tk[11])
        return E_cal, AFB

    # Cenário 1: Nominal
    E_cal, AFB = calc_energia_individual(Iarc, T_ms)
    
    # Cenário 2: Reduzida
    E_min_cal, AFB_min = calc_energia_individual(Imin, T_min_ms)
    
    # Pior Caso
    E_final = max(E_cal, E_min_cal)
    AFB_final = max(AFB, AFB_min)
    
    return {
        "ia_600": Iarc600, "i_arc": Iarc, "i_min": Imin,
        "box_type": box_type, "ees": EES, "cf": CF, "var_cf": VarCf,
        "h1": H1, "w1": W1,
        "e_nominal": E_cal, "afb_nominal": AFB,
        "e_min": E_min_cal, "afb_min": AFB_min,
        "e_final": E_final, "afb_final": AFB_final,
        "pior_caso": "Corrente Nominal" if E_final == E_cal else "Corrente Reduzida"
    }

# ==============================================================================
# 3. FRONTEND: STREAMLIT APP
# ==============================================================================
st.set_page_config(page_title="Calc. Energia Incidente", layout="wide")

# Estilos CSS para Inputs Numéricos limpos e Cards
st.markdown("""
<style>
    .result-box { background-color: #f8f9fa; padding: 15px; border-radius: 8px; border: 1px solid #e9ecef; }
    .stNumberInput input { text-align: center; }
</style>
""", unsafe_allow_html=True)

st.title("⚡ Calculadora IEEE 1584-2018 (Arc Flash)")
st.markdown("---")

# --- SIDEBAR: DADOS DO EQUIPAMENTO ---
with st.sidebar:
    st.header("⚙️ Equipamento")
    equip_name = st.text_input("Nome/TAG", value="QGBT-GERAL")
    voltage = st.selectbox("Tensão Nominal (V)", [208, 220, 380, 440, 480, 600], index=3)
    config_electrode = st.selectbox("Configuração Eletrodos", ["VCB", "VCBB", "HCB", "VOA", "HOA"], index=0)
    
    st.markdown("---")
    st.subheader("Dimensões do Painel")
    is_open = config_electrode in ['VOA', 'HOA']
    h_mm = st.number_input("Altura (H) [mm]", value=2000, disabled=is_open)
    w_mm = st.number_input("Largura (W) [mm]", value=800, disabled=is_open)
    d_mm = st.number_input("Profundidade (D) [mm]", value=400, disabled=is_open)

# --- ÁREA PRINCIPAL: INPUTS NUMÉRICOS (SEM SLIDERS) ---
col_in1, col_in2, col_in3 = st.columns(3)

with col_in1:
    st.subheader("1. Sistema")
    ibf_ka = st.number_input("Icc - Curto Circuito (kA)", min_value=0.1, value=20.0, format="%.2f")
    gap_mm = st.number_input("Gap entre Condutores (mm)", min_value=1.0, value=32.0, format="%.1f")
    dist_mm = st.number_input("Distância de Trabalho (mm)", min_value=100, value=610)

with col_in2:
    st.subheader("2. Tempo Nominal")
    st.info("Para Corrente de Arco (Iarc)")
    time_ms = st.number_input("Tempo de Atuação (ms)", min_value=0.0, value=100.0, format="%.1f")

with col_in3:
    st.subheader("3. Tempo Reduzido")
    st.warning("Para Corrente Reduzida (Imin)")
    # Por padrão, sugerimos o mesmo tempo, mas o usuário altera se a proteção for diferente
    time_min_ms = st.number_input("Tempo de Atuação (Imin) [ms]", min_value=0.0, value=100.0, format="%.1f")

# --- CÁLCULO ---
# O Streamlit recalcula a cada Enter/Mudança de foco no input
res = core_calculo_ieee1584(voltage, ibf_ka, config_electrode, gap_mm, dist_mm, time_ms, time_min_ms, h_mm, w_mm, d_mm)

st.markdown("---")

if res:
    # Cores e Categorias
    cat_txt, epi_txt, color_hex = obter_categoria_nfpa(res['e_final'])
    
    # --- RESULTADO PRINCIPAL ---
    st.header("📊 Resultados Finais")
    
    col_main1, col_main2, col_main3 = st.columns([1, 1, 1.5])
    
    with col_main1:
        st.metric("Energia Incidente", f"{res['e_final']:.2f} cal/cm²")
        st.caption(f"Cenário: {res['pior_caso']}")
        
    with col_main2:
        st.metric("Fronteira de Arco (AFB)", f"{res['afb_final']:.0f} mm")
        st.caption("Distância segura sem EPI")
        
    with col_main3:
        st.markdown(f"""
        <div style="background-color:{color_hex}; color:white; padding:15px; border-radius:10px; text-align:center;">
            <h4 style="margin:0; color:white;">{cat_txt}</h4>
            <p style="margin:0;">{epi_txt}</p>
        </div>
        """, unsafe_allow_html=True)

    # --- DETALHAMENTO TÉCNICO (CÉLULAS COMPLETAS) ---
    st.subheader("📝 Detalhamento Técnico")
    
    with st.expander("Ver Todos os Parâmetros Calculados", expanded=True):
        
        c1, c2, c3 = st.columns(3)
        
        # Coluna 1: Correntes e Fatores
        with c1:
            st.markdown("**Parâmetros de Corrente**")
            st.write(f"Iarc (600V): **{res['ia_600']:.3f} kA**")
            st.write(f"Iarc (Final): **{res['i_arc']:.3f} kA**")
            st.write(f"Imin (Reduzida): **{res['i_min']:.3f} kA**")
            st.markdown("---")
            st.write(f"Variação CF: **{res['var_cf']:.4f}**")
            
        # Coluna 2: Invólucro
        with c2:
            st.markdown("**Geometria e Invólucro**")
            st.write(f"Tipo: **{res['box_type']}**")
            if not is_open:
                st.write(f"Dim. Equiv: **{res['h1']:.1f} x {res['w1']:.1f} in**")
            st.write(f"Tam. Equiv (EES): **{res['ees']:.2f} in**")
            st.write(f"Fator Correção (CF): **{res['cf']:.3f}**")

        # Coluna 3: Cenários Individuais
        with c3:
            st.markdown("**Comparativo de Cenários**")
            st.markdown(f"**1. Nominal ({time_ms} ms)**")
            st.write(f"E = {res['e_nominal']:.3f} cal/cm²")
            st.write(f"AFB = {res['afb_nominal']:.0f} mm")
            
            st.markdown(f"**2. Reduzida ({time_min_ms} ms)**")
            st.write(f"E = {res['e_min']:.3f} cal/cm²")
            st.write(f"AFB = {res['afb_min']:.0f} mm")

else:
    st.error("Erro matemático: Verifique a combinação Tensão x Corrente (Raiz negativa na fórmula IEEE).")
