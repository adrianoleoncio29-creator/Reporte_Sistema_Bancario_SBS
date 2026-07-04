"""
generar_dashboard_sfp.py
Monitor Interactivo del Sistema Financiero Peruano
Desarrollado por: Adriano Castillo
"""
import pandas as pd
import json
import os
import math
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# 1. CARGA
# ─────────────────────────────────────────────────────────────
BASE = Path('Output_SBS')
df_bg   = pd.read_csv(BASE / 'Balance_General_y_Estado_de_Ganancias_y_Perdidas.csv')
df_ind  = pd.read_csv(BASE / 'Indicadores_Financieros.csv')
df_mor  = pd.read_csv(BASE / 'Morosidad_por_Tipo_y_Modalidad.csv')
df_tc   = pd.read_csv(BASE / 'Numero_Tarjetas_Credito_Consumo.csv')
df_td   = pd.read_csv(BASE / 'Numero_Tarjetas_Debito.csv')
df_rank = pd.read_csv(BASE / 'Ranking_Creditos_Directos_por_Tipo.csv')

if df_bg['Hoja'].dtype != int:
    df_bg['Hoja'] = pd.to_numeric(df_bg['Hoja'], errors='coerce').astype('Int64')

# ─────────────────────────────────────────────────────────────
# NORMALIZACIÓN DE NOMBRES DE BANCO
# Problema: cada CSV usa nombres ligeramente distintos.
# Solución: mapeamos todo al nombre canónico usado en df_ind.
# ─────────────────────────────────────────────────────────────
NOMBRE_CANONICO = {
    # Balance / EGP  →  canónico (igual que df_ind)
    'Banco BBVA Perú':                  'B. BBVA Perú',
    'Banco BCI Perú':                   'B. BCI Perú',
    'Banco Falabella Perú':             'B. Falabella Perú',
    'Banco GNB':                        'B. GNB',
    'Banco ICBC':                       'B. ICBC',
    'Banco Interamericano de Finanzas': 'B. Interamericano de Finanzas',
    'Banco Pichincha':                  'B. Pichincha',
    'Banco Ripley':                     'B. Ripley',
    'Banco Santander Perú':             'B. Santander Perú',
    'Banco de Crédito del Perú':        'B. De Crédito del Perú',
    'Banco de Comercio':                'B. de Comercio',
    # Morosidad / Ranking / TC / TD  →  canónico
    'B. de Crédito del Perú':           'B. De Crédito del Perú',  # minúscula "de"
}

def normalizar(df, col='Banco'):
    df[col] = df[col].replace(NOMBRE_CANONICO)
    return df

df_bg   = normalizar(df_bg)
df_mor  = normalizar(df_mor)
df_tc   = normalizar(df_tc)
df_td   = normalizar(df_td)
df_rank = normalizar(df_rank)
# df_ind es la fuente canónica, no se toca

periodos = sorted(df_ind['Periodo'].astype(str).unique().tolist())

# Todos los bancos: unión de todas las fuentes ya normalizadas
todos_bancos = sorted(set(
    df_ind['Banco'].dropna().unique().tolist() +
    df_bg['Banco'].dropna().unique().tolist()
))

# ─────────────────────────────────────────────────────────────
# 2. HELPERS
# ─────────────────────────────────────────────────────────────
def sv(series, default=None):
    s = series.dropna()
    if len(s) == 0:
        return default
    v = float(s.iloc[0])
    return None if math.isnan(v) else round(v, 4)

def ind(p, indicador, sub, banco):
    m = ((df_ind['Periodo'].astype(str) == str(p)) &
         (df_ind['Indicador'] == indicador) &
         (df_ind['Sub_Indicador'] == sub) &
         (df_ind['Banco'] == banco))
    return sv(df_ind.loc[m, 'Valor'])

def ind_sis(p, indicador, sub):
    m = ((df_ind['Periodo'].astype(str) == str(p)) &
         (df_ind['Indicador'] == indicador) &
         (df_ind['Sub_Indicador'] == sub))
    s = df_ind.loc[m, 'Valor'].dropna()
    return round(float(s.mean()), 4) if len(s) > 0 else None

def egp(p, concepto, banco):
    m = ((df_bg['Periodo'].astype(str) == str(p)) &
         (df_bg['Tabla'] == 'Estado de Ganancias y Pérdidas') &
         (df_bg['Concepto_Cuenta'] == concepto) &
         (df_bg['Banco'] == banco))
    v = sv(df_bg.loc[m, 'Monto_Total'])
    return round(v / 1000, 3) if v is not None else None

def egp_sis(p, concepto):
    m = ((df_bg['Periodo'].astype(str) == str(p)) &
         (df_bg['Tabla'] == 'Estado de Ganancias y Pérdidas') &
         (df_bg['Concepto_Cuenta'] == concepto))
    s = df_bg.loc[m, 'Monto_Total'].dropna()
    return round(float(s.sum()) / 1000, 3) if len(s) > 0 else None

def mor(p, concepto, banco):
    m = ((df_mor['Periodo'].astype(str) == str(p)) &
         (df_mor['Concepto'] == concepto) &
         (df_mor['Banco'] == banco) &
         (df_mor['Sub_Concepto'].isna()))
    return sv(df_mor.loc[m, 'Porcentaje'])

def mor_sis(p, concepto):
    m = ((df_mor['Periodo'].astype(str) == str(p)) &
         (df_mor['Concepto'] == concepto) &
         (df_mor['Sub_Concepto'].isna()))
    s = df_mor.loc[m, 'Porcentaje'].dropna()
    return round(float(s.mean()), 4) if len(s) > 0 else None

def rank(p, cat, banco):
    m = ((df_rank['Periodo'].astype(str) == str(p)) &
         (df_rank['Categoria'] == cat) &
         (df_rank['Banco'] == banco))
    monto = sv(df_rank.loc[m, 'Monto'])
    part  = sv(df_rank.loc[m, 'Participacion'])
    return (round(monto / 1000, 3) if monto else None), part

def rank_sis(p, cat):
    m = ((df_rank['Periodo'].astype(str) == str(p)) &
         (df_rank['Categoria'] == cat))
    s = df_rank.loc[m, 'Monto'].dropna()
    return round(float(s.sum()) / 1000, 3) if len(s) > 0 else None

def col_total_banco(p, banco):
    cats = ['Consumo Revolvente', 'Consumo no Revolvente']
    m = ((df_rank['Periodo'].astype(str) == str(p)) &
         (df_rank['Categoria'].isin(cats)) &
         (df_rank['Banco'] == banco))
    sub = df_rank.loc[m]
    monto_sum = sub['Monto'].dropna().sum()
    part_sum  = sub['Participacion'].dropna().sum()
    col = round(monto_sum / 1000, 3) if monto_sum > 0 else None
    ms  = round(float(part_sum), 4) if part_sum > 0 else None
    return col, ms

def col_total_sis(p):
    cats = ['Consumo Revolvente', 'Consumo no Revolvente']
    m = ((df_rank['Periodo'].astype(str) == str(p)) &
         (df_rank['Categoria'].isin(cats)))
    s = df_rank.loc[m, 'Monto'].dropna()
    return round(float(s.sum()) / 1000, 3) if len(s) > 0 else None

CONCEPTOS_SALDO = ['Atrasados', 'Refinanciados y Reestructurados', 'Vigentes']

def saldo_banco(p, banco):
    m = ((df_bg['Periodo'].astype(str) == str(p)) &
         (df_bg['Hoja'] == 1) &
         (df_bg['Tabla'] == 'Activo') &
         (df_bg['Banco'] == banco) &
         (df_bg['Concepto_Cuenta'].isin(CONCEPTOS_SALDO)))
    sub = df_bg.loc[m]
    def gc(c):
        v = sub.loc[sub['Concepto_Cuenta'] == c, 'Monto_Total'].dropna()
        return round(float(v.iloc[0]) / 1000, 3) if len(v) > 0 else None
    s_vig = gc('Vigentes')
    s_atr = gc('Atrasados')
    s_ref = gc('Refinanciados y Reestructurados')
    vals  = [x for x in [s_vig, s_atr, s_ref] if x is not None]
    s_tot = round(sum(vals), 3) if vals else None
    return s_tot, s_atr, s_ref, s_vig

def saldo_sis(p):
    m = ((df_bg['Periodo'].astype(str) == str(p)) &
         (df_bg['Hoja'] == 1) &
         (df_bg['Tabla'] == 'Activo') &
         (df_bg['Concepto_Cuenta'].isin(CONCEPTOS_SALDO)))
    sub = df_bg.loc[m]
    def gc(c):
        v = sub.loc[sub['Concepto_Cuenta'] == c, 'Monto_Total'].dropna()
        return round(float(v.sum()) / 1000, 3) if len(v) > 0 else None
    s_vig = gc('Vigentes')
    s_atr = gc('Atrasados')
    s_ref = gc('Refinanciados y Reestructurados')
    vals  = [x for x in [s_vig, s_atr, s_ref] if x is not None]
    s_tot = round(sum(vals), 3) if vals else None
    return s_tot, s_atr, s_ref, s_vig

def tc_val(p, banco):
    m = ((df_tc['Periodo'].astype(str) == str(p)) & (df_tc['Banco'] == banco))
    return sv(df_tc.loc[m, 'Nro_TC'])

def tc_sis(p):
    s = df_tc.loc[df_tc['Periodo'].astype(str) == str(p), 'Nro_TC'].dropna()
    return int(s.sum()) if len(s) > 0 else None

def td_val(p, banco):
    m = ((df_td['Periodo'].astype(str) == str(p)) & (df_td['Banco'] == banco))
    return sv(df_td.loc[m, 'Nro_TD'])

def td_sis(p):
    s = df_td.loc[df_td['Periodo'].astype(str) == str(p), 'Nro_TD'].dropna()
    return int(s.sum()) if len(s) > 0 else None

# ─────────────────────────────────────────────────────────────
# 3. CONSTRUCCIÓN DEL JSON — TODOS LOS BANCOS
# ─────────────────────────────────────────────────────────────
print("Procesando periodos...")
data_json = {}

for p in periodos:
    print(f"  {p}...")
    entry = {}

    tc_s = tc_sis(p)
    td_s = td_sis(p)
    col_tot_s = col_total_sis(p)

    # ── Bloque Sistema ──
    bs = {
        'mfn':        egp_sis(p, 'MARGEN FINANCIERO NETO'),
        'rne':        egp_sis(p, 'RESULTADO NETO DEL EJERCICIO'),
        'ing_fin':    egp_sis(p, 'INGRESOS FINANCIEROS'),
        'roe':        ind_sis(p, 'RENTABILIDAD', 'Utilidad Neta Anualizada / Patrimonio Promedio'),
        'roa':        ind_sis(p, 'RENTABILIDAD', 'Utilidad Neta Anualizada / Activo Promedio'),
        'mor_consumo':mor_sis(p, 'Créditos de consumo'),
        'mor_global': ind_sis(p, 'CALIDAD DE ACTIVOS', 'Créditos Atrasados (criterio SBS) / Créditos Directos'),
        'cobertura':  ind_sis(p, 'CALIDAD DE ACTIVOS', 'Provisiones / Créditos Atrasados'),
        'eficiencia': ind_sis(p, 'EFICIENCIA Y GESTIÓN', 'Gastos de Operación / Margen Financiero Total'),
        'liq_mn':     ind_sis(p, 'LIQUIDEZ', 'Ratio de Liquidez MN'),
        'liq_me':     ind_sis(p, 'LIQUIDEZ', 'Ratio de Liquidez ME'),
        'capital':    ind_sis(p, 'SOLVENCIA', 'Ratio de Capital Global'),
        'col_rev':    rank_sis(p, 'Consumo Revolvente'),
        'ms_rev':     None,
        'tc':         tc_s,
        'ms_tc':      100.0,
        'td':         td_s,
        'ms_td':      100.0,
        'col_total':  col_tot_s,
        'ms_total':   None,
    }
    s_tot, s_atr, s_ref, s_vig = saldo_sis(p)
    bs.update({'saldo_total': s_tot, 'saldo_atrasados': s_atr,
               'saldo_refinanciados': s_ref, 'saldo_vigentes': s_vig})
    entry['sistema'] = bs

    # ── Bloque por cada banco ──
    for banco in todos_bancos:
        col_r,  ms_r  = rank(p, 'Consumo Revolvente', banco)
        col_tot, ms_tot = col_total_banco(p, banco)
        tc_r  = tc_val(p, banco)
        td_r  = td_val(p, banco)
        ms_tc_r = round(tc_r / tc_s * 100, 2) if tc_r and tc_s else None
        ms_td_r = round(td_r / td_s * 100, 2) if td_r and td_s else None
        s_tot_b, s_atr_b, s_ref_b, s_vig_b = saldo_banco(p, banco)

        bloque = {
            'mfn':        egp(p, 'MARGEN FINANCIERO NETO', banco),
            'rne':        egp(p, 'RESULTADO NETO DEL EJERCICIO', banco),
            'ing_fin':    egp(p, 'INGRESOS FINANCIEROS', banco),
            'roe':        ind(p, 'RENTABILIDAD', 'Utilidad Neta Anualizada / Patrimonio Promedio', banco),
            'roa':        ind(p, 'RENTABILIDAD', 'Utilidad Neta Anualizada / Activo Promedio', banco),
            'mor_consumo':mor(p, 'Créditos de consumo', banco),
            'mor_global': ind(p, 'CALIDAD DE ACTIVOS', 'Créditos Atrasados (criterio SBS) / Créditos Directos', banco),
            'cobertura':  ind(p, 'CALIDAD DE ACTIVOS', 'Provisiones / Créditos Atrasados', banco),
            'eficiencia': ind(p, 'EFICIENCIA Y GESTIÓN', 'Gastos de Operación / Margen Financiero Total', banco),
            'liq_mn':     ind(p, 'LIQUIDEZ', 'Ratio de Liquidez MN', banco),
            'liq_me':     ind(p, 'LIQUIDEZ', 'Ratio de Liquidez ME', banco),
            'capital':    ind(p, 'SOLVENCIA', 'Ratio de Capital Global', banco),
            'col_rev':    col_r,   'ms_rev':   ms_r,
            'tc':         tc_r,    'ms_tc':    ms_tc_r,
            'td':         td_r,    'ms_td':    ms_td_r,
            'col_total':  col_tot, 'ms_total': ms_tot,
            'saldo_total': s_tot_b, 'saldo_atrasados': s_atr_b,
            'saldo_refinanciados': s_ref_b, 'saldo_vigentes': s_vig_b,
        }
        entry[banco] = bloque

    data_json[str(p)] = entry

print("JSON construido OK.")

# ─────────────────────────────────────────────────────────────
# 4. HELPERS PYTHON PARA HTML
# ─────────────────────────────────────────────────────────────
MESES_ES = {
    '01':'Enero','02':'Febrero','03':'Marzo','04':'Abril',
    '05':'Mayo','06':'Junio','07':'Julio','08':'Agosto',
    '09':'Setiembre','10':'Octubre','11':'Noviembre','12':'Diciembre'
}
MESES_CORTO = {
    '01':'Ene','02':'Feb','03':'Mar','04':'Abr','05':'May','06':'Jun',
    '07':'Jul','08':'Ago','09':'Set','10':'Oct','11':'Nov','12':'Dic'
}

def plabel(p):
    s = str(p)
    return f"{MESES_ES.get(s[4:6], s[4:6])} {s[:4]}"

periodos_12    = periodos[-12:]
options_html   = '\n'.join(
    f'<option value="{p}">{plabel(p)}</option>'
    for p in reversed(periodos_12)
)
data_json_str  = json.dumps(data_json, ensure_ascii=False)
periodos_12_js = json.dumps([str(p) for p in periodos_12])
bancos_js      = json.dumps(todos_bancos, ensure_ascii=False)
primer_banco_a = json.dumps(todos_bancos[0] if todos_bancos else '')
primer_banco_b = json.dumps(todos_bancos[1] if len(todos_bancos) > 1 else todos_bancos[0])

# ─────────────────────────────────────────────────────────────
# 5. CSS
# ─────────────────────────────────────────────────────────────
CSS = """
:root {
  --azul:    #1a3a5c;
  --azul2:   #2563a8;
  --azul3:   #4a90d9;
  --acento:  #e84393;
  --acento2: #f0a500;
  --gris1:   #1e2533;
  --gris2:   #3d4a5c;
  --gris3:   #6b7a8d;
  --gris4:   #c8d0db;
  --fondo:   #f0f2f5;
  --blanco:  #ffffff;
  --verde:   #0e7c42;
  --rojo:    #c0392b;
  --verde-bg:#e6f4ed;
  --rojo-bg: #faeaea;
}
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--fondo);
       color: var(--gris1); font-size: 14px; }

/* HEADER */
.hdr { background: linear-gradient(135deg, var(--azul) 0%, #0d2540 100%);
       padding: 0 28px; display: flex; align-items: center;
       justify-content: space-between; height: 66px;
       box-shadow: 0 3px 14px rgba(0,0,0,.35); position: sticky; top: 0; z-index: 100; }
.hdr-brand { display: flex; align-items: center; gap: 12px; }
.hdr-brand-icon { width: 38px; height: 38px; background: var(--azul2);
                  border-radius: 8px; display: flex; align-items: center;
                  justify-content: center; font-size: 20px; }
.hdr-brand-text { color: #fff; }
.hdr-brand-text strong { display: block; font-size: 15px; font-weight: 800;
                          letter-spacing: -.2px; }
.hdr-brand-text span { font-size: 11px; color: rgba(255,255,255,.55);
                        text-transform: uppercase; letter-spacing: .8px; }
.hdr-center { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
.sel-wrap { display: flex; flex-direction: column; gap: 3px; }
.sel-wrap label { font-size: 10px; color: rgba(255,255,255,.55); font-weight: 700;
                  text-transform: uppercase; letter-spacing: .7px; }
.sel-wrap select { background: rgba(255,255,255,.1); border: 1px solid rgba(255,255,255,.25);
                   color: #fff; border-radius: 6px; padding: 5px 10px;
                   font-size: 13px; font-weight: 700; cursor: pointer; outline: none;
                   min-width: 150px; }
.sel-wrap select option { background: #1a3a5c; color: #fff; }
.hdr-sep { width: 1px; height: 30px; background: rgba(255,255,255,.2); }
"""

CSS2 = """
/* LAYOUT */
.main { max-width: 1440px; margin: 0 auto; padding: 20px 20px 32px; }
.sec-label { font-size: 11px; font-weight: 800; color: var(--azul2);
             text-transform: uppercase; letter-spacing: 1.2px;
             margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
.sec-label::before { content: ''; display: block; width: 3px; height: 14px;
                     background: var(--azul2); border-radius: 2px; }

/* HERO */
.hero { background: linear-gradient(135deg, var(--azul) 0%, var(--azul2) 60%, #1e5fa8 100%);
        border-radius: 12px; padding: 22px 28px; margin-bottom: 20px;
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 4px 18px rgba(26,58,92,.35); }
.hero-left { color: #fff; }
.hero-left h2 { font-size: 22px; font-weight: 800; letter-spacing: -.3px; }
.hero-left h2 .hl { color: var(--azul3); }
.hero-left p { font-size: 13px; color: rgba(255,255,255,.65); margin-top: 5px; }
.hero-banks { display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }
.bank-badge { background: rgba(255,255,255,.15); border: 1px solid rgba(255,255,255,.3);
              border-radius: 20px; padding: 4px 12px; font-size: 12px;
              color: rgba(255,255,255,.9); font-weight: 700; }
.bank-badge.ba { background: rgba(74,144,217,.35); border-color: var(--azul3); }
.bank-badge.bb { background: rgba(232,67,147,.25); border-color: var(--acento); }
.hero-right { display: flex; gap: 12px; }
.hero-stat { text-align: center; background: rgba(255,255,255,.1);
             border: 1px solid rgba(255,255,255,.2); border-radius: 10px;
             padding: 14px 18px; min-width: 100px; }
.hero-stat .val { font-size: 22px; font-weight: 800; color: #fff; }
.hero-stat .lbl { font-size: 10px; color: rgba(255,255,255,.6);
                  text-transform: uppercase; letter-spacing: .6px; margin-top: 3px; }
.hero-stat .sub { font-size: 11px; color: rgba(255,255,255,.45); margin-top: 2px; }

/* KPI CARDS */
.kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 20px; }
.kpi { background: var(--blanco); border-radius: 10px; padding: 16px 18px;
       box-shadow: 0 2px 8px rgba(0,0,0,.07); position: relative; overflow: hidden; }
.kpi::before { content: ''; position: absolute; top: 0; left: 0; right: 0;
               height: 4px; background: var(--azul2); }
.kpi.k2::before { background: var(--acento); }
.kpi.k3::before { background: var(--acento2); }
.kpi.k4::before { background: var(--azul3); }
.kpi-icon { font-size: 20px; margin-bottom: 6px; }
.kpi-lbl { font-size: 10px; font-weight: 700; color: var(--gris3);
           text-transform: uppercase; letter-spacing: .8px; margin-bottom: 6px; }
.kpi-val { font-size: 24px; font-weight: 800; color: var(--azul); line-height: 1; }
.kpi.k2 .kpi-val { color: var(--acento); }
.kpi.k3 .kpi-val { color: var(--acento2); }
.kpi.k4 .kpi-val { color: var(--azul2); }
.kpi-unit { font-size: 11px; color: var(--gris3); margin-top: 3px; }
.kpi-sub  { font-size: 10px; color: var(--azul3); margin-top: 2px; font-weight: 600; }
.kpi-vars { display: flex; gap: 5px; margin-top: 8px; flex-wrap: wrap; }
.kv { display: inline-flex; align-items: center; gap: 2px; font-size: 10px;
      font-weight: 700; padding: 2px 7px; border-radius: 20px; }
.kv-up { background: var(--verde-bg); color: var(--verde); }
.kv-dn { background: var(--rojo-bg); color: var(--rojo); }
"""

CSS3 = """
/* MID GRID */
.mid-grid { display: grid; grid-template-columns: 1fr 400px; gap: 14px; margin-bottom: 20px; }

/* TABLA */
.tbl-card { background: var(--blanco); border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,.07); overflow: hidden; }
.tbl-card table { width: 100%; border-collapse: collapse; table-layout: fixed; }
.tbl-card thead th { padding: 11px 12px; font-size: 11px; font-weight: 700; color: #fff;
                     text-transform: uppercase; letter-spacing: .6px; text-align: center;
                     background: var(--azul); }
.tbl-card thead th:first-child { text-align: left; width: 32%; }
.tbl-card thead th.th-b { background: var(--acento); }
.tbl-card thead th.th-s { background: var(--gris2); }
.tbl-card tbody tr:nth-child(even) { background: #f4f6f9; }
.tbl-card tbody tr:hover { background: #e8eef5; }
.tbl-card tbody td { padding: 7px 10px; font-size: 12px;
                     border-bottom: 1px solid #e4e9f0;
                     text-align: center; vertical-align: middle; }
.tbl-card tbody td:first-child { text-align: left; font-weight: 600; color: var(--gris1); }
.sec-row td { background: #e8eef5 !important; font-size: 9px !important;
              font-weight: 800 !important; color: var(--azul) !important;
              text-transform: uppercase !important; letter-spacing: .8px !important;
              padding: 6px 10px !important; }
.va { color: var(--azul2); font-weight: 700; }
.vb { color: var(--acento); font-weight: 700; }
.vs { color: var(--gris2); font-weight: 600; }
.vc { display: flex; flex-direction: column; align-items: center; gap: 2px; }
.vc-main { font-weight: 700; }
.vc-tags { display: flex; gap: 3px; justify-content: center; flex-wrap: wrap; }
.vt { font-size: 9px; font-weight: 700; padding: 1px 5px; border-radius: 8px; white-space: nowrap; }
.vt-up { background: var(--verde-bg); color: var(--verde); }
.vt-dn { background: var(--rojo-bg); color: var(--rojo); }

/* SIDE PANEL */
.side-panel { display: flex; flex-direction: column; gap: 12px; }
.mini-card { background: var(--blanco); border-radius: 10px;
             box-shadow: 0 2px 8px rgba(0,0,0,.07); padding: 14px 16px; }
.mini-title { font-size: 11px; font-weight: 800; color: var(--azul);
              text-transform: uppercase; letter-spacing: .7px;
              margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }
.gauge-row { display: flex; justify-content: space-around; gap: 8px; }
.gauge-item { text-align: center; }
.gauge-val { font-size: 18px; font-weight: 800; color: var(--azul2); }
.gauge-lbl { font-size: 10px; color: var(--gris3); text-transform: uppercase;
             letter-spacing: .5px; margin-top: 2px; }
.bar-row { display: flex; flex-direction: column; gap: 7px; }
.bar-item { display: flex; align-items: center; gap: 8px; }
.bar-name { font-size: 11px; color: var(--gris3); width: 90px; text-align: right; flex-shrink: 0;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.bar-track { flex: 1; background: var(--fondo); border-radius: 4px; height: 8px; overflow: hidden; }
.bar-fill { height: 100%; border-radius: 4px; transition: width .4s ease; }
.bar-val { font-size: 11px; font-weight: 700; color: var(--azul); width: 55px;
           flex-shrink: 0; text-align: right; line-height: 1.3; }
.pie-btn { background: #fff; border: 1px solid var(--azul2); color: var(--azul2);
           border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: 700; cursor: pointer; }
.pie-btn.active { background: var(--azul2); color: #fff; }
"""

CSS4 = """
/* CHARTS */
.chart-toggle-wrap { display: flex; align-items: center; gap: 6px; margin-bottom: 10px; flex-wrap: wrap; }
.tog { font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 20px;
       border: 1px solid var(--azul2); color: var(--azul2); background: #fff; cursor: pointer; }
.tog.active { background: var(--azul2); color: #fff; }
.charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 20px; }
.chart-card { background: var(--blanco); border-radius: 10px;
              box-shadow: 0 2px 8px rgba(0,0,0,.07); padding: 16px; }
.chart-hdr h4 { font-size: 13px; font-weight: 700; color: var(--azul); }
.chart-hdr p  { font-size: 11px; color: var(--gris3); margin-top: 2px; margin-bottom: 8px; }

/* INSIGHTS */
.insights-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 14px; margin-bottom: 24px; }
.ins-card { background: var(--blanco); border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,.07);
            padding: 18px; border-left: 4px solid var(--azul2); }
.ins-card.risk   { border-left-color: var(--acento2); }
.ins-card.profit { border-left-color: var(--acento); }
.ins-card.market { border-left-color: var(--azul3); }
.ins-icon  { font-size: 22px; margin-bottom: 8px; }
.ins-title { font-size: 12px; font-weight: 800; text-transform: uppercase;
             letter-spacing: .6px; color: var(--azul2); margin-bottom: 8px; }
.ins-card.risk   .ins-title { color: var(--acento2); }
.ins-card.profit .ins-title { color: var(--acento); }
.ins-card.market .ins-title { color: var(--azul3); }
.ins-text  { font-size: 11.5px; color: var(--gris2); line-height: 1.65; }

/* FOOTER */
.footer { background: var(--azul); color: rgba(255,255,255,.7);
          padding: 28px 20px; margin-top: 8px; }
.footer-inner { max-width: 1440px; margin: 0 auto;
                display: flex; flex-wrap: wrap; gap: 20px;
                justify-content: space-between; align-items: center; }
.footer-brand strong { display: block; color: #fff; font-size: 15px; font-weight: 800; }
.footer-brand span   { font-size: 12px; margin-top: 4px; display: block; }
.footer-contact { display: flex; flex-direction: column; gap: 5px; font-size: 12px; }
.footer-contact a { color: var(--azul3); text-decoration: none; }
.footer-contact a:hover { text-decoration: underline; }
.footer-source { font-size: 11px; text-align: right; }

/* ── SIDEBAR GLOSARIO ─────────────────────────────────── */
.page-wrap { display: flex; min-height: 100vh; }

.sidebar {
  width: 260px; flex-shrink: 0;
  background: var(--blanco);
  border-right: 2px solid var(--gris4);
  display: flex; flex-direction: column;
  position: sticky; top: 66px; height: calc(100vh - 66px);
  overflow: hidden;
  box-shadow: 2px 0 8px rgba(0,0,0,.07);
  z-index: 90;
  transition: width .3s ease;
}
.sidebar.collapsed { width: 42px; }

.sb-toggle {
  display: flex; align-items: center; justify-content: center;
  background: var(--azul2); color: #fff;
  border: none; cursor: pointer;
  padding: 10px; font-size: 14px;
  flex-shrink: 0; width: 100%;
  transition: background .2s;
}
.sb-toggle:hover { background: var(--azul); }
.sb-toggle-icon { font-size: 16px; transition: transform .3s; }
.sidebar.collapsed .sb-toggle-icon { transform: rotate(180deg); }

.sb-header {
  padding: 14px 14px 10px;
  border-bottom: 1px solid var(--gris4);
  flex-shrink: 0; overflow: hidden;
}
.sb-header h3 {
  font-size: 13px; font-weight: 800; color: var(--azul2);
  text-transform: uppercase; letter-spacing: .8px; white-space: nowrap;
}
.sb-header p { font-size: 9.5px; color: var(--gris3); margin-top: 3px; white-space: nowrap; }
.sidebar.collapsed .sb-header { display: none; }

.sb-terms {
  flex: 1; overflow-y: auto; padding: 10px 10px 0;
  scrollbar-width: thin; scrollbar-color: var(--gris4) transparent;
}
.sb-terms::-webkit-scrollbar { width: 4px; }
.sb-terms::-webkit-scrollbar-thumb { background: var(--gris4); border-radius: 4px; }
.sidebar.collapsed .sb-terms { display: none; }

.term-btn {
  display: block; width: 100%; text-align: left;
  background: transparent; border: none;
  padding: 7px 10px; border-radius: 6px;
  font-size: 13px; font-weight: 600; color: var(--gris1);
  cursor: pointer; margin-bottom: 3px;
  transition: background .15s, color .15s;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.term-btn:hover  { background: #e8eef5; color: var(--azul2); }
.term-btn.active { background: var(--azul2); color: #fff; }
.term-btn .term-icon { margin-right: 6px; font-size: 12px; }

.sb-def-panel {
  flex-shrink: 0;
  border-top: 1px solid var(--gris4);
  padding: 12px;
  max-height: 260px;
  overflow-y: auto;
  background: var(--fondo);
  scrollbar-width: thin; scrollbar-color: var(--gris4) transparent;
}
.sb-def-panel::-webkit-scrollbar { width: 4px; }
.sb-def-panel::-webkit-scrollbar-thumb { background: var(--gris4); border-radius: 4px; }
.sidebar.collapsed .sb-def-panel { display: none; }

.def-empty { font-size: 12px; color: var(--gris3); text-align: center; padding: 10px 0; font-style: italic; }
.def-title {
  font-size: 13px; font-weight: 800; color: var(--azul2);
  text-transform: uppercase; letter-spacing: .5px;
  margin-bottom: 6px; display: flex; align-items: center; gap: 6px;
}
.def-badge {
  font-size: 10px; font-weight: 700; padding: 1px 6px;
  border-radius: 10px; background: #e8eef5; color: var(--azul2);
  text-transform: none; letter-spacing: 0;
}
.def-text    { font-size: 12.5px; color: var(--gris2); line-height: 1.65; }
.def-formula {
  background: #e8eef5; border-left: 3px solid var(--azul2);
  border-radius: 0 4px 4px 0; padding: 6px 10px;
  margin-top: 8px; font-size: 11.5px; color: var(--azul);
  font-weight: 600; font-family: monospace;
}
.def-tip { font-size: 11.5px; color: var(--gris3); margin-top: 6px; font-style: italic; }

.content-area { flex: 1; overflow-x: hidden; min-width: 0; }

@media(max-width:1100px) {
  .mid-grid    { grid-template-columns: 1fr; }
  .kpi-grid    { grid-template-columns: repeat(2,1fr); }
  .charts-grid { grid-template-columns: 1fr; }
  .insights-grid { grid-template-columns: 1fr; }
  .hdr-center  { display: none; }
  .sidebar { width: 42px; }
  .sidebar .sb-header, .sidebar .sb-terms, .sidebar .sb-def-panel { display: none; }
}
"""

# ─────────────────────────────────────────────────────────────
# 6. HTML HEADER + BODY
# ─────────────────────────────────────────────────────────────
HTML_HEAD = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Monitor Interactivo del Sistema Financiero Peruano</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
{CSS}{CSS2}{CSS3}{CSS4}
</style>
</head>
<body>
"""

HTML_BODY = f"""
<header class="hdr">
  <div class="hdr-brand">
    <div class="hdr-brand-icon">🏦</div>
    <div class="hdr-brand-text">
      <strong>Monitor del Sistema Financiero Peruano</strong>
      <span>SBS · Banca Múltiple · Cifras en millones de soles</span>
    </div>
  </div>
  <div class="hdr-center">
    <div class="sel-wrap">
      <label>Periodo</label>
      <select id="periodSelect" onchange="updateDashboard()">
        {options_html}
      </select>
    </div>
    <div class="hdr-sep"></div>
    <div class="sel-wrap">
      <label>Banco A</label>
      <select id="bancoASelect" onchange="updateDashboard()"></select>
    </div>
    <div class="sel-wrap">
      <label>Banco B</label>
      <select id="bancoBSelect" onchange="updateDashboard()"></select>
    </div>
  </div>
</header>

<!-- ══ PAGE WRAP: sidebar + content ══ -->
<div class="page-wrap">

  <!-- ══ SIDEBAR GLOSARIO ══ -->
  <aside class="sidebar" id="sidebar">
    <button class="sb-toggle" onclick="toggleSidebar()" title="Expandir / Colapsar glosario">
      <span class="sb-toggle-icon">◀</span>
    </button>
    <div class="sb-header">
      <h3>📚 Glosario</h3>
      <p>Selecciona un indicador</p>
    </div>
    <div class="sb-terms" id="sb-terms"></div>
    <div class="sb-def-panel" id="sb-def-panel">
      <div class="def-empty">← Selecciona un indicador<br>para ver su definición</div>
    </div>
  </aside>

  <!-- ══ CONTENIDO PRINCIPAL ══ -->
  <div class="content-area">
<div class="main">
  <!-- HERO -->
  <div class="hero">
    <div class="hero-left">
      <h2>Sistema Financiero <span class="hl" id="hero-year">2026</span></h2>
      <p>Al cierre de <span id="hero-month">—</span> · Datos SBS · Banca Múltiple</p>
      <div class="hero-banks">
        <span class="bank-badge ba" id="badge-a">Banco A</span>
        <span class="bank-badge bb" id="badge-b">Banco B</span>
        <span class="bank-badge">🇵🇪 Total Sistema</span>
      </div>
    </div>
    <div class="hero-right">
      <div class="hero-stat">
        <div class="val" id="hero-roe">—</div>
        <div class="lbl">ROE</div><div class="sub">Banco A</div>
      </div>
      <div class="hero-stat">
        <div class="val" id="hero-mor">—</div>
        <div class="lbl">Morosidad</div><div class="sub">Banco A</div>
      </div>
      <div class="hero-stat">
        <div class="val" id="hero-ms">—</div>
        <div class="lbl">Mkt Share</div><div class="sub">Banco A</div>
      </div>
    </div>
  </div>

  <!-- KPI CARDS -->
  <div class="sec-label" id="kpi-label">Indicadores Clave · Banco A</div>
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-icon">💰</div>
      <div class="kpi-lbl">Margen Financiero Neto</div>
      <div class="kpi-val" id="kpi-mfn">—</div>
      <div class="kpi-unit">Millones de Soles</div>
      <div class="kpi-vars" id="kpi-mfn-vars"></div>
    </div>
    <div class="kpi k2">
      <div class="kpi-icon">📈</div>
      <div class="kpi-lbl">Resultado Neto del Ejercicio</div>
      <div class="kpi-val" id="kpi-rne">—</div>
      <div class="kpi-unit">Millones de Soles</div>
      <div class="kpi-vars" id="kpi-rne-vars"></div>
    </div>
    <div class="kpi k3">
      <div class="kpi-icon">⚠️</div>
      <div class="kpi-lbl">Morosidad Consumo</div>
      <div class="kpi-val" id="kpi-mor">—</div>
      <div class="kpi-unit">% Créditos de Consumo</div>
      <div class="kpi-vars" id="kpi-mor-vars"></div>
    </div>
    <div class="kpi k4">
      <div class="kpi-icon">🏆</div>
      <div class="kpi-lbl">Market Share Consumo Total</div>
      <div class="kpi-val" id="kpi-ms">—</div>
      <div class="kpi-unit">% Participación de Mercado</div>
      <div class="kpi-vars" id="kpi-ms-vars"></div>
    </div>
  </div>

  <!-- TABLA COMPARATIVA -->
  <div class="sec-label" id="tbl-label">Comparativo · Banco A vs Banco B vs Total Sistema</div>
  <div class="mid-grid">
    <div class="tbl-card">
      <table>
        <thead>
          <tr>
            <th>Indicador</th>
            <th id="th-a">🔵 Banco A</th>
            <th class="th-b" id="th-b">🔴 Banco B</th>
            <th class="th-s">⚫ Total Sistema</th>
          </tr>
        </thead>
        <tbody id="tbl-body">
          <tr><td colspan="4" style="text-align:center;padding:24px;color:#8a96b0">Cargando...</td></tr>
        </tbody>
      </table>
    </div>

    <div class="side-panel">
      <div class="mini-card">
        <div class="mini-title">🛡️ Solvencia &amp; Liquidez · Banco A</div>
        <div class="gauge-row">
          <div class="gauge-item">
            <div class="gauge-val" id="g-capital">—</div>
            <div class="gauge-lbl">Capital Global</div>
          </div>
          <div class="gauge-item">
            <div class="gauge-val" id="g-liqmn">—</div>
            <div class="gauge-lbl">Liquidez MN</div>
          </div>
          <div class="gauge-item">
            <div class="gauge-val" id="g-liqme">—</div>
            <div class="gauge-lbl">Liquidez ME</div>
          </div>
        </div>
      </div>

      <div class="mini-card">
        <div class="mini-title">📊 ROE · ROA · Eficiencia · Banco A</div>
        <div class="gauge-row">
          <div class="gauge-item">
            <div class="gauge-val" id="g-roe">—</div>
            <div class="gauge-lbl">ROE</div>
          </div>
          <div class="gauge-item">
            <div class="gauge-val" id="g-roa">—</div>
            <div class="gauge-lbl">ROA</div>
          </div>
          <div class="gauge-item">
            <div class="gauge-val" id="g-efi">—</div>
            <div class="gauge-lbl">Eficiencia</div>
          </div>
        </div>
      </div>

      <div class="mini-card">
        <div class="mini-title">💳 Tarjetas de Crédito</div>
        <div class="bar-row" id="bar-tc"></div>
      </div>

      <div class="mini-card">
        <div class="mini-title">🏧 Tarjetas de Débito</div>
        <div class="bar-row" id="bar-td"></div>
      </div>

      <div class="mini-card">
        <div class="mini-title">🥧 Composición del Saldo de Activos</div>
        <div style="display:flex;gap:6px;margin-bottom:8px">
          <button class="pie-btn active" onclick="updatePie('A',this)">Banco A</button>
          <button class="pie-btn" onclick="updatePie('B',this)">Banco B</button>
          <button class="pie-btn" onclick="updatePie('S',this)">Sistema</button>
        </div>
        <div id="chart-pie" style="height:240px"></div>
      </div>
    </div>
  </div>

  <!-- TENDENCIA -->
  <div class="sec-label">Tendencia Histórica · Últimos 12 Periodos</div>
  <div class="chart-toggle-wrap">
    <span style="font-size:11px;font-weight:700;color:var(--gris3);text-transform:uppercase;letter-spacing:.6px">Métrica:</span>
    <button class="tog active" data-m="mor">Morosidad</button>
    <button class="tog" data-m="roe">ROE</button>
    <button class="tog" data-m="rne">Utilidad Neta</button>
    <button class="tog" data-m="saldo_total">Saldo Total</button>
    <button class="tog" data-m="ms_total">Market Share</button>
  </div>
  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-hdr">
        <h4 id="chart-left-title">Morosidad de Consumo (%)</h4>
        <p id="chart-left-sub">Banco A vs Banco B</p>
      </div>
      <div id="chart-left" style="height:260px"></div>
    </div>
    <div class="chart-card">
      <div class="chart-hdr">
        <h4 id="chart-right-title">Morosidad de Consumo (%)</h4>
        <p id="chart-right-sub">Banco A vs Total Sistema</p>
      </div>
      <div id="chart-right" style="height:260px"></div>
    </div>
  </div>

  <!-- INSIGHTS -->
  <div class="sec-label">Análisis Automático</div>
  <div class="insights-grid">
    <div class="ins-card risk">
      <div class="ins-icon">🛡️</div>
      <div class="ins-title">Gestión de Riesgos</div>
      <div class="ins-text" id="ins-risk">Cargando...</div>
    </div>
    <div class="ins-card profit">
      <div class="ins-icon">📈</div>
      <div class="ins-title">Eficiencia y Rentabilidad</div>
      <div class="ins-text" id="ins-profit">Cargando...</div>
    </div>
    <div class="ins-card market">
      <div class="ins-icon">🏦</div>
      <div class="ins-title">Posición Competitiva</div>
      <div class="ins-text" id="ins-market">Cargando...</div>
    </div>
  </div>
</div>

  </div><!-- /content-area -->
</div><!-- /page-wrap -->

<!-- FOOTER -->
<footer class="footer">
  <div class="footer-inner">
    <div class="footer-brand">
      <strong>Monitor Interactivo del Sistema Financiero Peruano</strong>
      <span>Herramienta de análisis sectorial basada en datos públicos de la SBS.</span>
    </div>
    <div class="footer-contact">
      <span>👤 Desarrollado por: <strong style="color:#fff">Adriano Castillo</strong></span>
      <span>📱 <a href="tel:+51961855271">961 855 271</a></span>
      <span>✉️ <a href="mailto:adrianoleoncio29@gmail.com">adrianoleoncio29@gmail.com</a></span>
      <span>🔗 <a href="https://www.linkedin.com/in/adrianoleonciocastillodelcastillo/" target="_blank">LinkedIn / adrianoleonciocastillodelcastillo</a></span>
    </div>
    <div class="footer-source">
      Fuente: Superintendencia de Banca,<br>Seguros y AFP (SBS) · Perú
    </div>
  </div>
</footer>
"""

# ─────────────────────────────────────────────────────────────
# 7. JAVASCRIPT — Bloque 1: constantes y helpers
# ─────────────────────────────────────────────────────────────
JS1 = f"""
<script>
const DATA   = {data_json_str};
const P12    = {periodos_12_js};
const BANCOS = {bancos_js};

const MESES = {{'01':'Ene','02':'Feb','03':'Mar','04':'Abr','05':'May','06':'Jun',
                '07':'Jul','08':'Ago','09':'Set','10':'Oct','11':'Nov','12':'Dic'}};
const MESES_F = {{'01':'Enero','02':'Febrero','03':'Marzo','04':'Abril','05':'Mayo',
                  '06':'Junio','07':'Julio','08':'Agosto','09':'Setiembre',
                  '10':'Octubre','11':'Noviembre','12':'Diciembre'}};

const C_A = '#2563a8';
const C_B = '#e84393';
const C_S = '#6b7a8d';

let currentPeriodo = '';
let currentBancoA  = {primer_banco_a};
let currentBancoB  = {primer_banco_b};
let currentPieMode = 'A';
let currentMetrica = 'mor';

function plabel(p)  {{ const s=String(p); return MESES[s.slice(4,6)]+' '+s.slice(0,4); }}
function plabelF(p) {{ const s=String(p); return MESES_F[s.slice(4,6)]+' '+s.slice(0,4); }}

function fN(v,d=2) {{
  if(v==null) return 'N/D';
  return Number(v).toLocaleString('es-PE',{{minimumFractionDigits:d,maximumFractionDigits:d}});
}}
function fP(v)  {{ return v==null?'N/D':fN(v,2)+'%'; }}
function fMM(v) {{
  if(v==null) return 'N/D';
  return 'S/ '+fN(v,1)+' MM';
}}
function fE(v) {{ if(v==null) return 'N/D'; return Number(v).toLocaleString('es-PE'); }}

function prevPeriodo(p,months) {{
  const s=String(p),yr=parseInt(s.slice(0,4)),mo=parseInt(s.slice(4,6));
  let nm=mo-months,ny=yr;
  while(nm<=0){{ nm+=12; ny--; }}
  return String(ny)+String(nm).padStart(2,'0');
}}

function varPct(curr,prev,isRatio) {{
  if(curr==null||prev==null||prev===0) return null;
  if(isRatio) return parseFloat((curr-prev).toFixed(2));
  return parseFloat(((curr-prev)/Math.abs(prev)*100).toFixed(2));
}}

function varChip(v,label,isRatio,invertColors) {{
  if(v==null) return '';
  const up=invertColors?v<0:v>0;
  const cls=up?'kv-up':'kv-dn';
  const arrow=v>0?'▲':'▼';
  const suffix=isRatio?' pp':'%';
  return `<span class="kv ${{cls}}">${{arrow}} ${{Math.abs(v).toFixed(2)}}${{suffix}} ${{label}}</span>`;
}}

function varTag(v,label,isRatio,invertColors) {{
  if(v==null) return '';
  const up=invertColors?v<0:v>0;
  const cls=up?'vt-up':'vt-dn';
  const arrow=v>0?'▲':'▼';
  const suffix=isRatio?'pp':'%';
  return `<span class="vt ${{cls}}">${{arrow}}${{Math.abs(v).toFixed(2)}}${{suffix}} ${{label}}</span>`;
}}

function kpiVars(p,banco,key,isRatio,invertColors) {{
  const curr=DATA[p]?.[banco]?.[key];
  const pM=prevPeriodo(p,1), pA=prevPeriodo(p,12);
  const vM=varPct(curr,DATA[pM]?.[banco]?.[key],isRatio);
  const vA=varPct(curr,DATA[pA]?.[banco]?.[key],isRatio);
  return varChip(vM,'vs MA',isRatio,invertColors)+varChip(vA,'vs AA',isRatio,invertColors);
}}

function tblVars(p,banco,key,isRatio,invertColors) {{
  const curr=DATA[p]?.[banco]?.[key];
  const pM=prevPeriodo(p,1), pA=prevPeriodo(p,12);
  const vM=varPct(curr,DATA[pM]?.[banco]?.[key],isRatio);
  const vA=varPct(curr,DATA[pA]?.[banco]?.[key],isRatio);
  return varTag(vM,'MA',isRatio,invertColors)+varTag(vA,'AA',isRatio,invertColors);
}}

function vcell(mainHtml,tagsHtml) {{
  if(!tagsHtml) return mainHtml;
  return `<div class="vc"><div class="vc-main">${{mainHtml}}</div>
          <div class="vc-tags">${{tagsHtml}}</div></div>`;
}}

// Poblar selectores de banco
function populateBancoSelects() {{
  ['bancoASelect','bancoBSelect'].forEach((id,idx) => {{
    const sel = document.getElementById(id);
    sel.innerHTML = BANCOS.map(b =>
      `<option value="${{b}}">${{b}}</option>`).join('');
    sel.value = idx === 0 ? currentBancoA : currentBancoB;
  }});
}}
</script>
"""

# ─────────────────────────────────────────────────────────────
# 8. JAVASCRIPT — Bloque 2: tabla comparativa
# ─────────────────────────────────────────────────────────────
JS2 = """
<script>
function buildTable(d, p, bancoA, bancoB) {
  const ra=d[bancoA]||{}, rb=d[bancoB]||{}, rs=d['sistema']||{};
  const T=(banco,key,isR,inv)=>tblVars(p,banco,key,isR,inv);

  const row=(lbl,av,at,bv,bt,sv,st)=>
    `<tr>
      <td>${lbl}</td>
      <td class="va">${vcell(av,at)}</td>
      <td class="vb">${vcell(bv,bt)}</td>
      <td class="vs">${vcell(sv,st)}</td>
    </tr>`;
  const sec=(t)=>`<tr class="sec-row"><td colspan="4">${t}</td></tr>`;

  return [
    sec('A · Volúmenes de Negocio'),
    row('Colocaciones Consumo Revolvente',
        fMM(ra.col_rev), T(bancoA,'col_rev',false,false),
        fMM(rb.col_rev), T(bancoB,'col_rev',false,false),
        fMM(rs.col_rev), ''),
    row('Colocaciones Consumo Total',
        fMM(ra.col_total), T(bancoA,'col_total',false,false),
        fMM(rb.col_total), T(bancoB,'col_total',false,false),
        fMM(rs.col_total), ''),
    row('Market Share Consumo Revolvente',
        fP(ra.ms_rev), T(bancoA,'ms_rev',true,false),
        fP(rb.ms_rev), T(bancoB,'ms_rev',true,false),
        '100.00%', ''),
    row('Market Share Consumo Total',
        fP(ra.ms_total), T(bancoA,'ms_total',true,false),
        fP(rb.ms_total), T(bancoB,'ms_total',true,false),
        '—', ''),
    row('Tarjetas de Crédito (Nro.)',
        fE(ra.tc)+' <small style="color:#6b7a8d">'+(ra.ms_tc?fP(ra.ms_tc):'')+'</small>', T(bancoA,'tc',false,false),
        fE(rb.tc)+' <small style="color:#6b7a8d">'+(rb.ms_tc?fP(rb.ms_tc):'')+'</small>', T(bancoB,'tc',false,false),
        fE(rs.tc), ''),
    row('Tarjetas de Débito (Nro.)',
        fE(ra.td)+' <small style="color:#6b7a8d">'+(ra.ms_td?fP(ra.ms_td):'')+'</small>', T(bancoA,'td',false,false),
        fE(rb.td)+' <small style="color:#6b7a8d">'+(rb.ms_td?fP(rb.ms_td):'')+'</small>', T(bancoB,'td',false,false),
        fE(rs.td), ''),

    sec('B · Calidad de Activos y Riesgo'),
    row('Morosidad Consumo (%)',
        fP(ra.mor_consumo), T(bancoA,'mor_consumo',true,true),
        fP(rb.mor_consumo), T(bancoB,'mor_consumo',true,true),
        fP(rs.mor_consumo), ''),
    row('Morosidad Global SBS (%)',
        fP(ra.mor_global), T(bancoA,'mor_global',true,true),
        fP(rb.mor_global), T(bancoB,'mor_global',true,true),
        fP(rs.mor_global), ''),
    row('Cobertura de Provisiones (%)',
        fP(ra.cobertura), T(bancoA,'cobertura',true,false),
        fP(rb.cobertura), T(bancoB,'cobertura',true,false),
        fP(rs.cobertura), ''),

    sec('C · Gestión y Eficiencia'),
    row('Ratio de Eficiencia (%)',
        fP(ra.eficiencia), T(bancoA,'eficiencia',true,true),
        fP(rb.eficiencia), T(bancoB,'eficiencia',true,true),
        fP(rs.eficiencia), ''),

    sec('D · Liquidez y Solvencia'),
    row('Ratio de Liquidez MN (%)',
        fP(ra.liq_mn), T(bancoA,'liq_mn',true,false),
        fP(rb.liq_mn), T(bancoB,'liq_mn',true,false),
        fP(rs.liq_mn), ''),
    row('Ratio de Liquidez ME (%)',
        fP(ra.liq_me), T(bancoA,'liq_me',true,false),
        fP(rb.liq_me), T(bancoB,'liq_me',true,false),
        fP(rs.liq_me), ''),
    row('Ratio de Capital Global (%)',
        fP(ra.capital), T(bancoA,'capital',true,false),
        fP(rb.capital), T(bancoB,'capital',true,false),
        fP(rs.capital), ''),

    sec('E · Rentabilidad (EGP)'),
    row('Margen Financiero Neto',
        fMM(ra.mfn), T(bancoA,'mfn',false,false),
        fMM(rb.mfn), T(bancoB,'mfn',false,false),
        fMM(rs.mfn), ''),
    row('Resultado Neto del Ejercicio',
        fMM(ra.rne), T(bancoA,'rne',false,false),
        fMM(rb.rne), T(bancoB,'rne',false,false),
        fMM(rs.rne), ''),
    row('ROE (%)',
        fP(ra.roe), T(bancoA,'roe',true,false),
        fP(rb.roe), T(bancoB,'roe',true,false),
        fP(rs.roe), ''),
    row('ROA (%)',
        fP(ra.roa), T(bancoA,'roa',true,false),
        fP(rb.roa), T(bancoB,'roa',true,false),
        fP(rs.roa), ''),

    sec('F · Saldo de Activos por Calidad'),
    row('Saldo Total de Activos',
        fMM(ra.saldo_total), T(bancoA,'saldo_total',false,false),
        fMM(rb.saldo_total), T(bancoB,'saldo_total',false,false),
        fMM(rs.saldo_total), ''),
    row('Saldo Vigentes',
        fMM(ra.saldo_vigentes), T(bancoA,'saldo_vigentes',false,false),
        fMM(rb.saldo_vigentes), T(bancoB,'saldo_vigentes',false,false),
        fMM(rs.saldo_vigentes), ''),
    row('Saldo Atrasados',
        fMM(ra.saldo_atrasados), T(bancoA,'saldo_atrasados',false,true),
        fMM(rb.saldo_atrasados), T(bancoB,'saldo_atrasados',false,true),
        fMM(rs.saldo_atrasados), ''),
    row('Saldo Refinanciados y Reest.',
        fMM(ra.saldo_refinanciados), T(bancoA,'saldo_refinanciados',false,true),
        fMM(rb.saldo_refinanciados), T(bancoB,'saldo_refinanciados',false,true),
        fMM(rs.saldo_refinanciados), ''),
  ].join('');
}
</script>
"""

# ─────────────────────────────────────────────────────────────
# 9. JAVASCRIPT — Bloque 3: barras, donut, gráficos tendencia
# ─────────────────────────────────────────────────────────────
JS3 = """
<script>
function buildBars(id, items) {
  const maxV = Math.max(...items.map(i => i.val||0));
  document.getElementById(id).innerHTML = items.map(i => {
    const pct = maxV > 0 ? Math.round((i.val||0)/maxV*100) : 0;
    const short = i.name.length > 14 ? i.name.slice(0,14)+'…' : i.name;
    return `<div class="bar-item">
      <div class="bar-name" title="${i.name}">${short}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${pct}%;background:${i.color}"></div>
      </div>
      <div class="bar-val">${fE(i.val)}<br>
        <span style="font-size:8.5px;color:#8a96b0">${i.ms ? fP(i.ms) : ''}</span>
      </div>
    </div>`;
  }).join('');
}

function updatePie(mode, btn) {
  currentPieMode = mode;
  document.querySelectorAll('.pie-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  buildPieChart(mode, currentPeriodo, currentBancoA, currentBancoB);
}

function buildPieChart(mode, p, bancoA, bancoB) {
  const banco = mode==='A' ? bancoA : (mode==='B' ? bancoB : 'sistema');
  const d = DATA[p]?.[banco];
  if (!d) return;
  const vig  = d.saldo_vigentes      || 0;
  const atr  = d.saldo_atrasados     || 0;
  const ref  = d.saldo_refinanciados || 0;
  const total = vig + atr + ref;
  const color = mode==='A' ? C_A : (mode==='B' ? C_B : C_S);

  function lighten(hex, pct) {
    const n = parseInt(hex.slice(1),16);
    const r = Math.min(255, (n>>16)+Math.round(pct*255));
    const g = Math.min(255, ((n>>8)&0xff)+Math.round(pct*255));
    const b = Math.min(255, (n&0xff)+Math.round(pct*255));
    return `rgb(${r},${g},${b})`;
  }

  const trace = {
    type: 'pie', hole: 0.45,
    labels: [
      `Vigentes — S/${vig.toLocaleString('es-PE',{minimumFractionDigits:1,maximumFractionDigits:1})} MM`,
      `Atrasados — S/${atr.toLocaleString('es-PE',{minimumFractionDigits:1,maximumFractionDigits:1})} MM`,
      `Refinanciados — S/${ref.toLocaleString('es-PE',{minimumFractionDigits:1,maximumFractionDigits:1})} MM`,
    ],
    values: [vig, atr, ref],
    marker: { colors: [color, lighten(color,.25), lighten(color,.42)] },
    textinfo: 'percent', textposition: 'inside',
    insidetextorientation: 'radial',
    hovertemplate: '%{label}<br>%{percent:.1%}<extra></extra>',
  };
  const layout = {
    margin:{t:6,b:8,l:8,r:8}, showlegend:true,
    legend:{orientation:'h',x:.5,xanchor:'center',y:-0.1,
            font:{size:9,family:'Segoe UI,Arial'}},
    paper_bgcolor:'#fff', font:{family:'Segoe UI,Arial',size:10},
    annotations:[{
      font:{size:11,color:color,family:'Segoe UI,Arial'},
      showarrow:false,
      text:'S/'+total.toLocaleString('es-PE',{minimumFractionDigits:1,maximumFractionDigits:1})+' MM',
      x:.5,y:.5,
    }],
  };
  Plotly.react('chart-pie',[trace],layout,{responsive:true,displayModeBar:false});
}

const METRICA_CFG = {
  'mor':         {key:'mor_consumo',  titulo:'Morosidad de Consumo (%)',          sufijo:'%', isRatio:true },
  'roe':         {key:'roe',          titulo:'ROE – Retorno sobre Patrimonio (%)',sufijo:'%', isRatio:true },
  'rne':         {key:'rne',          titulo:'Resultado Neto del Ejercicio (MM)', sufijo:'',  isRatio:false},
  'saldo_total': {key:'saldo_total',  titulo:'Saldo Total de Activos (MM)',       sufijo:'',  isRatio:false},
  'ms_total':    {key:'ms_total',     titulo:'Market Share Consumo Total (%)',    sufijo:'%', isRatio:true },
};

function buildCharts(selP, metrica, bancoA, bancoB) {
  currentMetrica = metrica;
  const cfg = METRICA_CFG[metrica];
  const key = cfg.key;
  const labels = P12.map(plabel);
  const arrA = P12.map(p => DATA[p]?.[bancoA]?.[key] ?? null);
  const arrB = P12.map(p => DATA[p]?.[bancoB]?.[key] ?? null);
  const arrS = P12.map(p => DATA[p]?.['sistema']?.[key] ?? null);
  const selLbl = plabel(selP);
  const vline  = {type:'line',x0:selLbl,x1:selLbl,y0:0,y1:1,yref:'paper',
                  line:{color:'#2563a8',width:1.5,dash:'dot'}};
  const txtTpl = cfg.isRatio ? '%{y:.2f}%' : '%{y:,.1f} MM';
  const mkLayout = () => ({
    margin:{t:8,b:55,l:46,r:8},
    legend:{orientation:'h',y:-0.22,font:{size:10}},
    xaxis:{tickfont:{size:9},tickangle:-45,shapes:[vline]},
    yaxis:{tickfont:{size:10},ticksuffix:cfg.sufijo},
    plot_bgcolor:'#fff',paper_bgcolor:'#fff',
    font:{family:'Segoe UI,Arial'},hovermode:'x unified',
  });
  const mkTrace = (arr, name, color, dash) => ({
    x:labels, y:arr, name,
    mode:'lines+markers+text', textposition:'top center',
    texttemplate:txtTpl, textfont:{size:8,color},
    line:{color,width:2.5,dash:dash||'solid'}, marker:{size:4,color},
  });

  Plotly.react('chart-left',
    [mkTrace(arrA,bancoA,C_A), mkTrace(arrB,bancoB,C_B,'dot')],
    mkLayout(), {responsive:true,displayModeBar:false});
  Plotly.react('chart-right',
    [mkTrace(arrA,bancoA,C_A), mkTrace(arrS,'Total Sistema',C_S,'dash')],
    mkLayout(), {responsive:true,displayModeBar:false});

  document.getElementById('chart-left-title').textContent  = cfg.titulo;
  document.getElementById('chart-left-sub').textContent    = `${bancoA} vs ${bancoB}`;
  document.getElementById('chart-right-title').textContent = cfg.titulo;
  document.getElementById('chart-right-sub').textContent   = `${bancoA} vs Total Sistema`;
}
</script>
"""

# ─────────────────────────────────────────────────────────────
# 10. JAVASCRIPT — Bloque 4: insights + updateDashboard
# ─────────────────────────────────────────────────────────────
JS4 = """
<script>
function buildInsights(d, p, bancoA, bancoB) {
  const ra = d[bancoA]||{}, rb = d[bancoB]||{}, rs = d['sistema']||{};
  const lbl = plabelF(p);
  let risk='', profit='', market='';

  // Riesgo
  if (ra.mor_consumo != null) {
    const diff = rs.mor_consumo != null ? (ra.mor_consumo - rs.mor_consumo).toFixed(2) : null;
    const vs = diff != null
      ? (parseFloat(diff) < 0
          ? `, <strong>${Math.abs(diff)} pp por debajo</strong> del sistema`
          : `, <strong>${diff} pp por encima</strong> del sistema`)
      : '';
    risk = `En ${lbl}, <strong>${bancoA}</strong> registró una morosidad de consumo de <strong>${fP(ra.mor_consumo)}</strong>${vs}. `;
    if (rb.mor_consumo != null)
      risk += `<strong>${bancoB}</strong> se ubicó en <strong>${fP(rb.mor_consumo)}</strong>. `;
    if (ra.cobertura != null)
      risk += ra.cobertura >= 100
        ? `La cobertura de provisiones de <strong>${fP(ra.cobertura)}</strong> supera el 100%, reflejando sólida gestión preventiva.`
        : `La cobertura de provisiones de <strong>${fP(ra.cobertura)}</strong> se encuentra por debajo del 100%.`;
  } else {
    risk = `No hay datos de morosidad disponibles para ${bancoA} en ${lbl}.`;
  }

  // Rentabilidad
  if (ra.roe != null) {
    profit = ra.roe > 0
      ? `<strong>${bancoA}</strong> alcanzó un ROE de <strong>${fP(ra.roe)}</strong> en ${lbl}`
      : `El ROE de <strong>${bancoA}</strong> fue <strong>${fP(ra.roe)}</strong> en ${lbl}, con presión sobre la rentabilidad patrimonial`;
    if (rb.roe != null) profit += ` vs <strong>${fP(rb.roe)}</strong> de ${bancoB}`;
    profit += '. ';
    if (ra.eficiencia != null)
      profit += ra.eficiencia < 60
        ? `El ratio de eficiencia de <strong>${fP(ra.eficiencia)}</strong> refleja estructura de costos competitiva.`
        : `El ratio de eficiencia de <strong>${fP(ra.eficiencia)}</strong> señala margen de mejora en gastos operativos.`;
    if (ra.rne != null) profit += ` Resultado neto: <strong>${fMM(ra.rne)}</strong>.`;
  } else {
    profit = `No hay datos de rentabilidad disponibles para ${bancoA} en ${lbl}.`;
  }

  // Posición competitiva
  if (ra.ms_total != null || ra.ms_rev != null) {
    const msA = ra.ms_total ?? ra.ms_rev;
    const msB = rb.ms_total ?? rb.ms_rev;
    market = `<strong>${bancoA}</strong> mantiene una participación de <strong>${fP(msA)}</strong> en Consumo Total`;
    if (msB != null) market += ` frente al <strong>${fP(msB)}</strong> de ${bancoB}`;
    market += '. ';
    if (ra.tc != null) {
      market += `Tarjetas de crédito: <strong>${fE(ra.tc)}</strong>`;
      if (rb.tc != null) market += ` vs <strong>${fE(rb.tc)}</strong> de ${bancoB}`;
      market += '. ';
    }
    if (ra.capital != null)
      market += ra.capital >= 14
        ? `Ratio de capital global de <strong>${fP(ra.capital)}</strong>, por encima del mínimo regulatorio.`
        : `Ratio de capital global de <strong>${fP(ra.capital)}</strong>, próximo al límite regulatorio.`;
  } else {
    market = `No hay datos de posición competitiva disponibles para ${bancoA} en ${lbl}.`;
  }
  return {risk, profit, market};
}

// ── UPDATE PRINCIPAL ──
function updateDashboard() {
  const p      = document.getElementById('periodSelect').value;
  const bancoA = document.getElementById('bancoASelect').value;
  const bancoB = document.getElementById('bancoBSelect').value;
  const d      = DATA[p];
  if (!d) return;
  currentPeriodo = p;
  currentBancoA  = bancoA;
  currentBancoB  = bancoB;

  const ra = d[bancoA]||{};

  // Hero
  document.getElementById('hero-year').textContent  = String(p).slice(0,4);
  document.getElementById('hero-month').textContent = plabelF(p);
  document.getElementById('hero-roe').textContent   = ra.roe        != null ? fP(ra.roe)        : 'N/D';
  document.getElementById('hero-mor').textContent   = ra.mor_consumo!= null ? fP(ra.mor_consumo): 'N/D';
  document.getElementById('hero-ms').textContent    = ra.ms_total   != null ? fP(ra.ms_total)   :
                                                       (ra.ms_rev    != null ? fP(ra.ms_rev)     : 'N/D');
  document.getElementById('badge-a').textContent    = bancoA;
  document.getElementById('badge-b').textContent    = bancoB;

  // KPI label y encabezados tabla
  document.getElementById('kpi-label').innerHTML =
    `<span style="width:3px;height:14px;background:var(--azul2);border-radius:2px;display:inline-block;margin-right:8px"></span>
     Indicadores Clave · ${bancoA}`;
  document.getElementById('tbl-label').innerHTML =
    `<span style="width:3px;height:14px;background:var(--azul2);border-radius:2px;display:inline-block;margin-right:8px"></span>
     Comparativo · ${bancoA} vs ${bancoB} vs Total Sistema`;
  document.getElementById('th-a').textContent = '🔵 ' + bancoA;
  document.getElementById('th-b').textContent = '🔴 ' + bancoB;

  // KPI cards
  document.getElementById('kpi-mfn').textContent    = fMM(ra.mfn);
  document.getElementById('kpi-mfn-vars').innerHTML = kpiVars(p,bancoA,'mfn',false,false);
  document.getElementById('kpi-rne').textContent    = fMM(ra.rne);
  document.getElementById('kpi-rne-vars').innerHTML = kpiVars(p,bancoA,'rne',false,false);
  document.getElementById('kpi-mor').textContent    = ra.mor_consumo!=null ? fP(ra.mor_consumo) : 'N/D';
  document.getElementById('kpi-mor-vars').innerHTML = kpiVars(p,bancoA,'mor_consumo',true,true);
  document.getElementById('kpi-ms').textContent     = ra.ms_total!=null ? fP(ra.ms_total) : 'N/D';
  document.getElementById('kpi-ms-vars').innerHTML  = kpiVars(p,bancoA,'ms_total',true,false);

  // Tabla
  document.getElementById('tbl-body').innerHTML = buildTable(d, p, bancoA, bancoB);

  // Gauges
  document.getElementById('g-capital').textContent = fP(ra.capital);
  document.getElementById('g-liqmn').textContent   = fP(ra.liq_mn);
  document.getElementById('g-liqme').textContent   = fP(ra.liq_me);
  document.getElementById('g-roe').textContent     = fP(ra.roe);
  document.getElementById('g-roa').textContent     = fP(ra.roa);
  document.getElementById('g-efi').textContent     = fP(ra.eficiencia);

  // Barras TC / TD
  const rb = d[bancoB]||{}, rs = d['sistema']||{};
  buildBars('bar-tc', [
    {name:bancoA, val:ra.tc, ms:ra.ms_tc, color:C_A},
    {name:bancoB, val:rb.tc, ms:rb.ms_tc, color:C_B},
    {name:'Sistema', val:rs.tc, ms:100, color:C_S},
  ]);
  buildBars('bar-td', [
    {name:bancoA, val:ra.td, ms:ra.ms_td, color:C_A},
    {name:bancoB, val:rb.td, ms:rb.ms_td, color:C_B},
    {name:'Sistema', val:rs.td, ms:100, color:C_S},
  ]);

  // Donut
  buildPieChart(currentPieMode, p, bancoA, bancoB);

  // Gráficos
  buildCharts(p, currentMetrica, bancoA, bancoB);

  // Insights
  const ins = buildInsights(d, p, bancoA, bancoB);
  document.getElementById('ins-risk').innerHTML   = ins.risk;
  document.getElementById('ins-profit').innerHTML = ins.profit;
  document.getElementById('ins-market').innerHTML = ins.market;
}

// Botones de métrica
document.addEventListener('DOMContentLoaded', () => {
  populateBancoSelects();
  document.querySelectorAll('.tog').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tog').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      buildCharts(currentPeriodo, btn.dataset.m, currentBancoA, currentBancoB);
    });
  });
  updateDashboard();
});
</script>
"""

# ─────────────────────────────────────────────────────────────
# 11. JAVASCRIPT — Bloque Glosario + Sidebar
# ─────────────────────────────────────────────────────────────
JS_GLOSARIO = """
<script>
// ═══════════════════════════════════════════════════════════
// GLOSARIO DE INDICADORES FINANCIEROS
// ═══════════════════════════════════════════════════════════
const GLOSARIO = [
  {
    id: 'roe',
    icon: '📈',
    nombre: 'ROE',
    categoria: 'Rentabilidad',
    definicion: 'El <strong>Retorno sobre el Patrimonio</strong> (Return on Equity) mide cuánto beneficio neto genera un banco por cada sol de capital propio invertido por sus accionistas. Es el indicador más usado para evaluar la rentabilidad desde la perspectiva del dueño.',
    formula: 'ROE = Utilidad Neta Anualizada / Patrimonio Promedio × 100',
    tip: '📌 Un ROE más alto indica mayor rentabilidad para los accionistas. En banca peruana, un ROE superior al 15% se considera sólido.'
  },
  {
    id: 'roa',
    icon: '🏦',
    nombre: 'ROA',
    categoria: 'Rentabilidad',
    definicion: 'El <strong>Retorno sobre Activos</strong> (Return on Assets) indica la eficiencia con la que el banco utiliza todos sus activos para generar utilidades. Mide la rentabilidad en relación al tamaño total del banco.',
    formula: 'ROA = Utilidad Neta Anualizada / Activo Total Promedio × 100',
    tip: '📌 Un ROA positivo y creciente refleja que el banco está usando bien sus recursos. Valores entre 1% y 2% son habituales en banca retail.'
  },
  {
    id: 'morosidad',
    icon: '⚠️',
    nombre: 'Morosidad',
    categoria: 'Calidad de Cartera',
    definicion: 'La <strong>tasa de morosidad</strong> mide el porcentaje de créditos que están en atraso (vencidos o en cobranza judicial) respecto al total de créditos otorgados. Es el principal indicador de la calidad de la cartera crediticia.',
    formula: 'Morosidad = Créditos Atrasados / Créditos Directos Totales × 100',
    tip: '📌 Una morosidad más baja es mejor. Indica que los clientes están cumpliendo con sus pagos. Si baja respecto al mes anterior, es señal positiva de gestión de riesgo.'
  },
  {
    id: 'mfn',
    icon: '💰',
    nombre: 'Margen Financiero Neto',
    categoria: 'Rentabilidad (EGP)',
    definicion: 'El <strong>Margen Financiero Neto</strong> es la ganancia que obtiene el banco por su actividad de intermediación financiera, después de restar las provisiones por créditos de alto riesgo. Representa el núcleo del negocio bancario: la diferencia entre lo que cobra por préstamos y lo que paga por depósitos, menos el costo del riesgo.',
    formula: 'MFN = Margen Financiero Bruto – Provisiones para Créditos Directos',
    tip: '📌 Un MFN creciente indica que el banco está mejorando su negocio principal. Se expresa en millones de soles (MM).'
  },
  {
    id: 'consumo_rev',
    icon: '💳',
    nombre: 'Consumo Revolvente',
    categoria: 'Volumen de Negocio',
    definicion: 'Los <strong>créditos de consumo revolvente</strong> son aquellos donde el cliente puede usar, pagar y volver a usar el crédito aprobado de forma flexible, principalmente a través de tarjetas de crédito. El saldo disponible se "recarga" conforme el cliente paga.',
    formula: 'Market Share = Saldo del banco / Saldo total del sistema × 100',
    tip: '📌 El consumo revolvente es el segmento más competido de la banca retail peruana, donde compiten bancos especializados como Falabella, Ripley y los bancos universales.'
  },
  {
    id: 'cobertura',
    icon: '🛡️',
    nombre: 'Cobertura de Provisiones',
    categoria: 'Calidad de Cartera',
    definicion: 'El <strong>ratio de cobertura de provisiones</strong> mide cuánto tiene reservado el banco (en provisiones) para cubrir sus créditos en mora. Una cobertura del 100% significa que el banco tiene reservado exactamente lo que le deben los clientes en atraso.',
    formula: 'Cobertura = Provisiones Constituidas / Créditos Atrasados × 100',
    tip: '📌 Una cobertura superior al 100% es señal de fortaleza: el banco tiene más reservas que deudas en mora. Cuanto más alto, más protegido está ante pérdidas inesperadas.'
  },
  {
    id: 'eficiencia',
    icon: '⚙️',
    nombre: 'Ratio de Eficiencia',
    categoria: 'Gestión y Eficiencia',
    definicion: 'El <strong>ratio de eficiencia operativa</strong> mide qué porcentaje de sus ingresos financieros destina el banco a cubrir sus gastos operativos (personal, locales, tecnología, etc.). Indica qué tan bien administra sus costos.',
    formula: 'Eficiencia = Gastos de Operación / Margen Financiero Total × 100',
    tip: '📌 Aquí un número MENOR es mejor. Un ratio de 50% significa que por cada S/ 1.00 de ingreso gasta S/ 0.50 en operar. Por debajo del 60% se considera eficiente en el sistema peruano.'
  },
  {
    id: 'liquidez',
    icon: '💧',
    nombre: 'Ratio de Liquidez',
    categoria: 'Liquidez',
    definicion: 'El <strong>ratio de liquidez</strong> mide la capacidad del banco para atender sus obligaciones de corto plazo (retiros, vencimientos) con sus activos líquidos disponibles. Se calcula por separado en Moneda Nacional (MN) y Moneda Extranjera (ME).',
    formula: 'Liquidez MN = Activos Líquidos MN / Pasivos de Corto Plazo MN × 100',
    tip: '📌 La SBS exige un mínimo del 8% para MN y 20% para ME. Un ratio mayor indica mayor colchón de seguridad frente a retiros masivos o shocks de liquidez.'
  },
  {
    id: 'capital',
    icon: '🏛️',
    nombre: 'Ratio de Capital Global',
    categoria: 'Solvencia',
    definicion: 'El <strong>Ratio de Capital Global</strong> (o ratio de adecuación de capital) mide la solidez financiera del banco: qué tan bien está respaldado su portafolio de activos riesgosos con capital propio. Es el principal indicador regulatorio de solvencia en el Perú.',
    formula: 'Capital Global = Patrimonio Efectivo / Activos Ponderados por Riesgo × 100',
    tip: '📌 La SBS exige un mínimo del 10% (Basilea III). Ratios por encima del 14% reflejan una posición de capital muy sólida y capacidad de crecimiento sin necesidad de capitalización inmediata.'
  },
  {
    id: 'solvencia',
    icon: '⚖️',
    nombre: 'Solvencia',
    categoria: 'Solvencia',
    definicion: 'La <strong>solvencia</strong> es la capacidad estructural del banco para hacer frente a sus obligaciones de largo plazo con todos sus activos. A diferencia de la liquidez (corto plazo), evalúa si el banco podría pagar todas sus deudas en un escenario de liquidación.',
    formula: 'Pasivo Total / Capital Social y Reservas = N° de veces (apalancamiento)',
    tip: '📌 Un banco solvente tiene más activos que pasivos. El ratio de capital global es la medida regulatoria más usada para cuantificar la solvencia en el sistema financiero peruano.'
  },
  {
    id: 'market_share',
    icon: '🏆',
    nombre: 'Market Share',
    categoria: 'Posición Competitiva',
    definicion: 'La <strong>participación de mercado</strong> (Market Share) mide el peso relativo de un banco dentro del total del sistema, para una métrica específica (colocaciones, tarjetas, depósitos). Indica la posición competitiva del banco en cada segmento.',
    formula: 'Market Share = Saldo del banco / Saldo total del sistema × 100',
    tip: '📌 Un market share creciente indica ganancia de posición competitiva. En consumo revolvente, los primeros 3 bancos concentran más del 60% del mercado.'
  },
  {
    id: 'saldo_activos',
    icon: '📊',
    nombre: 'Saldo de Activos',
    categoria: 'Balance',
    definicion: 'El <strong>saldo de activos crediticios</strong> representa el stock total de créditos otorgados, clasificados según su calidad de pago: <strong>Vigentes</strong> (al día), <strong>Atrasados</strong> (vencidos o en cobranza judicial) y <strong>Refinanciados y Reestructurados</strong> (con condiciones modificadas por dificultades del deudor).',
    formula: 'Saldo Total = Vigentes + Atrasados + Refinanciados y Reestructurados',
    tip: '📌 La composición del saldo revela la calidad de la cartera. Una mayor proporción de vigentes y una menor de atrasados refleja una cartera sana y bien gestionada.'
  },
];

// ── Renderizar botones del glosario ──
function renderGlosario() {
  const container = document.getElementById('sb-terms');
  container.innerHTML = GLOSARIO.map((t, i) =>
    `<button class="term-btn" onclick="showDef(${i})" id="term-btn-${i}">
       <span class="term-icon">${t.icon}</span>${t.nombre}
       <span style="font-size:8.5px;color:#8a96b0;display:block;margin-left:18px;margin-top:1px">${t.categoria}</span>
     </button>`
  ).join('');
}

// ── Mostrar definición ──
function showDef(idx) {
  document.querySelectorAll('.term-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('term-btn-' + idx).classList.add('active');
  const t = GLOSARIO[idx];
  const panel = document.getElementById('sb-def-panel');
  panel.innerHTML = `
    <div class="def-title">
      ${t.icon} ${t.nombre}
      <span class="def-badge">${t.categoria}</span>
    </div>
    <div class="def-text">${t.definicion}</div>
    <div class="def-formula">${t.formula}</div>
    <div class="def-tip">${t.tip}</div>
  `;
  panel.scrollTop = 0;
}

// ── Toggle sidebar ──
function toggleSidebar() {
  const sb = document.getElementById('sidebar');
  sb.classList.toggle('collapsed');
}

document.addEventListener('DOMContentLoaded', renderGlosario);
</script>
"""

# ─────────────────────────────────────────────────────────────
# 12. ENSAMBLAJE Y ESCRITURA DEL HTML
# ─────────────────────────────────────────────────────────────
full_html = (
    HTML_HEAD
    + HTML_BODY
    + JS1
    + JS2
    + JS3
    + JS4
    + JS_GLOSARIO
    + '\n</body>\n</html>'
)

out = 'dashboard_sistema_financiero.html'
with open(out, 'w', encoding='utf-8') as fh:
    fh.write(full_html)

print(f"\n✅  {out}  generado correctamente.")
print(f"   Bancos procesados : {len(todos_bancos)}")
print(f"   Periodos totales  : {len(periodos)}")
print(f"   Selector (12 meses): {plabel(periodos_12[0])} → {plabel(periodos_12[-1])}")
print(f"   Tamaño archivo    : {os.path.getsize(out)/1024:.1f} KB")
