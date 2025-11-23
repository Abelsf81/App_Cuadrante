import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import io
import random
import calendar 
import pandas as pd
from itertools import groupby
from operator import itemgetter

# --- CONSTANTES Y CONFIGURACIÓN ---
TEAMS = ['A', 'B', 'C']
ROLES = ["Jefe", "Subjefe", "Conductor", "Bombero"] 
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Plantilla por defecto
DEFAULT_ROSTER = [
    {"ID_Puesto": "Jefe A",       "Nombre": "Jefe A",       "Turno": "A", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe A",    "Nombre": "Subjefe A",    "Turno": "A", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond A",       "Nombre": "Cond A",       "Turno": "A", "Rol": "Conductor",  "SV": True},
    {"ID_Puesto": "Bombero A1",   "Nombre": "Bombero A1",   "Turno": "A", "Rol": "Bombero",    "SV": True},
    {"ID_Puesto": "Bombero A2",   "Nombre": "Bombero A2",   "Turno": "A", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero A3",   "Nombre": "Bombero A3",   "Turno": "A", "Rol": "Bombero",    "SV": False},
    
    {"ID_Puesto": "Jefe B",       "Nombre": "Jefe B",       "Turno": "B", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe B",    "Nombre": "Subjefe B",    "Turno": "B", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond B",       "Nombre": "Cond B",       "Turno": "B", "Rol": "Conductor",  "SV": True},
    {"ID_Puesto": "Bombero B1",   "Nombre": "Bombero B1",   "Turno": "B", "Rol": "Bombero",    "SV": True},
    {"ID_Puesto": "Bombero B2",   "Nombre": "Bombero B2",   "Turno": "B", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero B3",   "Nombre": "Bombero B3",   "Turno": "B", "Rol": "Bombero",    "SV": False},

    {"ID_Puesto": "Jefe C",       "Nombre": "Jefe C",       "Turno": "C", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe C",    "Nombre": "Subjefe C",    "Turno": "C", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond C",       "Nombre": "Cond C",       "Turno": "C", "Rol": "Conductor",  "SV": True},
    {"ID_Puesto": "Bombero C1",   "Nombre": "Bombero C1",   "Turno": "C", "Rol": "Bombero",    "SV": True},
    {"ID_Puesto": "Bombero C2",   "Nombre": "Bombero C2",   "Turno": "C", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero C3",   "Nombre": "Bombero C3",   "Turno": "C", "Rol": "Bombero",    "SV": False},
]

# -------------------------------------------------------------------
# 1. LÓGICA BASE
# -------------------------------------------------------------------

def generate_base_schedule(year):
    is_leap = calendar.isleap(year)
    total_days = 366 if is_leap else 365
    status = {'A': 0, 'B': 2, 'C': 1} 
    schedule = {team: [] for team in TEAMS}
    for _ in range(total_days):
        for t in TEAMS:
            if status[t] == 0: schedule[t].append('T')
            else: schedule[t].append('L')
            status[t] = (status[t] + 1) % 3
    return schedule, total_days

def is_in_night_period(day_idx, year, night_periods):
    current_date = datetime.date(year, 1, 1) + datetime.timedelta(days=day_idx)
    for start, end in night_periods:
        if start <= current_date <= end: return True
    return False

def get_night_transition_dates(night_periods):
    dates = set()
    for start, end in night_periods:
        dates.add(start); dates.add(end)
    return dates

def get_candidates(person_missing, roster_df, day_idx, current_schedule, adjustments_log_current_day=None):
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    
    blocked_turns = set()
    if adjustments_log_current_day:
        for coverer_name in adjustments_log_current_day:
            cov_p = roster_df[roster_df['Nombre'] == coverer_name]
            if not cov_p.empty:
                blocked_turns.add(cov_p.iloc[0]['Turno'])

    for _, candidate in roster_df.iterrows():
        if candidate['Turno'] == missing_turn: continue
        cand_status = current_schedule[candidate['Nombre']][day_idx]
        if cand_status != 'L': continue 
        if candidate['Turno'] in blocked_turns: continue

        is_compatible = False
        cand_role = candidate['Rol']
        
        if missing_role == "Jefe":
            if cand_role in ["Jefe", "Subjefe"]: is_compatible = True
        elif missing_role == "Subjefe":
             if cand_role in ["Jefe", "Subjefe"]: is_compatible = True
        elif missing_role == "Conductor":
            if cand_role == "Conductor": is_compatible = True
            if candidate['SV']: is_compatible = True
        elif missing_role == "Bombero":
            if cand_role == "Bombero": is_compatible = True
            if candidate['SV']: is_compatible = True
            
        if is_compatible:
            candidates.append(candidate['Nombre'])
    return candidates

# -------------------------------------------------------------------
# 2. VISUALIZADOR HTML
# -------------------------------------------------------------------
def render_annual_calendar(year, team, base_sch, night_periods):
    html = f"<div style='font-family:monospace; font-size:10px;'>"
    html += "<div style='display:flex; margin-bottom:2px;'><div style='width:30px;'></div>"
    for d in range(1, 32):
        html += f"<div style='width:20px; text-align:center; color:#888;'>{d}</div>"
    html += "</div>"

    for m_idx, mes in enumerate(MESES):
        m_num = m_idx + 1
        days_in_month = calendar.monthrange(year, m_num)[1]
        html += f"<div style='display:flex; margin-bottom:2px;'><div style='width:30px; font-weight:bold;'>{mes}</div>"
        
        for d in range(1, 32):
            if d <= days_in_month:
                dt = datetime.date(year, m_num, d)
                d_idx = dt.timetuple().tm_yday - 1
                state = base_sch[team][d_idx]
                
                bg_color = "#eee"
                text_color = "#ccc"
                border = "1px solid #fff"
                
                if state == 'T':
                    bg_color = "#d4edda"; text_color = "#155724"
                
                if is_in_night_period(d_idx, year, night_periods):
                    if state == 'T': bg_color = "#28a745"; text_color = "white"
                    else: bg_color = "#aaa"; text_color = "#555"
                
                if dt in get_night_transition_dates(night_periods):
                    border = "2px solid red"

                html += f"<div style='width:20px; background-color:{bg_color}; color:{text_color}; text-align:center; border:{border}; border-radius:2px;'>{state}</div>"
            else:
                html += "<div style='width:20px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

# -------------------------------------------------------------------
# 3. LÓGICA INTERACTIVA Y SUGERENCIAS
# -------------------------------------------------------------------

def find_valid_slot(person_name, duration, roster_df, year, night_periods, current_requests):
    """Busca un hueco válido cercano para sugerir."""
    base_sch, total_days = generate_base_schedule(year)
    person = roster_df[roster_df['Nombre'] == person_name].iloc[0]
    
    occ_map = {i:[] for i in range(total_days)}
    for req in current_requests:
        p_req = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0]
        s = req['Inicio'].timetuple().tm_yday - 1
        e = req['Fin'].timetuple().tm_yday - 1
        for d in range(s, e+1):
            if base_sch[p_req['Turno']][d] == 'T': occ_map[d].append(p_req)

    all_days = list(range(total_days))
    random.shuffle(all_days)
    
    for start_idx in all_days:
        end_idx = start_idx + duration - 1
        if end_idx >= total_days: continue
        
        valid_block = True
        credits_in_block = 0
        
        for d in range(start_idx, end_idx + 1):
            if base_sch[person['Turno']][d] == 'T':
                credits_in_block += 1
                occupants = occ_map[d]
                if len(occupants) >= 2: valid_block = False; break
                for occ in occupants:
                    if occ['Turno'] == person['Turno']: valid_block = False; break
                    if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: valid_block = False; break
                d_obj = datetime.date(year, 1, 1) + datetime.timedelta(days=d)
                if d_obj in get_night_transition_dates(night_periods): valid_block = False; break
        
        if valid_block and credits_in_block > 0:
            return datetime.date(year, 1, 1) + datetime.timedelta(days=start_idx)
            
    return None

def calculate_stats(roster_df, requests, year):
    base_sch, _ = generate_base_schedule(year)
    stats = {}
    for _, p in roster_df.iterrows():
        stats[p['Nombre']] = {'credits': 0, 'natural': 0}
        
    for req in requests:
        name = req['Nombre']
        if name not in stats: continue
        s_idx = req['Inicio'].timetuple().tm_yday - 1
        e_idx = req['Fin'].timetuple().tm_yday - 1
        row = roster_df[roster_df['Nombre'] == name].iloc[0]
        
        nat = (e_idx - s_idx) + 1
        cred = 0
        for d in range(s_idx, e_idx + 1):
            if base_sch[row['Turno']][d] == 'T': cred += 1
        stats[name]['credits'] += cred
        stats[name]['natural'] += nat
    return stats

# -------------------------------------------------------------------
# 4. GENERACIÓN FINAL (EXCEL)
# -------------------------------------------------------------------
def create_final_excel(roster_df, requests, year, night_periods):
    base_sch, total_days = generate_base_schedule(year)
    final_schedule = {row['Nombre']: base_sch[row['Turno']].copy() for _, row in roster_df.iterrows()}
    counters = {row['Nombre']: 0 for _, row in roster_df.iterrows()}
    name_to_turn = {row['Nombre']: row['Turno'] for _, row in roster_df.iterrows()}
    turn_coverage = {'A':0, 'B':0, 'C':0}

    day_vacs = {i:[] for i in range(total_days)}
    for r in requests:
        nm = r['Nombre']
        s = r['Inicio'].timetuple().tm_yday - 1
        e = r['Fin'].timetuple().tm_yday - 1
        for d in range(s, e+1):
            if final_schedule[nm][d] == 'T':
                final_schedule[nm][d] = 'V'
                day_vacs[d].append(nm)
            else:
                final_schedule[nm][d] = 'V(L)'
                
    adjustments_log = []
    for d in range(total_days):
        absent = day_vacs[d]
        if not absent: continue
        absent.sort(key=lambda x: 0 if "Jefe" in x or "Subjefe" in x else 1)
        coverers_today_turns = set()
        
        for missing in absent:
            p_miss = roster_df[roster_df['Nombre'] == missing].iloc[0]
            cands = []
            for _, cand in roster_df.iterrows():
                if cand['Turno'] == p_miss['Turno']: continue
                if final_schedule[cand['Nombre']][d] != 'L': continue
                if cand['Turno'] in coverers_today_turns: continue
                
                ok = False
                if p_miss['Rol'] == 'Jefe' and cand['Rol'] in ['Jefe', 'Subjefe']: ok=True
                elif p_miss['Rol'] == 'Subjefe' and cand['Rol'] in ['Jefe', 'Subjefe']: ok=True
                elif p_miss['Rol'] == 'Conductor' and (cand['Rol']=='Conductor' or cand['SV']): ok=True
                elif p_miss['Rol'] == 'Bombero' and (cand['Rol']=='Bombero' or cand['SV']): ok=True
                if ok: cands.append(cand['Nombre'])
            
            valid_c = []
            for c in cands:
                p1 = final_schedule[c][d-1] if d>0 else 'L'
                p2 = final_schedule[c][d-2] if d>1 else 'L'
                if not (p1.startswith('T') and p2.startswith('T')): valid_c.append(c)
                
            if valid_c:
                valid_c.sort(key=lambda x: (turn_coverage[name_to_turn[x]], counters[x], random.random()))
                chosen = valid_c[0]
                final_schedule[chosen][d] = f"T*({missing})"
                adjustments_log.append((d, chosen, missing))
                counters[chosen] += 1
                turn_coverage[name_to_turn[chosen]] += 1
                coverers_today_turns.add(name_to_turn[chosen])

    wb = Workbook()
    s_T = PatternFill("solid", fgColor="C6EFCE"); s_V = PatternFill("solid", fgColor="FFEB9C")
    s_VR = PatternFill("solid", fgColor="FFFFE0"); s_Cov = PatternFill("solid", fgColor="FFC7CE")
    s_L = PatternFill("solid", fgColor="F2F2F2"); s_Night = PatternFill("solid", fgColor="A6A6A6")
    font_bold = Font(bold=True); font_red = Font(color="9C0006", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_thin = Side(border_style="thin", color="000000")
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    ws1 = wb.active; ws1.title = "Cuadrante"
    ws1.column_dimensions['A'].width = 20
    for i in range(2, 34): ws1.column_dimensions[get_column_letter(i)].width = 4
    
    curr_row = 1
    for t in TEAMS:
        ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=32)
        cell_title = ws1.cell(curr_row, 1, f"TURNO {t}"); cell_title.font = Font(bold=True, color="FFFFFF"); cell_title.fill = PatternFill("solid", fgColor="000080"); cell_title.alignment = align_c
        curr_row += 2
        members = roster_df[roster_df['Turno'] == t]
        for _, p in members.iterrows():
            nm = p['Nombre']
            ws1.cell(curr_row, 1, f"{nm} ({p['Rol']})").font = font_bold
            for d in range(1, 32): c = ws1.cell(curr_row, d+1, d); c.alignment = align_c; c.font = font_bold; c.border = border_all; c.fill = PatternFill("solid", fgColor="E0E0E0")
            curr_row += 1
            for m_idx, mes in enumerate(MESES):
                ws1.cell(curr_row, 1, mes).font = font_bold; ws1.cell(curr_row, 1).border = border_all
                d_month = calendar.monthrange(year, m_idx+1)[1]
                for d in range(1, 32):
                    cell = ws1.cell(curr_row, d+1); cell.border = border_all; cell.alignment = align_c
                    if d <= d_month:
                        dt = datetime.date(year, m_idx+1, d); d_y = dt.timetuple().tm_yday - 1
                        st_val = final_schedule[nm][d_y]
                        fill = s_L; val = ""
                        if st_val == 'T': fill = s_T; val = "T"
                        elif st_val == 'V': fill = s_V; val = "V"
                        elif st_val.startswith('V('): fill = s_VR; val = "v"
                        elif st_val.startswith('T*'): fill = s_Cov; val = "*"; cell.font = font_red
                        if is_in_night_period(d_y, year, night_periods): fill = s_Night
                        cell.fill = fill; cell.value = val
                    else: cell.fill = PatternFill("solid", fgColor="808080")
                curr_row += 1
            curr_row += 2

    ws2 = wb.create_sheet("Estadísticas")
    ws2.append(["Nombre", "Turno", "Puesto", "Gastado (T)", "Coberturas", "Total (Nat)"])
    for _, p in roster_df.iterrows():
        nm = p['Nombre']; sch = final_schedule[nm]
        cred = sch.count('V'); cov = counters[nm]
        nat = cred + sch.count('V(L)') + sch.count('V(R)')
        ws2.append([nm, p['Turno'], p['Rol'], cred, cov, nat])

    ws4 = wb.create_sheet("Ajustes")
    ws4.append(["Fecha", "Cubre", "Ausente"])
    for d, c, a in adjustments_log:
        dt = datetime.date(year, 1, 1) + datetime.timedelta(days=d)
        ws3 = wb.create_sheet("Log") if 'Log' not in wb.sheetnames else wb["Log"]
        ws4.append([dt.strftime("%d/%m/%Y"), c, a])

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ STREAMLIT (V9.2 - INTERACTIVO + AUTO-EQUILIBRIO)
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V9.2 - Interactivo")

st.title("🚒 Gestor V9.2: Simulador Interactivo")
st.caption("Sube tus datos, visualiza conflictos y corrígelos en tiempo real.")

# 1. CONFIGURACIÓN INICIAL
with st.sidebar:
    st.header("1. Configuración")
    year_val = st.number_input("Año", value=2026)
    
    with st.expander("Plantilla"):
        if 'roster_data' not in st.session_state:
            st.session_state.roster_data = pd.DataFrame(DEFAULT_ROSTER)
        edited_df = st.data_editor(st.session_state.roster_data, use_container_width=True)
        st.session_state.roster_data = edited_df
        
    with st.expander("Nocturnas"):
        if 'nights' not in st.session_state: st.session_state.nights = []
        c1, c2 = st.columns(2)
        dn_s = c1.date_input("Inicio", key="n_s", value=None)
        dn_e = c2.date_input("Fin", key="n_e", value=None)
        if st.button("Añadir Nocturna"):
            if dn_s and dn_e: st.session_state.nights.append((dn_s, dn_e))
        
        st.write(f"Periodos: {len(st.session_state.nights)}")
        
        st.divider()
        st.caption("O sube Excel (Inicio | Fin)")
        uploaded_n = st.file_uploader("Excel Nocturnas", type=['xlsx'], key="n_up")
        if uploaded_n:
            try:
                df_n = pd.read_excel(uploaded_n)
                c = 0
                for _, row in df_n.iterrows():
                    try:
                        v1 = row.iloc[0]; v2 = row.iloc[1]
                        if not pd.isnull(v1) and not pd.isnull(v2):
                            d1 = pd.to_datetime(v1).date()
                            d2 = pd.to_datetime(v2).date()
                            st.session_state.nights.append((d1, d2)); c+=1
                    except: pass
                if c>0: st.success(f"Añadidos {c} periodos.")
            except: pass

        if st.button("Limpiar Nocturnas"): st.session_state.nights = []

# 2. ESTADO DE SOLICITUDES
if 'raw_requests_df' not in st.session_state:
    cols = ["Nombre", "Inicio", "Fin"]
    st.session_state.raw_requests_df = pd.DataFrame(columns=cols)

# 3. CARGA DE DATOS
st.divider()
c_up, c_info = st.columns([1, 2])

with c_up:
    uploaded_file = st.file_uploader("📂 Cargar Excel (Formato Horizontal)", type=['xlsx'])
    if uploaded_file:
        if st.button("Procesar Carga"):
            try:
                df = pd.read_excel(uploaded_file)
                new_reqs = []
                for _, row in df.iterrows():
                    name = row.get('Nombre')
                    if not name: continue
                    for i in range(1, 21):
                        ks = f"Inicio {i}"; ke = f"Fin {i}"
                        if ks in row and ke in row and not pd.isnull(row[ks]):
                            try:
                                new_reqs.append({
                                    "Nombre": name,
                                    "Inicio": pd.to_datetime(row[ks]).date(),
                                    "Fin": pd.to_datetime(row[ke]).date()
                                })
                            except: pass
                st.session_state.raw_requests_df = pd.DataFrame(new_reqs)
                st.success(f"Cargados {len(new_reqs)} registros.")
            except Exception as e: st.error(f"Error: {e}")

# 4. ZONA INTERACTIVA
st.subheader("2. Mesa de Trabajo Interactiva")

# Visualizar calendario
base_sch, total_days = generate_base_schedule(year_val)
tab_a, tab_b, tab_c = st.tabs(["Calendario A", "Calendario B", "Calendario C"])
with tab_a: st.markdown(render_annual_calendar(year_val, 'A', base_sch, st.session_state.nights), unsafe_allow_html=True)
with tab_b: st.markdown(render_annual_calendar(year_val, 'B', base_sch, st.session_state.nights), unsafe_allow_html=True)
with tab_c: st.markdown(render_annual_calendar(year_val, 'C', base_sch, st.session_state.nights), unsafe_allow_html=True)

# Lógica
current_requests = st.session_state.raw_requests_df.to_dict('records')
c_editor, c_stats = st.columns([2, 1])

with c_editor:
    st.markdown("### ✏️ Edición de Solicitudes")
    df_editor = st.session_state.raw_requests_df.copy()
    if not df_editor.empty:
        df_editor["Inicio"] = pd.to_datetime(df_editor["Inicio"])
        df_editor["Fin"] = pd.to_datetime(df_editor["Fin"])
    
    edited_requests_df = st.data_editor(df_editor, num_rows="dynamic", use_container_width=True, key="editor_main")
    
    if not edited_requests_df.empty:
        final_reqs_list = []
        for _, r in edited_requests_df.iterrows():
            try:
                final_reqs_list.append({
                    "Nombre": r["Nombre"],
                    "Inicio": r["Inicio"].date(),
                    "Fin": r["Fin"].date()
                })
            except: pass 
        current_requests = final_reqs_list

# Stats
stats = calculate_spent_credits(edited_df, current_requests, year_val)

with c_stats:
    st.markdown("### 📊 Asesor Inteligente")
    
    # --- BOTÓN MAESTRO: EQUILIBRADO GLOBAL ---
    if st.button("⚖️ Auto-Equilibrar Todo (Recortar >13 y Rellenar <13)", type="primary"):
        # 1. Fase de Recorte (>13)
        cleaned_reqs = []
        # Agrupar por nombre
        reqs_by_name = {}
        for r in current_requests:
            reqs_by_name.setdefault(r['Nombre'], []).append(r)
            
        temp_reqs = []
        # Procesar recortes
        for name, reqs in reqs_by_name.items():
            cred = stats.get(name, {}).get('credits', 0)
            # Ordenar cronológicamente
            reqs.sort(key=lambda x: x['Inicio'])
            
            while cred > 13 and reqs:
                # Recortar del último
                last_req = reqs[-1]
                # Calcular cuantos creditos vale el ultimo dia
                last_day = last_req['Fin']
                # Check if last day is T
                p_data = edited_df[edited_df['Nombre']==name].iloc[0]
                d_idx = last_day.timetuple().tm_yday - 1
                
                if base_sch[p_data['Turno']][d_idx] == 'T':
                    cred -= 1 # Hemos quitado un credito
                
                if last_req['Inicio'] == last_req['Fin']:
                    reqs.pop() # Eliminar periodo completo
                else:
                    last_req['Fin'] = last_day - datetime.timedelta(days=1)
            
            temp_reqs.extend(reqs)
            
        # Actualizar lista intermedia
        current_requests = temp_reqs
        # Recalcular stats tras recorte
        stats = calculate_spent_credits(edited_df, current_requests, year_val)
        
        # 2. Fase de Relleno (<13)
        added_count = 0
        for name, data in stats.items():
            credits_needed = 13 - data['credits']
            if credits_needed > 0:
                # Buscar N días
                for _ in range(credits_needed):
                    slot = find_valid_slot(name, 1, edited_df, year_val, st.session_state.nights, current_requests)
                    if slot:
                        current_requests.append({"Nombre": name, "Inicio": slot, "Fin": slot})
                        added_count += 1
        
        # Guardar resultado final
        df_update = pd.DataFrame(current_requests)
        df_update['Inicio'] = pd.to_datetime(df_update['Inicio'])
        df_update['Fin'] = pd.to_datetime(df_update['Fin'])
        st.session_state.raw_requests_df = df_update
        st.success(f"✅ Proceso completado. Se añadieron {added_count} días.")
        st.rerun()

    st.divider()
    
    # Estado Individual
    for name, data in stats.items():
        c = data['credits']
        col = "green" if c == 13 else "red"
        with st.expander(f"{name}: {c}/13", expanded=(c!=13)):
            if c < 13:
                if st.button(f"➕ Buscar 1 día ({name})", key=f"add_{name}"):
                    slot = find_valid_slot(name, 1, edited_df, year_val, st.session_state.nights, current_requests)
                    if slot:
                        current_requests.append({"Nombre": name, "Inicio": slot, "Fin": slot})
                        st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                        st.rerun()
                    else: st.error("Sin hueco.")
            elif c > 13:
                if st.button(f"✂️ Recortar 1 día ({name})", key=f"trim_{name}"):
                    # Lógica manual de recorte (mismo que arriba pero individual)
                    df_t = pd.DataFrame(current_requests)
                    idx_list = df_t[df_t['Nombre']==name].index.tolist()
                    if idx_list:
                        last = idx_list[-1]
                        r = df_t.loc[last]
                        if r['Inicio'] == r['Fin']: df_t = df_t.drop(last)
                        else: df_t.at[last, 'Fin'] = r['Fin'] - datetime.timedelta(days=1)
                        st.session_state.raw_requests_df = df_t
                        st.rerun()

# 5. GENERACIÓN FINAL
st.divider()
if st.button("🚀 Generar Excel Final (Solo si todo está OK)", type="primary", use_container_width=True):
    excel_io = create_final_excel(edited_df, current_requests, year_val, st.session_state.nights)
    st.download_button("📥 Descargar Cuadrante Validado", excel_io, f"Cuadrante_Final_{year_val}.xlsx")
