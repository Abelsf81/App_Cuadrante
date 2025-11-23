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
from datetime import timedelta

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
        dates.add(end) 
    return dates

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
# 2. EL ARQUITECTO (GENERADOR ROBUSTO)
# -------------------------------------------------------------------

def check_architect_conflict(start_idx, duration, person, occupation_map, base_sch, year, transition_dates):
    total_days = len(base_sch['A'])
    if start_idx + duration > total_days: return True

    for i in range(start_idx, start_idx + duration):
        d_obj = datetime.date(year, 1, 1) + timedelta(days=i)
        if d_obj in transition_dates:
            if base_sch[person['Turno']][i] == 'T': return True
        
        occupants = occupation_map.get(i, [])
        if len(occupants) >= 2: return True
        
        for occ in occupants:
            if occ['Turno'] == person['Turno']: return True
            if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: return True
            
    return False

def book_architect_slot(start_idx, duration, person, occupation_map):
    for i in range(start_idx, start_idx + duration):
        if i not in occupation_map: occupation_map[i] = []
        occupation_map[i].append(person)

def run_architect_mode(roster_df, year, night_periods):
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    occupation_map = {} 
    generated_requests = []
    
    people = roster_df.to_dict('records')
    priority_order = ["Jefe", "Subjefe", "Conductor", "Bombero"]
    people.sort(key=lambda x: priority_order.index(x['Rol']))
    
    quarters = [(0, 90), (91, 181), (182, 273), (274, 364)]
    
    for person in people:
        credits_got = 0
        my_slots = []
        
        # FASE 1: BLOQUES LARGOS
        target_blocks = 3
        q_indices = [0, 1, 2, 3]
        random.shuffle(q_indices)
        
        for i in range(target_blocks):
            q_start, q_end = quarters[q_indices[i % 4]]
            candidates = []
            for d in range(q_start, q_end - 10):
                if not check_architect_conflict(d, 10, person, occupation_map, base_sch, year, transition_dates):
                    c = 0
                    for k in range(d, d+10):
                        if base_sch[person['Turno']][k] == 'T': c += 1
                    candidates.append((d, c))
            
            candidates.sort(key=lambda x: x[1], reverse=True)
            
            if candidates:
                chosen_start, gained_credits = candidates[0]
                overlap = False
                for s in my_slots:
                    if abs(chosen_start - s[0]) < 15: overlap = True
                
                if not overlap:
                    book_architect_slot(chosen_start, 10, person, occupation_map)
                    my_slots.append((chosen_start, 10))
                    credits_got += gained_credits
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=chosen_start),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=chosen_start + 9)
                    })

        # FASE 2: RELLENO DE PRECISIÓN
        attempts = 0
        while credits_got < 13 and attempts < 1000:
            d = random.randint(0, total_days - 1)
            if base_sch[person['Turno']][d] == 'T':
                if not check_architect_conflict(d, 1, person, occupation_map, base_sch, year, transition_dates):
                    book_architect_slot(d, 1, person, occupation_map)
                    credits_got += 1
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=d),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=d)
                    })
            attempts += 1
            
    return generated_requests

# -------------------------------------------------------------------
# 3. VISUALIZADOR HTML
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
                
                bg_color = "#eee"; text_color = "#ccc"; border = "1px solid #fff"
                if state == 'T': bg_color = "#d4edda"; text_color = "#155724"
                if is_in_night_period(d_idx, year, night_periods):
                    if state == 'T': bg_color = "#28a745"; text_color = "white"
                    else: bg_color = "#aaa"; text_color = "#555"
                if dt in get_night_transition_dates(night_periods): border = "2px solid red"

                html += f"<div style='width:20px; background-color:{bg_color}; color:{text_color}; text-align:center; border:{border}; border-radius:2px;'>{state}</div>"
            else:
                html += "<div style='width:20px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

# -------------------------------------------------------------------
# 4. FUNCIONES DE COBERTURA Y EXCEL
# -------------------------------------------------------------------

def get_candidates(person_missing, roster_df, day_idx, current_schedule, year, night_periods, adjustments_log_current_day=None):
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    
    blocked_turns = set()
    if adjustments_log_current_day:
        for coverer_name in adjustments_log_current_day:
            cov_p = roster_df[roster_df['Nombre'] == coverer_name]
            if not cov_p.empty:
                blocked_turns.add(cov_p.iloc[0]['Turno'])

    turn_exhausted_from_night = None
    if day_idx > 0:
        prev_day_idx = day_idx - 1
        if is_in_night_period(prev_day_idx, year, night_periods):
            base_sch_temp, _ = generate_base_schedule(year)
            for t in TEAMS:
                if base_sch_temp[t][prev_day_idx] == 'T':
                    turn_exhausted_from_night = t
                    break

    for _, candidate in roster_df.iterrows():
        if candidate['Turno'] == missing_turn: continue
        cand_status = current_schedule[candidate['Nombre']][day_idx]
        if cand_status != 'L': continue 
        
        if candidate['Turno'] in blocked_turns: continue
        if turn_exhausted_from_night and candidate['Turno'] == turn_exhausted_from_night: continue

        is_compatible = False
        cand_role = candidate['Rol']
        if missing_role == "Jefe" and cand_role in ["Jefe", "Subjefe"]: is_compatible = True
        elif missing_role == "Subjefe" and cand_role in ["Jefe", "Subjefe"]: is_compatible = True
        elif missing_role == "Conductor" and (cand_role == "Conductor" or candidate['SV']): is_compatible = True
        elif missing_role == "Bombero" and (cand_role == "Bombero" or candidate['SV']): is_compatible = True
            
        if is_compatible: candidates.append(candidate['Nombre'])
    return candidates

def validate_and_generate_final(roster_df, requests, year, night_periods):
    base_schedule_turn, total_days = generate_base_schedule(year)
    final_schedule = {} 
    turn_coverage_counters = {'A': 0, 'B': 0, 'C': 0}
    person_coverage_counters = {name: 0 for name in roster_df['Nombre']}
    name_to_turn = {row['Nombre']: row['Turno'] for _, row in roster_df.iterrows()}
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    natural_days_count = {name: 0 for name in roster_df['Nombre']}
    
    for req in requests:
        name = req['Nombre']
        s_idx = req['Inicio'].timetuple().tm_yday - 1
        e_idx = req['Fin'].timetuple().tm_yday - 1
        duration = (e_idx - s_idx) + 1
        natural_days_count[name] += duration
        for d in range(s_idx, e_idx + 1):
            if final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V'
            else:
                final_schedule[name][d] = 'V(L)'

    adjustments_log = []
    for d in range(total_days):
        absent_people = day_vacations[d]
        if not absent_people: continue
        current_day_coverers = []
        absent_people.sort(key=lambda x: 0 if "Jefe" in x or "Subjefe" in x else 1)

        for name_missing in absent_people:
            person_row = roster_df[roster_df['Nombre'] == name_missing].iloc[0]
            candidates = get_candidates(person_row, roster_df, d, final_schedule, year, night_periods, current_day_coverers)
            
            if candidates:
                valid = []
                for c in candidates:
                    prev = final_schedule[c][d-1] if d>0 else 'L'
                    prev2 = final_schedule[c][d-2] if d>1 else 'L'
                    if not (prev.startswith('T') and prev2.startswith('T')): valid.append(c)
                
                if valid:
                    valid.sort(key=lambda x: (turn_coverage_counters[name_to_turn[x]], person_coverage_counters[x], random.random()))
                    chosen = valid[0]
                    final_schedule[chosen][d] = f"T*({name_missing})"
                    adjustments_log.append((d, chosen, name_missing))
                    current_day_coverers.append(chosen)
                    turn_coverage_counters[name_to_turn[chosen]] += 1
                    person_coverage_counters[chosen] += 1

    fill_log = {}
    for name in roster_df['Nombre']:
        current = natural_days_count.get(name, 0)
        needed = 39 - current
        if needed > 0:
            available_idx = [i for i, x in enumerate(final_schedule[name]) if x == 'L']
            if len(available_idx) >= needed:
                fill_idxs = get_clustered_dates(available_idx, needed)
                added_dates = []
                for idx in fill_idxs:
                    final_schedule[name][idx] = 'V(R)'
                    d_obj = datetime.date(year, 1, 1) + datetime.timedelta(days=idx)
                    added_dates.append(d_obj)
                fill_log[name] = added_dates

    return final_schedule, adjustments_log, person_coverage_counters, fill_log

def create_final_excel(schedule, roster_df, year, requests, fill_log, counters, night_periods, adjustments_log):
    wb = Workbook()
    s_T = PatternFill("solid", fgColor="C6EFCE"); s_V = PatternFill("solid", fgColor="FFEB9C")
    s_VR = PatternFill("solid", fgColor="FFFFE0"); s_Cov = PatternFill("solid", fgColor="FFC7CE")
    s_L = PatternFill("solid", fgColor="F2F2F2"); s_Night = PatternFill("solid", fgColor="A6A6A6")
    font_bold = Font(bold=True); font_red = Font(color="9C0006", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_thin = Side(border_style="thin", color="000000")
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    ws1 = wb.active; ws1.title = "Cuadrante"
    ws1.column_dimensions['A'].width = 15
    for i in range(2, 34): ws1.column_dimensions[get_column_letter(i)].width = 4
    
    curr_row = 1
    for t in TEAMS:
        ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=32)
        cell_title = ws1.cell(curr_row, 1, f"TURNO {t}"); cell_title.font = Font(bold=True, color="FFFFFF"); cell_title.fill = PatternFill("solid", fgColor="000080"); cell_title.alignment = align_c
        curr_row += 2
        members = roster_df[roster_df['Turno'] == t]
        for _, p in members.iterrows():
            nm = p['Nombre']; role = p['Rol']
            ws1.cell(curr_row, 1, f"{nm} ({role})").font = font_bold
            for d in range(1, 32): c = ws1.cell(curr_row, d+1, d); c.alignment = align_c; c.font = font_bold; c.border = border_all; c.fill = PatternFill("solid", fgColor="E0E0E0")
            curr_row += 1
            for m_idx, mes in enumerate(MESES):
                ws1.cell(curr_row, 1, mes).font = font_bold; ws1.cell(curr_row, 1).border = border_all
                d_month = calendar.monthrange(year, m_idx+1)[1]
                for d in range(1, 32):
                    cell = ws1.cell(curr_row, d+1); cell.border = border_all; cell.alignment = align_c
                    if d <= d_month:
                        dt = datetime.date(year, m_idx+1, d); d_y = dt.timetuple().tm_yday - 1
                        st_val = schedule[nm][d_y]
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
    headers = ["Nombre", "Turno", "Puesto", "Gastado (T)", "Coberturas (T*)", "Total Vacs (Nat)"]
    ws2.append(headers)
    for _, p in roster_df.iterrows():
        name = p['Nombre']; sch = schedule[name]
        v_credits = sch.count('V'); t_cover = counters[name]
        v_natural = sch.count('V') + sch.count('V(L)') + sch.count('V(R)')
        ws2.append([name, p['Turno'], p['Rol'], v_credits, t_cover, v_natural])

    ws4 = wb.create_sheet("Ajustes")
    ws4.append(["Fecha", "Cubre", "Ausente"])
    for d, c, a in adjustments_log:
        dt = datetime.date(year, 1, 1) + datetime.timedelta(days=d)
        ws4.append([dt.strftime("%d/%m/%Y"), c, a])
    
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ STREAMLIT (V14.1 - ARQUITECTO CORREGIDO)
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V14.1 - Arquitecto")

st.title("🚒 Gestor V14.1: El Arquitecto")
st.caption("Sistema de Construcción de Cuadrantes desde Cero.")

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

    # --- BOTÓN DE GENERACIÓN AUTOMÁTICA (EL ARQUITECTO) ---
    st.divider()
    st.markdown("#### 🏗️ Construcción Automática")
    if st.button("Construir Cuadrante Completo (13 Cr)", type="primary"):
        with st.spinner("El Arquitecto está diseñando el año..."):
            new_reqs = run_architect_mode(edited_df, year_val, st.session_state.nights)
            st.session_state.raw_requests_df = pd.DataFrame(new_reqs)
        st.success(f"¡Construido! {len(new_reqs)} periodos creados.")
        st.rerun()

# 2. ESTADO DE SOLICITUDES
if 'raw_requests_df' not in st.session_state:
    cols = ["Nombre", "Inicio", "Fin"]
    st.session_state.raw_requests_df = pd.DataFrame(columns=cols)

# 3. ZONA INTERACTIVA
st.divider()
st.subheader("2. Mesa de Trabajo Interactiva")

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
stats = calculate_stats(edited_df, current_requests, year_val)

with c_stats:
    st.markdown("### 📊 Estado de Créditos")
    st.divider()
    
    for name, data in stats.items():
        c = data['credits']
        col = "green" if c == 13 else "red"
        
        with st.expander(f"{name}: {c}/13", expanded=(c!=13)):
            if c < 13:
                st.warning(f"Faltan {13-c} créditos.")
            elif c > 13:
                st.error(f"Sobran {c-13} créditos.")

# 5. GENERACIÓN FINAL
st.divider()
if st.button("🚀 Generar Excel Final", type="primary", use_container_width=True):
    sch, adj, count, fill = validate_and_generate_final(edited_df, current_requests, year_val, st.session_state.nights)
    excel_io = create_final_excel(sch, edited_df, year_val, current_requests, fill, count, st.session_state.nights, adj)
    st.download_button("📥 Descargar Cuadrante Validado", excel_io, f"Cuadrante_Final_{year_val}.xlsx")
