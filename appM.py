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
# 2. MOTOR DE DRAFT (LA IA DE SELECCIÓN)
# -------------------------------------------------------------------

def get_available_blocks_for_person(person_name, roster_df, current_requests, year, night_periods, month_range):
    """
    Genera TODAS las opciones válidas restantes para una persona.
    """
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    person = roster_df[roster_df['Nombre'] == person_name].iloc[0]
    
    # Filtrar meses
    start_month_idx = MESES.index(month_range[0]) + 1
    end_month_idx = MESES.index(month_range[1]) + 1
    
    # 1. Construir Mapa de Ocupación Actual (Lo que ya está cogido)
    occupation_map = {i:[] for i in range(total_days)}
    my_current_slots = [] 

    for req in current_requests:
        if req['Nombre'] != person_name:
            p_req = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0]
            s = req['Inicio'].timetuple().tm_yday - 1
            e = req['Fin'].timetuple().tm_yday - 1
            for d in range(s, e+1):
                if base_sch[p_req['Turno']][d] == 'T': occupation_map[d].append(p_req)
        else:
            s = req['Inicio'].timetuple().tm_yday - 1
            e = req['Fin'].timetuple().tm_yday - 1
            my_current_slots.append((s, e))

    options = {'gold': [], 'silver': [], 'bronze': []}
    
    for d in range(total_days - 10): 
        # Filtro fecha inicio por mes seleccionado
        d_date = datetime.date(year, 1, 1) + timedelta(days=d)
        if not (start_month_idx <= d_date.month <= end_month_idx): continue

        # --- ANALISIS DE BLOQUE DE 10 DÍAS ---
        duration = 10
        credits = 0
        valid = True
        
        for k in range(d, d+duration):
            if base_sch[person['Turno']][k] == 'T':
                credits += 1
                d_obj = datetime.date(year, 1, 1) + timedelta(days=k)
                if d_obj in transition_dates: valid = False; break
                
                occupants = occupation_map[k]
                if len(occupants) >= 2: valid = False; break
                for occ in occupants:
                    if occ['Turno'] == person['Turno']: valid = False; break
                    if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: valid = False; break
            
            for ms in my_current_slots:
                if not (k < ms[0] - 2 or k > ms[1] + 2): valid = False; break
            
            if not valid: break
        
        if valid:
            start_date = datetime.date(year, 1, 1) + timedelta(days=d)
            end_date = start_date + timedelta(days=duration-1)
            label = f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
            
            if credits == 4:
                options['gold'].append({'label': label, 'start': start_date, 'end': end_date, 'cr': 4})
            elif credits == 3:
                options['silver'].append({'label': label, 'start': start_date, 'end': end_date, 'cr': 3})

        # --- ANALISIS DE BLOQUE DE 9 DÍAS ---
        duration = 9
        credits = 0
        valid = True
        for k in range(d, d+duration):
            if base_sch[person['Turno']][k] == 'T':
                credits += 1
                d_obj = datetime.date(year, 1, 1) + timedelta(days=k)
                if d_obj in transition_dates: valid = False; break
                occupants = occupation_map[k]
                if len(occupants) >= 2: valid = False; break
                for occ in occupants:
                    if occ['Turno'] == person['Turno']: valid = False; break
                    if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: valid = False; break
            for ms in my_current_slots:
                if not (k < ms[0] - 2 or k > ms[1] + 2): valid = False; break
            if not valid: break
            
        if valid and credits == 3: 
             start_date = datetime.date(year, 1, 1) + timedelta(days=d)
             end_date = start_date + timedelta(days=duration-1)
             label = f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
             options['bronze'].append({'label': label, 'start': start_date, 'end': end_date, 'cr': 3})

    return options

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
# 4. GENERACIÓN FINAL (EXCEL)
# -------------------------------------------------------------------
def get_candidates(person_missing, roster_df, day_idx, current_schedule, year, night_periods, adjustments_log_current_day=None):
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    
    blocked_turns = set()
    if adjustments_log_current_day:
        for coverer_name in adjustments_log_current_day:
            cov_p = roster_df[roster_df['Nombre'] == coverer_name]
            if not cov_p.empty: blocked_turns.add(cov_p.iloc[0]['Turno'])

    turn_exhausted_from_night = None
    if day_idx > 0:
        prev_day_idx = day_idx - 1
        if is_in_night_period(prev_day_idx, year, night_periods):
            base_sch_temp, _ = generate_base_schedule(year)
            for t in TEAMS:
                if base_sch_temp[t][prev_day_idx] == 'T':
                    turn_exhausted_from_night = t; break

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

def create_final_excel(roster_df, requests, year, night_periods):
    base_sch, total_days = generate_base_schedule(year)
    final_schedule = {row['Nombre']: base_sch[row['Turno']].copy() for _, row in roster_df.iterrows()}
    counters = {row['Nombre']: 0 for _, row in roster_df.iterrows()}
    name_to_turn = {row['Nombre']: row['Turno'] for _, row in roster_df.iterrows()}
    turn_coverage = {'A':0, 'B':0, 'C':0}

    day_vacs = {i:[] for i in range(total_days)}
    natural_days_count = {name: 0 for name in roster_df['Nombre']}
    
    for r in requests:
        nm = r['Nombre']
        s = r['Inicio'].timetuple().tm_yday - 1
        e = r['Fin'].timetuple().tm_yday - 1
        natural_days_count[nm] += (e - s + 1)
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
            cands = get_candidates(p_miss, roster_df, d, final_schedule, year, night_periods, list(coverers_today_turns))
            
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

    # Relleno Visual
    for name in roster_df['Nombre']:
        needed = 39 - natural_days_count.get(name, 0)
        if needed > 0:
            avl = [i for i, x in enumerate(final_schedule[name]) if x == 'L']
            if len(avl) >= needed:
                for i in range(needed): final_schedule[name][avl[i]] = 'V(R)'

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
    headers = ["Nombre", "Turno", "Puesto", "Gastado (T)", "Coberturas (T*)", "Total Vacs (Nat)"]
    ws2.append(headers)
    for _, p in roster_df.iterrows():
        nm = p['Nombre']; sch = final_schedule[nm]
        cred = sch.count('V'); cov = counters[nm]
        nat = cred + sch.count('V(L)') + sch.count('V(R)')
        ws2.append([nm, p['Turno'], p['Rol'], cred, cov, nat])

    ws4 = wb.create_sheet("Ajustes")
    ws4.append(["Fecha", "Cubre", "Ausente"])
    for d, c, a in adjustments_log:
        dt = datetime.date(year, 1, 1) + datetime.timedelta(days=d)
        ws4.append([dt.strftime("%d/%m/%Y"), c, a])
    
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ STREAMLIT (V15.1 - SALA DE DRAFT CON FILTRO MES Y SCROLL)
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V15.1 - Sala de Draft")

st.title("🚒 Gestor V15.1: Sala de Selección (Draft)")
st.caption("Selecciona trabajador y elige fichas. Usa el filtro de meses para ver más opciones.")

# 1. CONFIGURACIÓN
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
                            d1 = pd.to_datetime(v1).date(); d2 = pd.to_datetime(v2).date()
                            st.session_state.nights.append((d1, d2)); c+=1
                    except: pass
                if c>0: st.success(f"Cargadas {c}")
            except: pass
        if st.button("Limpiar Nocturnas"): st.session_state.nights = []

# 2. ESTADO
if 'raw_requests_df' not in st.session_state:
    st.session_state.raw_requests_df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
current_requests = st.session_state.raw_requests_df.to_dict('records')
stats = calculate_stats(edited_df, current_requests, year_val)

# 3. DRAFT ROOM
st.divider()
c_main, c_vis = st.columns([1, 2])

with c_main:
    st.subheader("2. Selección de Personal")
    
    all_names = edited_df['Nombre'].tolist()
    names_sorted = sorted(all_names, key=lambda x: (
        0 if "Jefe" in x else 1 if "Subjefe" in x else 2 if "Cond" in x else 3
    ))
    
    selected_person = st.selectbox("Selecciona Trabajador:", names_sorted)
    
    if selected_person:
        st.markdown("---")
        curr_stats = stats.get(selected_person, {'credits': 0, 'natural': 0})
        c = curr_stats['credits']
        remaining = 13 - c
        
        st.metric("Créditos Gastados", f"{c} / 13", delta=remaining, delta_color="normal")
        
        if remaining <= 0:
            st.success("✅ Cupo cubierto.")
        else:
            # FILTRO DE MESES
            month_range = st.select_slider("📅 Filtrar sugerencias por meses:", options=MESES, value=(MESES[0], MESES[-1]))
            
            st.info(f"🔍 Buscando bloques de {month_range[0]} a {month_range[1]}...")
            
            options = get_available_blocks_for_person(selected_person, edited_df, current_requests, year_val, st.session_state.nights, month_range)
            
            t_gold, t_silver, t_bronze = st.tabs([
                "🏆 Gold (4 Cr)", 
                "🥈 Silver (3 Cr - 10d)", 
                "🥉 Bronze (3 Cr - 9d)"
            ])
            
            with t_gold:
                with st.container(height=300):
                    if not options['gold']: st.write("Sin opciones.")
                    for opt in options['gold']:
                        if st.button(f"➕ {opt['label']}", key=f"g_{selected_person}_{opt['start']}"):
                            current_requests.append({"Nombre": selected_person, "Inicio": opt['start'], "Fin": opt['end']})
                            st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                            st.rerun()
                            
            with t_silver:
                with st.container(height=300):
                    if not options['silver']: st.write("Sin opciones.")
                    for opt in options['silver']:
                        if st.button(f"➕ {opt['label']}", key=f"s_{selected_person}_{opt['start']}"):
                            current_requests.append({"Nombre": selected_person, "Inicio": opt['start'], "Fin": opt['end']})
                            st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                            st.rerun()
                            
            with t_bronze:
                with st.container(height=300):
                    if not options['bronze']: st.write("Sin opciones.")
                    for opt in options['bronze']:
                        if st.button(f"➕ {opt['label']}", key=f"b_{selected_person}_{opt['start']}"):
                            current_requests.append({"Nombre": selected_person, "Inicio": opt['start'], "Fin": opt['end']})
                            st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                            st.rerun()

    st.markdown("---")
    st.write(f"**Mis Periodos ({selected_person}):**")
    my_reqs = [r for r in current_requests if r['Nombre'] == selected_person]
    if not my_reqs: st.caption("Ninguno")
    else:
        for i, r in enumerate(my_reqs):
            c1, c2 = st.columns([4, 1])
            c1.write(f"{r['Inicio'].strftime('%d/%m')} - {r['Fin'].strftime('%d/%m')}")
            if c2.button("🗑️", key=f"del_{selected_person}_{i}"):
                current_requests.remove(r)
                st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                st.rerun()

with c_vis:
    st.subheader("3. Visor Global")
    base_sch, _ = generate_base_schedule(year_val)
    st.markdown(render_annual_calendar(year_val, 'A', base_sch, st.session_state.nights), unsafe_allow_html=True)
    st.markdown(render_annual_calendar(year_val, 'B', base_sch, st.session_state.nights), unsafe_allow_html=True)
    st.markdown(render_annual_calendar(year_val, 'C', base_sch, st.session_state.nights), unsafe_allow_html=True)

# 4. FINAL
st.divider()
if st.button("🚀 Generar Excel Final", type="primary", use_container_width=True):
    excel_io = create_final_excel(edited_df, current_requests, year_val, st.session_state.nights)
    st.download_button("📥 Descargar Cuadrante", excel_io, f"Cuadrante_Final_{year_val}.xlsx")
