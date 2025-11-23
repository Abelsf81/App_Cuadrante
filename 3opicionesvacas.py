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

# --- ESTRATEGIAS DE VACACIONES (EL MENÚ COMPLETO - 5 OPCIONES) ---
STRATEGIES = {
    "standard": {
        "name": "🛡️ Estándar (4 Bloques)",
        "desc": "10+10+10+9 días. Requiere iniciar uno en T para cuadrar.",
        "blocks": [
            {"dur": 10, "cred": 4, "label": "Bloque 10d (4 Cr)"},
            {"dur": 10, "cred": 3, "label": "Bloque 10d (3 Cr)"},
            {"dur": 9,  "cred": 3, "label": "Bloque 9d (3 Cr)"}
        ],
        "auto_recipe": [ 
            {"dur": 10, "target": 4}, 
            {"dur": 10, "target": 3}, 
            {"dur": 10, "target": 3}, 
            {"dur": 9, "target": 3}
        ]
    },
    "safe": {
        "name": "🔢 Matemática Pura (4 Bloques)",
        "desc": "12+12+9+6 días. Indestructible: Múltiplos de 3 siempre cuadran.",
        "blocks": [
            {"dur": 12, "cred": 4, "label": "Largo 12d (4 Cr)"},
            {"dur": 9,  "cred": 3, "label": "Medio 9d (3 Cr)"},
            {"dur": 6,  "cred": 2, "label": "Corto 6d (2 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 12, "target": 4}, 
            {"dur": 12, "target": 4}, 
            {"dur": 9, "target": 3},
            {"dur": 6, "target": 2}
        ]
    },
    "balanced": {
        "name": "⚖️ Tridente (3 Bloques)",
        "desc": "13+13+13 días. Reparto equitativo del año.",
        "blocks": [
            {"dur": 13, "cred": 5, "label": "Bloque Mayor 13d (5 Cr)"},
            {"dur": 13, "cred": 4, "label": "Bloque Menor 13d (4 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 13, "target": 5}, 
            {"dur": 13, "target": 4}, 
            {"dur": 13, "target": 4}
        ]
    },
    "long": {
        "name": "✈️ Larga Estancia (3 Bloques)",
        "desc": "15+15+9 días. Ideal para viajes largos.",
        "blocks": [
            {"dur": 15, "cred": 5, "label": "Gran Viaje 15d (5 Cr)"},
            {"dur": 9,  "cred": 3, "label": "Escapada 9d (3 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 15, "target": 5}, 
            {"dur": 15, "target": 5}, 
            {"dur": 9, "target": 3}
        ]
    },
    "micro": {
        "name": "🐜 Hormiga (6 Bloques)",
        "desc": "5x6 días + 1x9 días. Muchos cortes pequeños.",
        "blocks": [
            {"dur": 6, "cred": 2, "label": "Semana 6d (2 Cr)"},
            {"dur": 9, "cred": 3, "label": "Semana+ 9d (3 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 6, "target": 2}, {"dur": 6, "target": 2}, 
            {"dur": 6, "target": 2}, {"dur": 6, "target": 2}, 
            {"dur": 6, "target": 2}, {"dur": 9, "target": 3}
        ]
    }
}

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
        dates.add(end) # Solo el final es crítico
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
# 2. MOTOR INTELIGENTE (VALIDADOR & GENERADOR)
# -------------------------------------------------------------------

def check_global_conflict_generic(start_idx, duration, person, occupation_map, base_sch, year, transition_dates):
    """Verifica conflictos contra un mapa de ocupación dado."""
    total_days = len(base_sch['A'])
    if start_idx + duration > total_days: return True

    for i in range(start_idx, start_idx + duration):
        # 1. Nocturna (Fin de periodo y trabajo)
        d_obj = datetime.date(year, 1, 1) + timedelta(days=i)
        if d_obj in transition_dates:
            if base_sch[person['Turno']][i] == 'T': return True
        
        # 2. Ocupación
        occupants = occupation_map.get(i, [])
        if len(occupants) >= 2: return True
        
        for occ in occupants:
            # 3. Mismo Turno
            if occ['Turno'] == person['Turno']: return True
            # 4. Misma Categoría
            if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: return True
            
    return False

def get_available_blocks_for_person(person_name, roster_df, current_requests, year, night_periods, month_range, strategy_key):
    """Genera el MENÚ DE OPCIONES DINÁMICO según estrategia."""
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    person = roster_df[roster_df['Nombre'] == person_name].iloc[0]
    
    start_month_idx = MESES.index(month_range[0]) + 1
    end_month_idx = MESES.index(month_range[1]) + 1
    
    occupation_map = {i:[] for i in range(total_days)}
    my_current_slots = [] 

    for req in current_requests:
        p_req = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0]
        s = req['Inicio'].timetuple().tm_yday - 1
        e = req['Fin'].timetuple().tm_yday - 1
        
        if req['Nombre'] != person_name:
            for d in range(s, e+1):
                if base_sch[p_req['Turno']][d] == 'T': occupation_map[d].append(p_req)
        else:
            my_current_slots.append((s, e))

    block_defs = STRATEGIES[strategy_key]['blocks']
    options = {b['label']: [] for b in block_defs}
    
    for d in range(total_days - 15): 
        d_date = datetime.date(year, 1, 1) + timedelta(days=d)
        if not (start_month_idx <= d_date.month <= end_month_idx): continue

        for b_def in block_defs:
            duration = b_def['dur']
            target_cred = b_def['cred']
            label_key = b_def['label']
            
            if not check_global_conflict_generic(d, duration, person, occupation_map, base_sch, year, transition_dates):
                overlap = False
                for ms in my_current_slots:
                    if not (d + duration - 1 < ms[0] - 2 or d > ms[1] + 2): overlap = True; break
                
                if not overlap:
                    credits = 0
                    for k in range(d, d+duration):
                        if base_sch[person['Turno']][k] == 'T': credits += 1
                    
                    if credits == target_cred:
                        start_date = d_date
                        end_date = start_date + timedelta(days=duration-1)
                        txt = f"{start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
                        options[label_key].append({'label': txt, 'start': start_date, 'end': end_date})

    return options

def auto_generate_schedule(roster_df, year, night_periods, strategy_key):
    """El Arquitecto Multi-Estrategia."""
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    occupation_map = {} 
    generated_requests = []
    
    people = roster_df.to_dict('records')
    priority_order = ["Jefe", "Subjefe", "Conductor", "Bombero"]
    people.sort(key=lambda x: priority_order.index(x['Rol']))
    
    RECIPE = STRATEGIES[strategy_key]['auto_recipe']
    
    for person in people:
        my_slots = []
        current_recipe = RECIPE.copy()
        random.shuffle(current_recipe) 
        
        for block in current_recipe:
            duration = block['dur']
            target = block['target']
            
            options = []
            for d in range(0, total_days - duration):
                c = sum([1 for k in range(d, d+duration) if base_sch[person['Turno']][k] == 'T'])
                if c == target:
                     if not check_global_conflict_generic(d, duration, person, occupation_map, base_sch, year, transition_dates):
                        options.append(d)
            
            random.shuffle(options)
            
            for start in options:
                overlap = any(start < s[0]+s[1]+2 and start+duration > s[0]-2 for s in my_slots)
                if not overlap:
                    for k in range(start, start+duration):
                        if k not in occupation_map: occupation_map[k] = []
                        occupation_map[k].append(person)
                    my_slots.append((start, duration))
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=start),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=start+duration-1)
                    })
                    break 

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
# 4. GENERACIÓN FINAL (EXCEL) Y AUXILIARES
# -------------------------------------------------------------------
def get_clustered_dates(available_idxs, needed_count):
    if not available_idxs: return []
    groups = []
    for k, g in groupby(enumerate(available_idxs), lambda ix: ix[0] - ix[1]):
        groups.append(list(map(itemgetter(1), g)))
    groups.sort(key=len, reverse=True)
    selected = []
    for group in groups:
        if len(selected) < needed_count:
            take = min(len(group), needed_count - len(selected))
            selected.extend(group[:take])
        else: break
    return sorted(selected)

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

    # Relleno Administrativo (V(R))
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
# INTERFAZ STREAMLIT (V18.0 - MULTI-ESTRATEGIA)
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V18.0")

st.title("🚒 Gestor V18.0: El Estratega")
st.caption("Modo Copiloto: Elige estrategia y selecciona las mejores fechas.")

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
                            d1 = pd.to_datetime(v1).date()
                            d2 = pd.to_datetime(v2).date()
                            st.session_state.nights.append((d1, d2)); c+=1
                    except: pass
                if c>0: st.success(f"Cargadas {c}")
            except: pass
        if st.button("Limpiar Nocturnas"): st.session_state.nights = []

    # --- SELECTOR DE ESTRATEGIA ---
    st.divider()
    strategy_key = st.selectbox(
        "🎯 Estrategia de Vacaciones", 
        options=list(STRATEGIES.keys()), 
        format_func=lambda x: STRATEGIES[x]['name']
    )
    st.info(STRATEGIES[strategy_key]['desc'])

    # BOTÓN AUTO
    if st.button("🎲 Generar Automático", type="primary"):
        with st.spinner("Generando..."):
            new_reqs = auto_generate_schedule(edited_df, year_val, st.session_state.nights, strategy_key)
            st.session_state.raw_requests_df = pd.DataFrame(new_reqs)
        st.success("¡Hecho!")
        st.rerun()

# 2. ESTADO
if 'raw_requests_df' not in st.session_state:
    st.session_state.raw_requests_df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
current_requests = st.session_state.raw_requests_df.to_dict('records')
stats = calculate_stats(edited_df, current_requests, year_val)

# 3. DRAFT ROOM (COPILOTO)
st.divider()
c_main, c_vis = st.columns([1, 2])

with c_main:
    st.subheader("2. Selección Manual (Copiloto)")
    
    all_names = edited_df['Nombre'].tolist()
    names_sorted = sorted(all_names, key=lambda x: (0 if "Jefe" in x else 1 if "Subjefe" in x else 2 if "Cond" in x else 3))
    
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
            month_range = st.select_slider("📅 Filtrar Meses:", options=MESES, value=(MESES[0], MESES[-1]))
            st.info(f"🔍 Buscando fichas de tipo: {STRATEGIES[strategy_key]['name']}")
            
            # OBTENER OPCIONES DINÁMICAS
            options = get_available_blocks_for_person(
                selected_person, edited_df, current_requests, year_val, st.session_state.nights, month_range, strategy_key
            )
            
            # PESTAÑAS DINÁMICAS
            block_defs = STRATEGIES[strategy_key]['blocks']
            tabs = st.tabs([b['label'] for b in block_defs])
            
            for i, b_def in enumerate(block_defs):
                key = b_def['label']
                with tabs[i]:
                    available_opts = options.get(key, [])
                    if not available_opts:
                        st.warning("Sin opciones disponibles.")
                    else:
                        with st.container(height=200):
                            # CLAVE ÚNICA (FIXED)
                            for opt in available_opts[:20]: 
                                if st.button(f"➕ {opt['label']}", key=f"add_{selected_person}_{opt['start']}_{i}"):
                                    current_requests.append({"Nombre": selected_person, "Inicio": opt['start'], "Fin": opt['end']})
                                    st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                                    st.rerun()

    st.markdown("---")
    st.write(f"**Mis Periodos:**")
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
    sch, adj, count, fill = validate_and_generate_final(edited_df, current_requests, year_val, st.session_state.nights)
    excel_io = create_final_excel(sch, edited_df, year_val, current_requests, fill, count, st.session_state.nights, adj)
    st.download_button("📥 Descargar Cuadrante", excel_io, f"Cuadrante_Final_{year_val}.xlsx")
