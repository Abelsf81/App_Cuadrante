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
ROLES = ["Mando", "Conductor", "Bombero"]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Plantilla por defecto
DEFAULT_ROSTER = [
    {"ID_Puesto": "Jefe A",      "Nombre": "Jefe A",      "Turno": "A", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Subjefe A",   "Nombre": "Subjefe A",   "Turno": "A", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Cond A",      "Nombre": "Cond A",      "Turno": "A", "Rol": "Conductor", "SV": True},
    {"ID_Puesto": "Bombero A1",  "Nombre": "Bombero A1",  "Turno": "A", "Rol": "Bombero",   "SV": True},
    {"ID_Puesto": "Bombero A2",  "Nombre": "Bombero A2",  "Turno": "A", "Rol": "Bombero",   "SV": False},
    {"ID_Puesto": "Bombero A3",  "Nombre": "Bombero A3",  "Turno": "A", "Rol": "Bombero",   "SV": False},
    
    {"ID_Puesto": "Jefe B",      "Nombre": "Jefe B",      "Turno": "B", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Subjefe B",   "Nombre": "Subjefe B",   "Turno": "B", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Cond B",      "Nombre": "Cond B",      "Turno": "B", "Rol": "Conductor", "SV": True},
    {"ID_Puesto": "Bombero B1",  "Nombre": "Bombero B1",  "Turno": "B", "Rol": "Bombero",   "SV": True},
    {"ID_Puesto": "Bombero B2",  "Nombre": "Bombero B2",  "Turno": "B", "Rol": "Bombero",   "SV": False},
    {"ID_Puesto": "Bombero B3",  "Nombre": "Bombero B3",  "Turno": "B", "Rol": "Bombero",   "SV": False},

    {"ID_Puesto": "Jefe C",      "Nombre": "Jefe C",      "Turno": "C", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Subjefe C",   "Nombre": "Subjefe C",   "Turno": "C", "Rol": "Mando",     "SV": False},
    {"ID_Puesto": "Cond C",      "Nombre": "Cond C",      "Turno": "C", "Rol": "Conductor", "SV": True},
    {"ID_Puesto": "Bombero C1",  "Nombre": "Bombero C1",  "Turno": "C", "Rol": "Bombero",   "SV": True},
    {"ID_Puesto": "Bombero C2",  "Nombre": "Bombero C2",  "Turno": "C", "Rol": "Bombero",   "SV": False},
    {"ID_Puesto": "Bombero C3",  "Nombre": "Bombero C3",  "Turno": "C", "Rol": "Bombero",   "SV": False},
]

# -------------------------------------------------------------------
# 1. MOTOR LÓGICO
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

def get_candidates(person_missing, roster_df, day_idx, current_schedule):
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    
    for _, candidate in roster_df.iterrows():
        if candidate['Turno'] == missing_turn: continue
        cand_status = current_schedule[candidate['Nombre']][day_idx]
        if cand_status != 'L': continue 
        
        is_compatible = False
        if missing_role == "Mando":
            if candidate['Rol'] == "Mando": is_compatible = True
        elif missing_role == "Conductor":
            if candidate['Rol'] == "Conductor": is_compatible = True
            if candidate['SV']: is_compatible = True
        elif missing_role == "Bombero":
            if candidate['Rol'] == "Bombero": is_compatible = True
            if candidate['SV']: is_compatible = True
            
        if is_compatible:
            candidates.append(candidate['Nombre'])
    return candidates

def is_night_restricted(date_obj, night_periods):
    for start, end in night_periods:
        if date_obj == start or date_obj == end: return True
    return False

def is_in_night_period(day_idx, year, night_periods):
    current_date = datetime.date(year, 1, 1) + datetime.timedelta(days=day_idx)
    for start, end in night_periods:
        if start <= current_date <= end: return True
    return False

def calculate_spent_credits(roster_df, requests, year):
    base_sch, _ = generate_base_schedule(year)
    credits = {name: 0 for name in roster_df['Nombre']}
    for req in requests:
        name = req['Nombre']
        if name not in credits: continue 
        row = roster_df[roster_df['Nombre'] == name]
        if row.empty: continue
        turn = row.iloc[0]['Turno']
        start_idx = req['Inicio'].timetuple().tm_yday - 1
        end_idx = req['Fin'].timetuple().tm_yday - 1
        cost = 0
        for d in range(start_idx, end_idx + 1):
            if base_sch[turn][d] == 'T': cost += 1
        credits[name] += cost
    return credits

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

def validate_and_generate(roster_df, requests, year, night_periods):
    base_schedule_turn, total_days = generate_base_schedule(year)
    final_schedule = {} 
    turn_coverage_counters = {'A': 0, 'B': 0, 'C': 0}
    person_coverage_counters = {name: 0 for name in roster_df['Nombre']}
    name_to_turn = {row['Nombre']: row['Turno'] for _, row in roster_df.iterrows()}
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    daily_error_codes = {} 

    for req in requests:
        name = req['Nombre']
        start_idx = req['Inicio'].timetuple().tm_yday - 1
        end_idx = req['Fin'].timetuple().tm_yday - 1
        for d in range(start_idx, end_idx + 1):
            if final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V'
            else:
                final_schedule[name][d] = 'V(L)'

    errors = []
    adjustments_log = []
    
    for d in range(total_days):
        absent_people = day_vacations[d]
        if not absent_people: continue
        
        if len(absent_people) > 2:
            date_str = (datetime.date(year, 1, 1) + datetime.timedelta(days=d)).strftime("%d-%m")
            errors.append(f"{date_str}: Hay {len(absent_people)} personas de vacaciones (Máx 2).")
            daily_error_codes[d] = "RED"
            continue
            
        if len(absent_people) == 2:
            p1 = roster_df[roster_df['Nombre'] == absent_people[0]].iloc[0]
            p2 = roster_df[roster_df['Nombre'] == absent_people[1]].iloc[0]
            if p1['Turno'] == p2['Turno']:
                errors.append(f"Día {d+1}: {p1['Nombre']} y {p2['Nombre']} son del mismo turno.")
                daily_error_codes[d] = "YELLOW"

        for name_missing in absent_people:
            person_row = roster_df[roster_df['Nombre'] == name_missing].iloc[0]
            candidates = get_candidates(person_row, roster_df, d, final_schedule)
            if not candidates:
                errors.append(f"Día {d+1}: Sin cobertura para {name_missing}.")
                if d not in daily_error_codes: daily_error_codes[d] = "ORANGE"
                continue
                
            valid_candidates = []
            for cand in candidates:
                prev_day = final_schedule[cand][d-1] if d > 0 else 'L'
                prev_prev = final_schedule[cand][d-2] if d > 1 else 'L'
                if prev_day.startswith('T') and prev_prev.startswith('T'): continue 
                valid_candidates.append(cand)
            
            if not valid_candidates:
                date_str = (datetime.date(year, 1, 1) + datetime.timedelta(days=d)).strftime("%d-%m")
                errors.append(f"{date_str}: {name_missing} no tiene cobertura válida (Regla Máx 2T).")
                if d not in daily_error_codes: daily_error_codes[d] = "ORANGE"
                continue
            
            if d not in daily_error_codes:
                def sort_key(cand_name):
                    cand_turn = name_to_turn[cand_name]
                    return (turn_coverage_counters[cand_turn], person_coverage_counters[cand_name], random.random())
                valid_candidates.sort(key=sort_key)
                chosen = valid_candidates[0]
                chosen_turn = name_to_turn[chosen]
                final_schedule[chosen][d] = f"T*({name_missing})"
                adjustments_log.append((d, chosen, name_missing))
                turn_coverage_counters[chosen_turn] += 1
                person_coverage_counters[chosen] += 1

    fill_log = {} 
    if not errors:
        for name in roster_df['Nombre']:
            current_v_days = [i for i, x in enumerate(final_schedule[name]) if x.startswith('V')]
            needed = 39 - len(current_v_days)
            added_dates = []
            if needed > 0:
                available_idx = [i for i, x in enumerate(final_schedule[name]) if x == 'L']
                if len(available_idx) >= needed:
                    fill_idxs = get_clustered_dates(available_idx, needed)
                    for idx in fill_idxs:
                        final_schedule[name][idx] = 'V(R)'
                        d_obj = datetime.date(year, 1, 1) + datetime.timedelta(days=idx)
                        added_dates.append(d_obj)
            fill_log[name] = added_dates

    return final_schedule, errors, person_coverage_counters, fill_log, adjustments_log, daily_error_codes

# -------------------------------------------------------------------
# 4. MOTOR AUTO-SOLVER Y REPARADOR (V6.2)
# -------------------------------------------------------------------

def check_request_conflict(req, occupation_map, base_schedule_turn, roster_df, night_periods, total_days):
    """Verifica si una solicitud choca con el mapa de ocupación actual."""
    start_idx = req['Inicio'].timetuple().tm_yday - 1
    end_idx = req['Fin'].timetuple().tm_yday - 1
    person = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0]
    
    # Check Nocturna
    if is_night_restricted(req['Inicio'], night_periods) or is_night_restricted(req['Fin'], night_periods):
        return "Nocturna"

    for d in range(start_idx, end_idx + 1):
        if d >= total_days: return "Fuera de rango"
        
        # Si es dia de trabajo, genera ocupación
        if base_schedule_turn[person['Turno']][d] == 'T':
            occupants = occupation_map[d]
            # Regla 1: Max 2
            if len(occupants) >= 2: return "Max 2 Personas"
            # Regla 2: Mismo Turno
            for occ in occupants:
                if occ['Turno'] == person['Turno']: return f"Conflicto Turno con {occ['Nombre']}"
            # Regla 3: Misma Categoria
            for occ in occupants:
                if occ['Rol'] == person['Rol'] and person['Rol'] != "Bombero": return f"Conflicto Rol con {occ['Nombre']}"
    return None

def book_request(req, occupation_map, base_schedule_turn, roster_df):
    start_idx = req['Inicio'].timetuple().tm_yday - 1
    end_idx = req['Fin'].timetuple().tm_yday - 1
    person = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0].to_dict()
    credits = 0
    for d in range(start_idx, end_idx + 1):
        if base_schedule_turn[person['Turno']][d] == 'T':
            occupation_map[d].append(person)
            credits += 1
    return credits

def smart_repair_requests(roster_df, imported_requests, year, night_periods):
    """
    Intenta arreglar conflictos moviendo las fechas +/- X días.
    """
    base_schedule_turn, total_days = generate_base_schedule(year)
    occupation_map = {i: [] for i in range(total_days)}
    
    accepted_requests = []
    change_log = [] # Strings describiendo cambios
    
    # Ordenar por prioridad (Jefes primero, o simplemente orden de lista)
    # Asumimos el orden del Excel es la prioridad
    
    for req in imported_requests:
        # Intentar bookear original
        conflict = check_request_conflict(req, occupation_map, base_schedule_turn, roster_df, night_periods, total_days)
        
        if not conflict:
            book_request(req, occupation_map, base_schedule_turn, roster_df)
            accepted_requests.append(req)
        else:
            # CONFLICTO DETECTADO -> INICIAR REPARACIÓN
            original_start = req['Inicio']
            duration = (req['Fin'] - req['Inicio']).days + 1
            found_fix = False
            
            # Probar desplazamientos: +1, -1, +2, -2 ... hasta +/- 15 días
            shifts = []
            for i in range(1, 16):
                shifts.append(i)
                shifts.append(-i)
            
            for delta in shifts:
                new_start = original_start + datetime.timedelta(days=delta)
                new_end = new_start + datetime.timedelta(days=duration-1)
                
                # Verificar si cae en año correcto
                if new_start.year != year or new_end.year != year: continue
                
                new_req = {"Nombre": req['Nombre'], "Inicio": new_start, "Fin": new_end}
                
                # Verificar si el nuevo hueco es válido
                # IMPORTANTE: Verificar que el nuevo periodo tenga sentido (empiece en T si el original era T)
                # Para simplificar, solo validamos reglas de oro. El usuario decidirá si le gusta.
                
                new_conflict = check_request_conflict(new_req, occupation_map, base_schedule_turn, roster_df, night_periods, total_days)
                
                if not new_conflict:
                    book_request(new_req, occupation_map, base_schedule_turn, roster_df)
                    accepted_requests.append(new_req)
                    change_log.append(f"🔄 {req['Nombre']}: {req['Inicio'].strftime('%d/%m')} movido a {new_start.strftime('%d/%m')} ({conflict})")
                    found_fix = True
                    break
            
            if not found_fix:
                change_log.append(f"❌ {req['Nombre']}: {req['Inicio'].strftime('%d/%m')} RECHAZADO (Imposible encajar).")

    return accepted_requests, change_log

def run_auto_solver_fill(roster_df, year, night_periods, existing_requests):
    # Wrapper para rellenar huecos tras la carga manual/reparada
    # Reutilizamos logica V6.1 pero asegurando que existing_requests ya están en occupation_map
    
    base_schedule_turn, total_days = generate_base_schedule(year)
    occupation_map = {i: [] for i in range(total_days)}
    
    # Pre-llenar con lo existente
    for req in existing_requests:
        book_request(req, occupation_map, base_schedule_turn, roster_df)
        
    # ... (Resto lógica de relleno automático V6.1) ...
    # Por brevedad, aquí repetimos la lógica de relleno simple
    
    final_requests = list(existing_requests)
    people = roster_df.to_dict('records')
    people.sort(key=lambda x: x['Turno'])
    
    all_days = [datetime.date(year, 1, 1) + datetime.timedelta(days=i) for i in range(total_days)]
    
    for p in people:
        # Calcular créditos actuales
        credits_got = 0
        # ... (misma logica de conteo) ...
        # Recalcular porque es complejo pasar estado
        for r in final_requests:
            if r['Nombre'] == p['Nombre']:
                s_idx = r['Inicio'].timetuple().tm_yday - 1
                e_idx = r['Fin'].timetuple().tm_yday - 1
                for d in range(s_idx, e_idx + 1):
                    if base_schedule_turn[p['Turno']][d] == 'T': credits_got += 1
                    
        if credits_got >= 13: continue
        
        # Intentar rellenar
        attempts = 0
        while credits_got < 13 and attempts < 200:
            day = random.choice(all_days)
            d_idx = day.timetuple().tm_yday - 1
            
            # Solo días T
            if base_schedule_turn[p['Turno']][d_idx] == 'T':
                req = {"Nombre": p['Nombre'], "Inicio": day, "Fin": day}
                if not check_request_conflict(req, occupation_map, base_schedule_turn, roster_df, night_periods, total_days):
                    # Check self overlap
                    overlap = False
                    for r in final_requests:
                        if r['Nombre'] == p['Nombre'] and r['Inicio'] == day: overlap = True
                    
                    if not overlap:
                        book_request(req, occupation_map, base_schedule_turn, roster_df)
                        final_requests.append(req)
                        credits_got += 1
            attempts += 1
            
    return final_requests


# -------------------------------------------------------------------
# 3. EXPORTADORES
# -------------------------------------------------------------------
def generate_visual_error_report(schedule, roster_df, year, night_periods, error_heatmap, text_errors):
    wb = Workbook()
    fill_red = PatternFill("solid", fgColor="FF0000")
    fill_yellow = PatternFill("solid", fgColor="FFD700")
    fill_orange = PatternFill("solid", fgColor="FFA500")
    fill_purple = PatternFill("solid", fgColor="CBC3E3")
    font_bold = Font(bold=True)
    font_white = Font(color="FFFFFF", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_all = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws1 = wb.active; ws1.title = "Mapa de Conflictos"
    ws1.column_dimensions['A'].width = 15
    for i in range(2, 34): ws1.column_dimensions[get_column_letter(i)].width = 4
    current_row = 1
    for t in TEAMS:
        ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=32)
        cell_title = ws1.cell(current_row, 1, f"TURNO {t}")
        cell_title.font = Font(bold=True, size=14, color="FFFFFF")
        cell_title.fill = PatternFill("solid", fgColor="B22222")
        cell_title.alignment = align_c
        current_row += 2
        team_members = roster_df[roster_df['Turno'] == t]
        for _, p in team_members.iterrows():
            name = p['Nombre']
            ws1.cell(current_row, 1, name).font = font_bold
            for d in range(1, 32):
                c = ws1.cell(current_row, d+1, d); c.alignment=align_c; c.border=border_all
            current_row += 1
            for m_idx, mes in enumerate(MESES):
                month_num = m_idx + 1
                ws1.cell(current_row, 1, mes).font = font_bold
                days_in_month = calendar.monthrange(year, month_num)[1]
                for d in range(1, 32):
                    cell = ws1.cell(current_row, d+1); cell.border=border_all; cell.alignment=align_c
                    if d <= days_in_month:
                        date_obj = datetime.date(year, month_num, d)
                        d_idx = date_obj.timetuple().tm_yday - 1
                        status = schedule[name][d_idx]
                        val = ""; fill = PatternFill("solid", fgColor="F2F2F2")
                        if status == 'T': val = "T"; fill = PatternFill("solid", fgColor="C6EFCE")
                        elif 'V' in status: val = "V"; fill = PatternFill("solid", fgColor="FFEB9C")
                        if is_in_night_period(d_idx, year, night_periods): fill = PatternFill("solid", fgColor="A6A6A6")
                        if (name, d_idx) in error_heatmap:
                            fill = fill_red
                            val = error_heatmap[(name, d_idx)]
                            if val == "ERR: Turno": fill = fill_yellow
                            if val == "ERR: Max 2T": fill = fill_orange
                            cell.font = font_white
                        cell.value = val; cell.fill = fill
                    else: cell.fill = PatternFill("solid", fgColor="808080")
                current_row += 1
            current_row += 2

    ws2 = wb.create_sheet("Lista de Errores")
    ws2.column_dimensions['A'].width = 80
    ws2.append(["Descripción del Conflicto"])
