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
    # Secuencia A(T) -> B(T) -> C(T)
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
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    
    # 1. Aplicar Vacaciones
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
    
    # 2. Resolver Coberturas
    for d in range(total_days):
        absent_people = day_vacations[d]
        if not absent_people: continue
        
        if len(absent_people) > 2:
            date_str = (datetime.date(year, 1, 1) + datetime.timedelta(days=d)).strftime("%d-%m")
            errors.append(f"{date_str}: Hay {len(absent_people)} personas de vacaciones (Máx 2).")
            continue
            
        if len(absent_people) == 2:
            p1 = roster_df[roster_df['Nombre'] == absent_people[0]].iloc[0]
            p2 = roster_df[roster_df['Nombre'] == absent_people[1]].iloc[0]
            if p1['Turno'] == p2['Turno']:
                errors.append(f"Día {d+1}: {p1['Nombre']} y {p2['Nombre']} son del mismo turno.")

        for name_missing in absent_people:
            person_row = roster_df[roster_df['Nombre'] == name_missing].iloc[0]
            candidates = get_candidates(person_row, roster_df, d, final_schedule)
            
            if not candidates:
                errors.append(f"Día {d+1}: Sin cobertura para {name_missing}.")
                continue
                
            valid_candidates = []
            for cand in candidates:
                prev_day = final_schedule[cand][d-1] if d > 0 else 'L'
                prev_prev = final_schedule[cand][d-2] if d > 1 else 'L'
                
                is_prev_work = prev_day.startswith('T')
                is_prev_prev_work = prev_prev.startswith('T')
                
                if is_prev_work and is_prev_prev_work:
                    continue 
                
                valid_candidates.append(cand)
            
            if not valid_candidates:
                date_str = (datetime.date(year, 1, 1) + datetime.timedelta(days=d)).strftime("%d-%m")
                errors.append(f"{date_str}: {name_missing} no tiene cobertura válida (Todos violarían Máx 2T).")
                continue
            
            def sort_key(cand_name):
                cand_turn = roster_df[roster_df['Nombre'] == cand_name].iloc[0]['Turno']
                return turn_coverage_counters[cand_turn]
            
            valid_candidates.sort(key=sort_key)
            chosen = valid_candidates[0]
            chosen_turn = roster_df[roster_df['Nombre'] == chosen].iloc[0]['Turno']
            
            final_schedule[chosen][d] = f"T*({person_row['Turno']})"
            turn_coverage_counters[chosen_turn] += 1
            person_coverage_counters[chosen] += 1

    # 3. Relleno Administrativo
    fill_log = {} 
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

    return final_schedule, errors, person_coverage_counters, fill_log

# -------------------------------------------------------------------
# 2. GENERACIÓN EXCEL
# -------------------------------------------------------------------
def create_excel(schedule, roster_df, year, requests, fill_log, counters, night_periods):
    wb = Workbook()
    
    s_T = PatternFill("solid", fgColor="C6EFCE") 
    s_V = PatternFill("solid", fgColor="FFEB9C") 
    s_VR = PatternFill("solid", fgColor="FFFFE0") 
    s_Cov = PatternFill("solid", fgColor="FFC7CE") 
    s_L = PatternFill("solid", fgColor="F2F2F2") 
    s_Night = PatternFill("solid", fgColor="A6A6A6") 
    
    font_bold = Font(bold=True)
    font_red = Font(color="9C0006", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_thin = Side(border_style="thin", color="000000")
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    # HOJA 1: CUADRANTE
    ws1 = wb.active
    ws1.title = "Cuadrante"
    ws1.column_dimensions['A'].width = 15
    for i in range(2, 34): ws1.column_dimensions[get_column_letter(i)].width = 4
    current_row = 1
    for t in TEAMS:
        ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=32)
        cell_title = ws1.cell(current_row, 1, f"TURNO {t}")
        cell_title.font = Font(bold=True, size=14, color="FFFFFF")
        cell_title.fill = PatternFill("solid", fgColor="000080")
        cell_title.alignment = align_c
        current_row += 2
        team_members = roster_df[roster_df['Turno'] == t]
        for _, p in team_members.iterrows():
            name = p['Nombre']
            role = p['Rol']
            ws1.cell(current_row, 1, f"{name} ({role})").font = font_bold
            for d in range(1, 32):
                c = ws1.cell(current_row, d+1, d)
                c.alignment = align_c
                c.font = font_bold
                c.border = border_all
                c.fill = PatternFill("solid", fgColor="E0E0E0")
            current_row += 1
            for m_idx, mes in enumerate(MESES):
                month_num = m_idx + 1
                ws1.cell(current_row, 1, mes).font = font_bold
                ws1.cell(current_row, 1).border = border_all
                days_in_month = calendar.monthrange(year, month_num)[1]
                for d in range(1, 32):
                    cell = ws1.cell(current_row, d+1)
                    cell.border = border_all
                    cell.alignment = align_c
                    if d <= days_in_month:
                        date_obj = datetime.date(year, month_num, d)
                        day_of_year = date_obj.timetuple().tm_yday - 1
                        status = schedule[name][day_of_year]
                        val = ""
                        fill = s_L 
                        if status == 'T':
                            val = "T"
                            fill = s_T
                        elif status == 'V':
                            val = "V"
                            fill = s_V
                        elif status == 'V(L)' or status == 'V(R)':
                            val = "v"
                            fill = s_VR
                        elif status.startswith('T*'):
                            val = status.split('(')[1][0] 
                            fill = s_Cov
                            cell.font = font_red
                        if is_in_night_period(day_of_year, year, night_periods):
                            fill = s_Night
                        cell.value = val
                        cell.fill = fill
                    else:
                        cell.fill = PatternFill("solid", fgColor="808080")
                current_row += 1
            current_row += 2 

    # HOJA 2: ESTADÍSTICAS
    ws2 = wb.create_sheet("Estadísticas")
    ws2.column_dimensions['A'].width = 20
    headers = ["Nombre", "Turno", "Puesto", "Gastado (T)", "Coberturas (T*)", "Total Días (T+T*)", "Total Vacs (Nat)"]
    ws2.append(headers)
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        sch = schedule[name]
        base_sch_turn, _ = generate_base_schedule(year)
        original_ts = base_sch_turn[p['Turno']].count('T')
        v_credits = sch.count('V')
        t_cover = counters[name]
        total_work = (original_ts - v_credits) + t_cover
        v_natural = sch.count('V') + sch.count('V(L)') + sch.count('V(R)')
        ws2.append([name, p['Turno'], p['Rol'], v_credits, t_cover, total_work, v_natural])

    # HOJA 3: RESUMEN
    ws3 = wb.create_sheet("Resumen Solicitudes")
    ws3.append(["Nombre", "Turno", "Rol", "Periodos Solicitados", "Días Relleno (Automático)"])
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        person_reqs = [f"{r['Inicio'].strftime('%d/%m')} al {r['Fin'].strftime('%d/%m')}" for r in requests if r['Nombre'] == name]
        req_str = " | ".join(person_reqs) if person_reqs else "Sin solicitudes"
        
        fill_dates = fill_log[name]
        fill_str = "Ninguno"
        if fill_dates:
            date_ranges = []
            fill_dates.sort()
            if fill_dates:
                range_start = fill_dates[0]
                range_end = fill_dates[0]
                for i in range(1, len(fill_dates)):
                    if (fill_dates[i] - fill_dates[i-1]).days == 1:
                        range_end = fill_dates[i]
                    else:
                        if range_start == range_end:
                            date_ranges.append(range_start.strftime('%d/%m'))
                        else:
                            date_ranges.append(f"{range_start.strftime('%d/%m')}-{range_end.strftime('%d/%m')}")
                        range_start = fill_dates[i]
                        range_end = fill_dates[i]
                if range_start == range_end:
                    date_ranges.append(range_start.strftime('%d/%m'))
                else:
                    date_ranges.append(f"{range_start.strftime('%d/%m')}-{range_end.strftime('%d/%m')}")
            fill_str = ", ".join(date_ranges)

        ws3.append([name, p['Turno'], p['Rol'], req_str, fill_str])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ STREAMLIT
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V3.8")

st.title("🚒 Gestor Integral V3.8")

# 1. CONFIGURACIÓN
c1, c2 = st.columns([2, 1])
with c1:
    with st.expander("1. Configuración de Plantilla", expanded=False):
        if 'roster_data' not in st.session_state:
            st.session_state.roster_data = pd.DataFrame(DEFAULT_ROSTER)
        edited_df = st.data_editor(
            st.session_state.roster_data,
            column_config={
                "ID_Puesto": st.column_config.TextColumn(disabled=True),
                "Turno": st.column_config.SelectboxColumn(options=TEAMS, required=True),
                "Rol": st.column_config.SelectboxColumn(options=ROLES, required=True),
                "SV": st.column_config.CheckboxColumn(label="¿Es SV?", help="Puede cubrir conductor")
            },
            num_rows="dynamic",
            use_container_width=True
        )
        st.session_state.roster_data = edited_df

with c2:
    with st.expander("🌑 Periodos Nocturnos", expanded=False):
        if 'nights' not in st.session_state: st.session_state.nights = []
        dn_start = st.date_input("Inicio Noche", value=None)
        dn_end = st.date_input("Fin Noche", value=None)
        if st.button("Añadir Periodo"):
            if dn_start and dn_end: st.session_state.nights.append((dn_start, dn_end))
        if st.session_state.nights:
            for i, (s, e) in enumerate(st.session_state.nights):
                col_del, col_tx = st.columns([1,4])
                if col_del.button("x", key=f"n_{i}"):
                    st.session_state.nights.pop(i)
                    st.rerun()
                col_tx.text(f"{s.strftime('%d/%m')} - {e.strftime('%d/%m')}")

# 2. GESTOR
st.divider()
col_main, col_list = st.columns([2, 1])
names_list = edited_df['Nombre'].tolist()
today = datetime.date.today()
year_val = st.number_input("Año", value=today.year + 1)

if 'requests' not in st.session_state: st.session_state.requests = []
credits_map = calculate_spent_credits(edited_df, st.session_state.requests, year_val)

with col_main:
    with st.expander("📂 Carga Masiva Horizontal"):
        template_df = edited_df[['ID_Puesto', 'Nombre']].copy()
        # --- AUMENTADO A 20 PERIODOS ---
        for i in range(1, 21): 
            template_df[f'Inicio {i}'] = ""
            template_df[f'Fin {i}'] = ""
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            template_df.to_excel(writer, index=False)
        st.download_button("⬇️ Descargar Plantilla", buffer.getvalue(), "plantilla_h.xlsx")
        
        uploaded_file = st.file_uploader("Sube Excel", type=['xlsx'])
        if uploaded_file and st.button("Procesar Archivo"):
            try:
                df_upload = pd.read_excel(uploaded_file)
                count = 0
                for index, row in df_upload.iterrows():
                    target_name = None
                    if 'ID_Puesto' in row and not pd.isnull(row['ID_Puesto']):
                        match = edited_df[edited_df['ID_Puesto'] == row['ID_Puesto']]
                        if not match.empty: target_name = match.iloc[0]['Nombre']
                    if not target_name and 'Nombre' in row:
                        if row['Nombre'] in names_list: target_name = row['Nombre']
                    if target_name:
                        # --- BUCLE AUMENTADO A 20 ---
                        for i in range(1, 21):
                            col_start = f'Inicio {i}'
                            col_end = f'Fin {i}'
                            if col_start in row and col_end in row:
                                val_start = row[col_start]
                                val_end = row[col_end]
                                if not pd.isnull(val_start) and not pd.isnull(val_end):
                                    st.session_state.requests.append({
                                        "Nombre": target_name,
                                        "Inicio": pd.to_datetime(val_start).date(),
                                        "Fin": pd.to_datetime(val_end).date()
                                    })
                                    count += 1
                if count > 0: st.success(f"✅ Importados {count}.")
                st.rerun()
            except Exception as e: st.error(f"Error: {e}")

    st.subheader("2. Añadir Solicitud Manual")
    sel_name = st.selectbox("Trabajador", names_list)
    if sel_name:
        spent = credits_map.get(sel_name, 0)
        st.progress(min(spent/13, 1.0), text=f"Créditos T: {spent} / 13")
        
        row_p = edited_df[edited_df['Nombre'] == sel_name].iloc[0]
        base_sch, _ = generate_base_schedule(year_val)
        my_sch = base_sch[row_p['Turno']]
        view_months = list(range(1, 13))
        
        html_cal = "<div style='display:flex; flex-wrap:wrap; gap:5px; margin-bottom:10px;'>"
        for m in view_months:
            html_cal += f"<div style='border:1px solid #ddd; padding:2px; border-radius:3px; width:100px;'><strong>{MESES[m-1]}</strong>"
            days_in_m = calendar.monthrange(year_val, m)[1]
            html_cal += "<div style='display:grid; grid-template-columns:repeat(7, 1fr); gap:1px; font-size:9px; text-align:center;'>"
            for d in range(1, days_in_m + 1):
                dt = datetime.date(year_val, m, d)
                d_idx = dt.timetuple().tm_yday - 1
                status = my_sch[d_idx]
                color = "#C6EFCE" if status == 'T' else "#F2F2F2"
                border = "2px solid #555" if is_in_night_period(d_idx, year_val, st.session_state.nights) else "1px solid #eee"
                html_cal += f"<div style='background-color:{color}; padding:1px; border:{border}'>{d}</div>"
            html_cal += "</div></div>"
        html_cal += "</div>"
        st.markdown(html_cal, unsafe_allow_html=True)

    d_range = st.date_input("Selecciona Rango", [], help="Inicio - Fin")
    if st.button("Añadir Periodo", use_container_width=True):
        if len(d_range) == 2:
            start, end = d_range
            conflict = False
            if is_night_restricted(start, st.session_state.nights) or is_night_restricted(end, st.session_state.nights):
                st.error("⛔ Conflicto periodo nocturno.")
                conflict = True
            if not conflict:
                st.session_state.requests.append({"Nombre": sel_name, "Inicio": start, "Fin": end})
                st.success(f"Añadido: {sel_name}")
                st.rerun()
        else: st.warning("Selecciona fechas.")

with col_list:
    st.subheader("Listado")
    if st.session_state.requests:
        for i, r in enumerate(st.session_state.requests):
            c_txt, c_btn = st.columns([4, 1])
            c_txt.text(f"{r['Nombre']}\n{r['Inicio'].strftime('%d/%m')} - {r['Fin'].strftime('%d/%m')}")
            if c_btn.button("🗑️", key=f"del_{i}"):
                st.session_state.requests.pop(i)
                st.rerun()

st.divider()
if st.button("🚀 Generar Excel Final", type="primary", use_container_width=True):
    if not st.session_state.requests:
        st.error("Faltan solicitudes.")
    else:
        final_sch, errs, counters, fill_log = validate_and_generate(
            edited_df, st.session_state.requests, year_val, st.session_state.nights
        )
        if errs:
            st.error("❌ Conflictos:")
            for e in errs: st.write(f"- {e}")
        else:
            st.success("✅ Éxito")
            excel_data = create_excel(
                final_sch, edited_df, year_val, st.session_state.requests, fill_log, counters, st.session_state.nights
            )
            st.download_button("📥 Descargar", excel_data, f"Cuadrante_V3.8_{year_val}.xlsx")
