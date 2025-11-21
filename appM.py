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

# --- CONSTANTES Y CONFIGURACIÓN ---
TEAMS = ['A', 'B', 'C']
ROLES = ["Mando", "Conductor", "Bombero"]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Plantilla por defecto
DEFAULT_ROSTER = [
    {"Nombre": "Jefe A", "Turno": "A", "Rol": "Mando", "SV": False},
    {"Nombre": "Subjefe A", "Turno": "A", "Rol": "Mando", "SV": False},
    {"Nombre": "Cond A", "Turno": "A", "Rol": "Conductor", "SV": True},
    {"Nombre": "Bombero A1", "Turno": "A", "Rol": "Bombero", "SV": True},
    {"Nombre": "Bombero A2", "Turno": "A", "Rol": "Bombero", "SV": False},
    {"Nombre": "Bombero A3", "Turno": "A", "Rol": "Bombero", "SV": False},
    
    {"Nombre": "Jefe B", "Turno": "B", "Rol": "Mando", "SV": False},
    {"Nombre": "Subjefe B", "Turno": "B", "Rol": "Mando", "SV": False},
    {"Nombre": "Cond B", "Turno": "B", "Rol": "Conductor", "SV": True},
    {"Nombre": "Bombero B1", "Turno": "B", "Rol": "Bombero", "SV": True},
    {"Nombre": "Bombero B2", "Turno": "B", "Rol": "Bombero", "SV": False},
    {"Nombre": "Bombero B3", "Turno": "B", "Rol": "Bombero", "SV": False},

    {"Nombre": "Jefe C", "Turno": "C", "Rol": "Mando", "SV": False},
    {"Nombre": "Subjefe C", "Turno": "C", "Rol": "Mando", "SV": False},
    {"Nombre": "Cond C", "Turno": "C", "Rol": "Conductor", "SV": True},
    {"Nombre": "Bombero C1", "Turno": "C", "Rol": "Bombero", "SV": True},
    {"Nombre": "Bombero C2", "Turno": "C", "Rol": "Bombero", "SV": False},
    {"Nombre": "Bombero C3", "Turno": "C", "Rol": "Bombero", "SV": False},
]

# -------------------------------------------------------------------
# 1. MOTOR LÓGICO
# -------------------------------------------------------------------

def generate_base_schedule(year):
    is_leap = calendar.isleap(year)
    total_days = 366 if is_leap else 365
    status = {'A': 0, 'B': 1, 'C': 2} # 0=T, 1=L1, 2=L2
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
    """Comprueba si es el primer o último día de un periodo nocturno."""
    for start, end in night_periods:
        if date_obj == start or date_obj == end:
            return True
    return False

def is_in_night_period(day_idx, year, night_periods):
    """Comprueba si un día cae DENTRO de un periodo nocturno (para pintar gris)."""
    current_date = datetime.date(year, 1, 1) + datetime.timedelta(days=day_idx)
    for start, end in night_periods:
        if start <= current_date <= end:
            return True
    return False

def validate_and_generate(roster_df, requests, year, night_periods):
    base_schedule_turn, total_days = generate_base_schedule(year)
    final_schedule = {} 
    coverage_counters = {name: 0 for name in roster_df['Nombre']}
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    
    # Procesar solicitudes
    for req in requests:
        name = req['Nombre']
        start = req['Inicio']
        end = req['Fin']
        start_idx = start.timetuple().tm_yday - 1
        end_idx = end.timetuple().tm_yday - 1
        
        for d in range(start_idx, end_idx + 1):
            if final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V'
            else:
                final_schedule[name][d] = 'V(L)'

    errors = []
    
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
            if p1['Rol'] == p2['Rol'] and p1['Rol'] != "Bombero":
                # Permitimos si hay cobertura, la lógica de abajo decidirá
                pass 

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
                if prev_day.startswith('T') and prev_prev.startswith('T'):
                    continue
                valid_candidates.append(cand)
                
            if not valid_candidates:
                errors.append(f"Día {d+1}: Falta cobertura para {name_missing} (Regla Máx 2T).")
                continue
                
            valid_candidates.sort(key=lambda x: coverage_counters[x])
            chosen = valid_candidates[0]
            
            # Solo marcamos la letra del turno (A, B, C) para ahorrar espacio en Excel
            final_schedule[chosen][d] = f"T*({person_row['Turno']})"
            coverage_counters[chosen] += 1

    # Relleno Administrativo
    fill_log = {} 
    for name in roster_df['Nombre']:
        current_v_days = [i for i, x in enumerate(final_schedule[name]) if x.startswith('V')]
        needed = 39 - len(current_v_days)
        added_dates = []
        if needed > 0:
            available_idx = [i for i, x in enumerate(final_schedule[name]) if x == 'L']
            # Filtrar días que sean NOCTURNOS para no poner V de relleno ahí (opcional, pero mejor)
            # Dejamos que el relleno caiga donde sea, es administrativo.
            if len(available_idx) >= needed:
                fill_idxs = random.sample(available_idx, needed)
                fill_idxs.sort()
                for idx in fill_idxs:
                    final_schedule[name][idx] = 'V(R)'
                    d_obj = datetime.date(year, 1, 1) + datetime.timedelta(days=idx)
                    added_dates.append(d_obj)
        fill_log[name] = added_dates

    return final_schedule, errors, coverage_counters, fill_log

# -------------------------------------------------------------------
# 2. GENERACIÓN EXCEL (MATRIZ Y ESTADÍSTICAS MEJORADAS)
# -------------------------------------------------------------------
def create_excel(schedule, roster_df, year, requests, fill_log, counters, night_periods):
    wb = Workbook()
    
    # Estilos
    s_T = PatternFill("solid", fgColor="C6EFCE") # Verde
    s_V = PatternFill("solid", fgColor="FFEB9C") # Amarillo
    s_VR = PatternFill("solid", fgColor="FFFFE0") # Amarillo Claro
    s_Cov = PatternFill("solid", fgColor="FFC7CE") # Rojo Claro
    s_L = PatternFill("solid", fgColor="F2F2F2") # Gris Claro
    s_Night = PatternFill("solid", fgColor="A6A6A6") # Gris Oscuro (Nocturna)
    
    font_bold = Font(bold=True)
    font_red = Font(color="9C0006", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_thin = Side(border_style="thin", color="000000")
    border_med = Side(border_style="medium", color="000000")
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    border_box = Border(left=border_med, right=border_med, top=border_med, bottom=border_med)

    # --- HOJA 1: CUADRANTE (MATRIZ ANUAL POR PERSONA) ---
    ws1 = wb.active
    ws1.title = "Cuadrante"
    
    # Configurar anchos
    ws1.column_dimensions['A'].width = 15
    for i in range(2, 34):
        ws1.column_dimensions[get_column_letter(i)].width = 4

    current_row = 1
    
    # Agrupar por Turno para orden visual
    for t in TEAMS:
        # Título del Turno
        ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=32)
        cell_title = ws1.cell(current_row, 1, f"TURNO {t}")
        cell_title.font = Font(bold=True, size=14, color="FFFFFF")
        cell_title.fill = PatternFill("solid", fgColor="000080") # Azul Oscuro
        cell_title.alignment = align_c
        current_row += 2

        team_members = roster_df[roster_df['Turno'] == t]
        
        for _, p in team_members.iterrows():
            name = p['Nombre']
            role = p['Rol']
            
            # Cabecera Persona
            ws1.cell(current_row, 1, f"{name} ({role})").font = font_bold
            
            # Cabecera Días (1-31)
            for d in range(1, 32):
                c = ws1.cell(current_row, d+1, d)
                c.alignment = align_c
                c.font = font_bold
                c.border = border_all
                c.fill = PatternFill("solid", fgColor="E0E0E0")
            
            current_row += 1
            
            # Filas Meses (1-12)
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
                        
                        # Lógica de Estilos
                        val = ""
                        fill = s_L # Default Libra
                        
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
                            # Formato T*(A) -> Ponemos solo A en rojo
                            val = status.split('(')[1][0] 
                            fill = s_Cov
                            cell.font = font_red
                        
                        # SUPERPOSICIÓN NOCTURNA (Gris Oscuro si coincide)
                        if is_in_night_period(day_of_year, year, night_periods):
                            fill = s_Night
                            # Si hay texto (T, V, etc), se mantiene, pero el fondo cambia
                        
                        cell.value = val
                        cell.fill = fill
                    else:
                        # Día inexistente
                        cell.fill = PatternFill("solid", fgColor="808080")
                
                current_row += 1
            current_row += 2 # Espacio entre personas

    # --- HOJA 2: ESTADÍSTICAS MAQUETADAS ---
    ws2 = wb.create_sheet("Estadísticas")
    
    # Anchos
    ws2.column_dimensions['A'].width = 20 # Nombre
    ws2.column_dimensions['B'].width = 10 # Turno
    ws2.column_dimensions['C'].width = 15 # Rol
    ws2.column_dimensions['D'].width = 18 # Creditos
    ws2.column_dimensions['E'].width = 18 # Cobertura
    ws2.column_dimensions['F'].width = 18 # Total Trab
    ws2.column_dimensions['G'].width = 20 # Total Vacs
    
    # Título
    ws2.merge_cells('A1:G1')
    ws2['A1'] = f"RESUMEN DE HORAS Y VACACIONES - AÑO {year}"
    ws2['A1'].font = Font(bold=True, size=14)
    ws2['A1'].alignment = align_c
    
    # Cabecera Tabla
    headers = ["Nombre", "Turno", "Puesto", "Gastado (T)", "Coberturas (T*)", "Total Días (T+T*)", "Total Vacs (Nat)"]
    for col, h in enumerate(headers, 1):
        c = ws2.cell(3, col, h)
        c.font = font_bold
        c.fill = PatternFill("solid", fgColor="BDD7EE") # Azulito
        c.border = border_box
        c.alignment = align_c

    row = 4
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        sch = schedule[name]
        
        # Cálculos
        t_base = sum(1 for x in sch if x == 'T' or x.startswith('T*') and not 'V' in x) # Simplificación conteo
        # Mejor: Contar exactamente del array
        real_t_base = 0
        # Recalcular T base originales
        base_sch_turn, _ = generate_base_schedule(year)
        original_ts = base_sch_turn[p['Turno']].count('T')
        
        # Créditos gastados = Original T - T actuales (sin contar coberturas)
        # Pero es más fácil: contar cuantas 'V' (vacaciones en dia de trabajo) hay
        v_credits = sch.count('V')
        t_cover = counters[name]
        
        # Total Trabajo = (Original T - V) + T*
        total_work = (original_ts - v_credits) + t_cover
        
        v_natural = sch.count('V') + sch.count('V(L)') + sch.count('V(R)')
        
        data = [name, p['Turno'], p['Rol'], v_credits, t_cover, total_work, v_natural]
        
        for col, val in enumerate(data, 1):
            c = ws2.cell(row, col, val)
            c.border = border_all
            c.alignment = align_c
            # Zebra striping
            if row % 2 == 0: c.fill = s_L
            
        row += 1

    # --- HOJA 3: RESUMEN TEXTO ---
    ws3 = wb.create_sheet("Resumen Solicitudes")
    ws3.column_dimensions['A'].width = 25
    ws3.column_dimensions['D'].width = 60
    ws3.append(["Nombre", "Turno", "Rol", "Periodos Solicitados"])
    
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        person_reqs = [f"{r['Inicio'].strftime('%d/%m')} al {r['Fin'].strftime('%d/%m')}" for r in requests if r['Nombre'] == name]
        req_str = " | ".join(person_reqs) if person_reqs else "Sin solicitudes"
        ws3.append([name, p['Turno'], p['Rol'], req_str])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V3.0")

st.title("🚒 Gestor Integral V3.0 (Matriz + Nocturnas)")

# 1. CONFIGURACIÓN Y NOCTURNAS
c1, c2 = st.columns([2, 1])

with c1:
    with st.expander("1. Configuración de Plantilla", expanded=False):
        if 'roster_data' not in st.session_state:
            st.session_state.roster_data = pd.DataFrame(DEFAULT_ROSTER)
        
        edited_df = st.data_editor(
            st.session_state.roster_data,
            column_config={
                "Turno": st.column_config.SelectboxColumn(options=TEAMS, required=True),
                "Rol": st.column_config.SelectboxColumn(options=ROLES, required=True),
                "SV": st.column_config.CheckboxColumn(label="¿Es SV?", help="Puede cubrir conductor")
            },
            num_rows="dynamic",
            use_container_width=True
        )
        st.session_state.roster_data = edited_df

with c2:
    with st.expander("🌑 Periodos Nocturnos (Gris)", expanded=True):
        st.caption("No se puede pedir vacaciones el primer ni último día.")
        if 'nights' not in st.session_state: st.session_state.nights = []
        
        dn_start = st.date_input("Inicio Noche", value=None)
        dn_end = st.date_input("Fin Noche", value=None)
        
        if st.button("Añadir Periodo Nocturno"):
            if dn_start and dn_end:
                if dn_start > dn_end:
                    st.error("Fecha fin anterior a inicio")
                else:
                    st.session_state.nights.append((dn_start, dn_end))
                    st.success("Añadido")
        
        # Listar
        if st.session_state.nights:
            st.write("Periodos activos:")
            for i, (s, e) in enumerate(st.session_state.nights):
                c_del, c_txt = st.columns([1, 4])
                if c_del.button("X", key=f"dn_{i}"):
                    st.session_state.nights.pop(i)
                    st.rerun()
                c_txt.text(f"{s.strftime('%d/%m')} - {e.strftime('%d/%m')}")

# 2. GESTOR DE VACACIONES
st.divider()
col_req, col_prev = st.columns([1, 2])

names_list = edited_df['Nombre'].tolist()
today = datetime.date.today()
year_val = st.number_input("Año", value=today.year + 1)

with col_req:
    st.subheader("2. Añadir Solicitud")
    sel_name = st.selectbox("Trabajador", names_list)
    
    # Selector Fechas
    d_range = st.date_input("Rango Vacaciones", [], help="Elige inicio y fin")
    
    if st.button("Añadir Periodo"):
        if len(d_range) == 2:
            start, end = d_range
            
            # VALIDACIÓN NOCTURNA INMEDIATA
            conflict = False
            if is_night_restricted(start, st.session_state.nights) or is_night_restricted(end, st.session_state.nights):
                st.error("⛔ Error: No puedes empezar ni terminar vacaciones en el primer/último día de un periodo nocturno.")
                conflict = True
            
            if not conflict:
                if 'requests' not in st.session_state: st.session_state.requests = []
                st.session_state.requests.append({"Nombre": sel_name, "Inicio": start, "Fin": end})
                st.success(f"Añadido: {sel_name}")
        else:
            st.warning("Selecciona fecha inicio y fin.")

with col_prev:
    st.subheader(f"Calendario de Trabajo: {sel_name}")
    # Visualizador de ayuda
    if sel_name:
        row_p = edited_df[edited_df['Nombre'] == sel_name].iloc[0]
        base_sch, _ = generate_base_schedule(year_val)
        my_sch = base_sch[row_p['Turno']]
        
        # Crear un mini dataframe para visualizar el mes seleccionado en el date input (o Enero por defecto)
        # Truco: mostrar todo el año es mucho, mostremos los meses relevantes si hay rango
        view_months = [1, 2, 3] # Default
        if len(d_range) > 0:
            view_months = [d_range[0].month]
            if len(d_range) > 1 and d_range[1].month != d_range[0].month:
                view_months.append(d_range[1].month)
        
        # Pintar mini calendario horizontal
        # Creamos HTML simple para visualizar T (Verde) y L (Gris)
        html_cal = "<div style='display:flex; flex-wrap:wrap; gap:10px;'>"
        for m in view_months:
            html_cal += f"<div style='border:1px solid #ddd; padding:5px; border-radius:5px;'><strong>{MESES[m-1]}</strong><br>"
            days_in_m = calendar.monthrange(year_val, m)[1]
            html_cal += "<div style='display:grid; grid-template-columns:repeat(7, 1fr); gap:2px; font-size:10px; text-align:center;'>"
            for d in range(1, days_in_m + 1):
                dt = datetime.date(year_val, m, d)
                d_idx = dt.timetuple().tm_yday - 1
                status = my_sch[d_idx]
                color = "#C6EFCE" if status == 'T' else "#F2F2F2"
                # Borde si es nocturno
                border = "2px solid #555" if is_in_night_period(d_idx, year_val, st.session_state.nights) else "none"
                
                html_cal += f"<div style='background-color:{color}; padding:2px; border:{border}'>{d}</div>"
            html_cal += "</div></div>"
        html_cal += "</div>"
        st.markdown(html_cal, unsafe_allow_html=True)
        st.caption("Verde = Trabajo (T), Gris = Libre (L). Borde Oscuro = Periodo Nocturno.")

# Lista Solicitudes
st.divider()
if 'requests' in st.session_state and st.session_state.requests:
    st.write("📝 Solicitudes Pendientes:")
    for i, r in enumerate(st.session_state.requests):
        cols = st.columns([4, 1])
        cols[0].write(f"{r['Nombre']}: {r['Inicio'].strftime('%d/%m')} - {r['Fin'].strftime('%d/%m')}")
        if cols[1].button("🗑️", key=f"del_{i}"):
            st.session_state.requests.pop(i)
            st.rerun()

# 3. GENERACIÓN
st.divider()
if st.button("🚀 Generar Cuadrante (Excel)", type="primary", use_container_width=True):
    if 'requests' not in st.session_state or not st.session_state.requests:
        st.error("Añade solicitudes primero.")
    else:
        final_sch, errs, counters, fill_log = validate_and_generate(
            edited_df, st.session_state.requests, year_val, st.session_state.nights
        )
        
        if errs:
            st.error("❌ Conflictos detectados:")
            for e in errs: st.write(f"- {e}")
        else:
            st.success("✅ Todo correcto.")
            excel_data = create_excel(
                final_sch, edited_df, year_val, st.session_state.requests, fill_log, counters, st.session_state.nights
            )
            st.download_button("📥 Descargar Excel V3.0", excel_data, f"Cuadrante_V3_{year_val}.xlsx")
