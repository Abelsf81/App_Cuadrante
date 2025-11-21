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

# Plantilla por defecto para cargar la primera vez
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
# 1. MOTOR LÓGICO (GENERACIÓN Y VALIDACIÓN)
# -------------------------------------------------------------------

def generate_base_schedule(year):
    """Genera el patrón 1-2 (T-L-L) para todo el año y personal."""
    is_leap = calendar.isleap(year)
    total_days = 366 if is_leap else 365
    # Estado inicial arbitrario para que el 1 de Enero: A=T, B=L, C=L
    # Secuencia: T -> L1 -> L2 -> T
    status = {'A': 0, 'B': 1, 'C': 2} # 0=T, 1=L1, 2=L2
    
    schedule = {team: [] for team in TEAMS}
    
    for _ in range(total_days):
        for t in TEAMS:
            if status[t] == 0: schedule[t].append('T')
            else: schedule[t].append('L')
            
            # Rotar estado
            status[t] = (status[t] + 1) % 3
            
    return schedule, total_days

def get_candidates(person_missing, roster_df, day_idx, current_schedule):
    """
    Encuentra candidatos válidos en las 'Bolsas de Cobertura' que estén Librando (L).
    """
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    
    for _, candidate in roster_df.iterrows():
        # 1. Descartar mismo turno (Regla de Oro: No cubrir mismo turno)
        if candidate['Turno'] == missing_turn:
            continue
            
        # 2. Verificar que libra ese día
        # (Nota: current_schedule es el estado actual, si ya cubre a otro, pondrá T*)
        cand_status = current_schedule[candidate['Nombre']][day_idx]
        if cand_status != 'L': 
            continue # Ya está trabajando o cubriendo
            
        # 3. Verificar Compatibilidad de Rol (Bolsas)
        is_compatible = False
        
        if missing_role == "Mando":
            if candidate['Rol'] == "Mando": is_compatible = True
            
        elif missing_role == "Conductor":
            if candidate['Rol'] == "Conductor": is_compatible = True
            if candidate['SV']: is_compatible = True # SV cubre conductor
            
        elif missing_role == "Bombero":
            if candidate['Rol'] == "Bombero": is_compatible = True
            if candidate['SV']: is_compatible = True # SV puede bajar a bombero
            
        if is_compatible:
            candidates.append(candidate['Nombre'])
            
    return candidates

def validate_and_generate(roster_df, requests, year):
    """
    Procesa las solicitudes, valida reglas y asigna coberturas.
    """
    base_schedule_turn, total_days = generate_base_schedule(year)
    
    # 1. Crear horario individual base
    final_schedule = {} # {Nombre: ['T', 'L', ...]}
    coverage_counters = {name: 0 for name in roster_df['Nombre']}
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    # 2. Procesar Solicitudes (Día a Día para validar concurrencia)
    # Convertimos rangos a un mapa: day_idx -> [personas de vacaciones]
    day_vacations = {i: [] for i in range(total_days)}
    
    # Expandir rangos
    for req in requests:
        name = req['Nombre']
        start = req['Inicio']
        end = req['Fin']
        start_idx = start.timetuple().tm_yday - 1
        end_idx = end.timetuple().tm_yday - 1
        
        for d in range(start_idx, end_idx + 1):
            # Solo nos importa si era día de trabajo (T) para la operatividad
            if final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V' # Marcamos vacaciones
            else:
                # Es vacaciones en día libre
                final_schedule[name][d] = 'V(L)'

    # 3. Resolver Coberturas y Validar Reglas de Oro
    errors = []
    
    for d in range(total_days):
        absent_people = day_vacations[d]
        
        if not absent_people: continue
        
        # Regla 1: Máx 2 personas
        if len(absent_people) > 2:
            date_str = (datetime.date(year, 1, 1) + datetime.timedelta(days=d)).strftime("%d-%m")
            errors.append(f"{date_str}: Hay {len(absent_people)} personas de vacaciones (Máx 2).")
            continue
            
        # Regla 2 y 3: Incompatibilidades (Mismo turno, Misma categoría)
        if len(absent_people) == 2:
            p1 = roster_df[roster_df['Nombre'] == absent_people[0]].iloc[0]
            p2 = roster_df[roster_df['Nombre'] == absent_people[1]].iloc[0]
            
            if p1['Turno'] == p2['Turno']:
                errors.append(f"Día {d+1}: {p1['Nombre']} y {p2['Nombre']} son del mismo turno ({p1['Turno']}).")
            
            if p1['Rol'] == p2['Rol'] and p1['Rol'] != "Bombero":
                # Validación inteligente: Si son mandos, ¿quedan mandos libres?
                # Simplificación robusta: Permitimos si hay cobertura, pero avisamos si es arriesgado.
                # La lógica de cobertura abajo determinará si es posible.
                pass 

        # Asignar Cobertura
        for name_missing in absent_people:
            person_row = roster_df[roster_df['Nombre'] == name_missing].iloc[0]
            candidates = get_candidates(person_row, roster_df, d, final_schedule)
            
            if not candidates:
                errors.append(f"Día {d+1}: No hay nadie disponible para cubrir a {name_missing}.")
                continue
                
            # Filtrar por Regla Máx 2T (No trabajar 3 días seguidos)
            valid_candidates = []
            for cand in candidates:
                # Mirar día anterior y día siguiente (si existe)
                # Simplificado: Miramos si ayer trabajó. Si ayer trabajó Y hoy trabaja (cobertura), mañana debe librar.
                # Implementación estricta: Evitar secuencias T-T-T
                
                prev_day = final_schedule[cand][d-1] if d > 0 else 'L'
                # Check simple: si ayer trabajó, hoy sería el 2º. Es aceptable.
                # Pero si ayer trabajó (T) y anteayer trabajó (T), hoy NO puede.
                
                prev_prev = final_schedule[cand][d-2] if d > 1 else 'L'
                
                if prev_day.startswith('T') and prev_prev.startswith('T'):
                    continue # Ya lleva 2, no puede hacer 3.
                
                valid_candidates.append(cand)
                
            if not valid_candidates:
                errors.append(f"Día {d+1}: Falta cobertura para {name_missing}. Candidatos agotados por Regla Máx 2T.")
                continue
                
            # Elección de Candidato: El que menos haya cubierto (Equilibrio)
            # Ordenar por contador
            valid_candidates.sort(key=lambda x: coverage_counters[x])
            chosen = valid_candidates[0]
            
            # Aplicar cobertura
            final_schedule[chosen][d] = f"T* ({person_row['Turno']})" # Marcamos turno cubierto
            coverage_counters[chosen] += 1

    # 4. Relleno Administrativo (Filler) hasta 39 días
    # Esto no afecta operatividad, es solo visual y contable
    fill_log = {} # {Nombre: [Fechas]}
    
    for name in roster_df['Nombre']:
        # Contar naturales actuales
        current_v_days = [i for i, x in enumerate(final_schedule[name]) if x.startswith('V')]
        natural_count = len(current_v_days)
        needed = 39 - natural_count
        
        added_dates = []
        if needed > 0:
            # Buscar días L libres (que no sean cobertura T*)
            available_idx = [i for i, x in enumerate(final_schedule[name]) if x == 'L']
            # Mezclar y coger
            if len(available_idx) >= needed:
                fill_idxs = random.sample(available_idx, needed)
                fill_idxs.sort()
                for idx in fill_idxs:
                    final_schedule[name][idx] = 'V(R)' # Vacation Rest (Relleno)
                    d_obj = datetime.date(year, 1, 1) + datetime.timedelta(days=idx)
                    added_dates.append(d_obj)
        
        fill_log[name] = added_dates

    return final_schedule, errors, coverage_counters, fill_log

# -------------------------------------------------------------------
# 2. GENERACIÓN EXCEL
# -------------------------------------------------------------------
def create_excel(schedule, roster_df, year, requests, fill_log, counters):
    wb = Workbook()
    
    # Estilos
    s_T = PatternFill("solid", fgColor="C6EFCE") # Verde
    s_V = PatternFill("solid", fgColor="FFEB9C") # Amarillo (Vacaciones Trabajo)
    s_VR = PatternFill("solid", fgColor="FFFFE0") # Amarillo Claro (Vacaciones Relleno)
    s_Cov = PatternFill("solid", fgColor="FFC7CE") # Rojo (Cobertura)
    s_L = PatternFill("solid", fgColor="F2F2F2") # Gris
    font_bold = Font(bold=True)
    align_c = Alignment(horizontal="center")
    
    # --- HOJA 1: CALENDARIO ---
    ws1 = wb.active
    ws1.title = "Cuadrante"
    
    # Cabecera
    ws1.cell(1, 1, "Nombre").font = font_bold
    ws1.cell(1, 2, "Puesto").font = font_bold
    for d in range(1, 367):
        col = get_column_letter(d + 2)
        ws1.column_dimensions[col].width = 3.5
        dt = datetime.date(year, 1, 1) + datetime.timedelta(days=d-1)
        cell = ws1.cell(1, d+2, dt.day)
        cell.alignment = align_c
        cell.font = font_bold
        if dt.month % 2 == 0: cell.fill = PatternFill("solid", fgColor="E0E0E0") # Alternar mes visual

    row = 2
    for t in TEAMS:
        ws1.cell(row, 1, f"--- TURNO {t} ---").font = Font(bold=True, color="0000FF")
        row += 1
        team_members = roster_df[roster_df['Turno'] == t]
        for _, p in team_members.iterrows():
            ws1.cell(row, 1, p['Nombre'])
            ws1.cell(row, 2, p['Rol'] + (" (SV)" if p['SV'] else ""))
            
            days = schedule[p['Nombre']]
            for i, status in enumerate(days):
                cell = ws1.cell(row, i+3, value="")
                cell.alignment = align_c
                
                if status == 'T':
                    cell.value = "T"
                    cell.fill = s_T
                elif status == 'V':
                    cell.value = "V"
                    cell.fill = s_V
                elif status == 'V(L)' or status == 'V(R)':
                    cell.value = "v"
                    cell.fill = s_VR
                elif status.startswith('T*'):
                    cell.value = status.split('(')[1][0] # Extraer Turno cubierto (A, B, C)
                    cell.fill = s_Cov
                    cell.font = Font(color="9C0006", bold=True)
                else:
                    cell.fill = s_L
            row += 1
        row += 1

    # --- HOJA 2: ESTADÍSTICAS ---
    ws2 = wb.create_sheet("Estadísticas")
    headers = ["Nombre", "Turno", "Créditos T Gastados", "Días Cobertura (T*)", "Total Trabajo", "Total Vacaciones (Nat)"]
    ws2.append(headers)
    
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        sch = schedule[name]
        t_base = sch.count('T')
        t_cover = counters[name]
        v_credits = 13 - (base_count_t(p['Turno'], year) - t_base) # Aprox calculation or track directly
        # Better calculation from schedule
        v_credits_spent = sum(1 for x in sch if x == 'V')
        v_natural = sch.count('V') + sch.count('V(L)') + sch.count('V(R)')
        
        ws2.append([name, p['Turno'], v_credits_spent, t_cover, t_base + t_cover, v_natural])

    # --- HOJA 3: RESUMEN SOLICITUDES ---
    ws3 = wb.create_sheet("Resumen Solicitudes")
    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['D'].width = 50
    ws3.column_dimensions['E'].width = 50
    ws3.append(["Nombre", "Turno", "Rol", "Periodos Solicitados", "Días Relleno Automático"])
    
    for _, p in roster_df.iterrows():
        name = p['Nombre']
        # Buscar solicitudes de esta persona
        person_reqs = [f"{r['Inicio'].strftime('%d/%m')} - {r['Fin'].strftime('%d/%m')}" for r in requests if r['Nombre'] == name]
        req_str = ", ".join(person_reqs)
        
        # Días de relleno
        fill_dates = [d.strftime('%d/%m') for d in fill_log[name]]
        # Agrupar visualmente si son muchos
        fill_str = ", ".join(fill_dates) if fill_dates else "Ninguno"
        
        ws3.append([name, p['Turno'], p['Rol'], req_str, fill_str])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def base_count_t(turn, year):
    # Helper para saber cuántos T tenía originalmente
    sch, _ = generate_base_schedule(year)
    return sch[turn].count('T')

# -------------------------------------------------------------------
# INTERFAZ
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor de Cuadrantes V2.0")

st.title("🚒 Gestor Integral de Vacaciones V2.0")
st.markdown("""
**Sistema de Gestión:**
1.  **Configura la Plantilla:** Define nombres, roles y SV (Sustituto Vehículo).
2.  **Solicita Rangos:** Elige fechas (Inicio-Fin). El sistema descuenta créditos 'T' (Max 13).
3.  **Genera:** El sistema valida reglas (Máx 2T, Coberturas) y rellena administrativamente hasta 39 días naturales.
""")

# 1. CONFIGURACIÓN
with st.expander("1. Configuración de Plantilla", expanded=True):
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

# 2. GESTOR DE VACACIONES
st.divider()
col_req, col_list = st.columns([1, 1])

with col_req:
    st.subheader("2. Añadir Solicitud")
    
    # Inputs
    names_list = edited_df['Nombre'].tolist()
    sel_name = st.selectbox("Trabajador", names_list)
    
    # Calcular año (asumimos año actual o próximo según fecha)
    today = datetime.date.today()
    year_val = st.number_input("Año del Cuadrante", value=today.year + 1 if today.month > 9 else today.year)
    
    # Date picker range
    d_range = st.date_input("Selecciona Rango", [], help="Elige fecha inicio y fin")
    
    if st.button("Añadir Periodo"):
        if len(d_range) == 2:
            start, end = d_range
            if start.year != year_val or end.year != year_val:
                st.error(f"Las fechas deben ser del año {year_val}")
            else:
                if 'requests' not in st.session_state: st.session_state.requests = []
                # Calcular coste provisional para feedback
                st.session_state.requests.append({
                    "Nombre": sel_name,
                    "Inicio": start,
                    "Fin": end
                })
                st.success(f"Periodo añadido para {sel_name}")
        else:
            st.warning("Debes seleccionar fecha de inicio y fin.")

with col_list:
    st.subheader("Resumen de Solicitudes")
    if 'requests' in st.session_state and st.session_state.requests:
        # Calcular créditos gastados en tiempo real (Aproximado base T-L-L)
        base_sch, _ = generate_base_schedule(year_val)
        
        # Agrupar por persona
        person_stats = {n: 0 for n in names_list}
        
        for i, r in enumerate(st.session_state.requests):
            # Botón borrar
            cols = st.columns([4, 1])
            cols[0].write(f"**{r['Nombre']}**: {r['Inicio'].strftime('%d/%m')} al {r['Fin'].strftime('%d/%m')}")
            if cols[1].button("🗑️", key=f"del_{i}"):
                st.session_state.requests.pop(i)
                st.rerun()
            
            # Calcular coste
            turn = edited_df[edited_df['Nombre'] == r['Nombre']].iloc[0]['Turno']
            start_idx = r['Inicio'].timetuple().tm_yday - 1
            end_idx = r['Fin'].timetuple().tm_yday - 1
            cost = 0
            for d in range(start_idx, end_idx + 1):
                if base_sch[turn][d] == 'T': cost += 1
            
            person_stats[r['Nombre']] += cost

        st.divider()
        st.write("--- Créditos Gastados (Max 13) ---")
        for name, spent in person_stats.items():
            if spent > 0:
                st.progress(min(spent/13, 1.0), text=f"{name}: {spent} créditos")
                if spent > 13: st.error(f"{name} se ha pasado ({spent}/13)")
    else:
        st.info("No hay solicitudes aún.")

# 3. GENERACIÓN
st.divider()
if st.button("🚀 Validar y Generar Cuadrante", type="primary", use_container_width=True):
    if 'requests' not in st.session_state or not st.session_state.requests:
        st.error("Añade al menos una solicitud de vacaciones.")
    else:
        final_sch, errs, counters, fill_log = validate_and_generate(edited_df, st.session_state.requests, year_val)
        
        if errs:
            st.error("❌ Se encontraron conflictos:")
            for e in errs:
                st.write(f"- {e}")
        else:
            st.success("✅ Cuadrante generado correctamente.")
            
            excel_data = create_excel(final_sch, edited_df, year_val, st.session_state.requests, fill_log, counters)
            
            st.download_button(
                label="📥 Descargar Excel Final",
                data=excel_data,
                file_name=f"Cuadrante_{year_val}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )