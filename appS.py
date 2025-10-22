import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import datetime
import io
import sys
import calendar 
from collections import defaultdict

# --- Constantes ---
TEAMS = ['A', 'B', 'C']
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


# -------------------------------------------------------------------
# LÓGICA DE GENERACIÓN (CORREGIDA)
# -------------------------------------------------------------------

def generate_base_schedule(total_days):
    """
    Genera el cuadrante base (T-L-L) para 365 o 366 días.
    """
    base_schedule = {team: [] for team in TEAMS}
    internal_status = {'A': 'L2', 'B': 'L1', 'C': 'T'} 
    for day in range(1, total_days + 1):
        new_status = {}
        for team in TEAMS:
            prev = internal_status[team]
            if prev == 'L2': new_status[team] = 'T'
            elif prev == 'T': new_status[team] = 'L1'
            elif prev == 'L1': new_status[team] = 'L2'
        internal_status = new_status
        for team in TEAMS:
            if internal_status[team] == 'T':
                base_schedule[team].append('T') # 'T' = Trabajo Base
            else:
                base_schedule[team].append('L') # 'L' = Libra
    return base_schedule


# --- ¡FUNCIÓN CORREGIDA! ---
def generate_schedule_interchange(base_schedule, vacation_requests_map, total_days):
    """
    Aplica el nuevo modelo de "intercambio" sobre el cuadrante base.
    Lógica de equilibrio 13-13-13 (desempate rotativo) corregida.
    """
    
    final_schedule = {team: list(days) for team, days in base_schedule.items()}
    # Contador de T* TOTALES (la única verdad)
    coverage_counts = {'A': 0, 'B': 0, 'C': 0}

    for day_index in range(total_days):
        day_of_year = day_index + 1
        team_on_v = vacation_requests_map.get(day_of_year)
        
        if team_on_v:
            final_schedule[team_on_v][day_index] = 'V' 
            # Obtener los dos candidatos a cubrir
            cover_teams = [team for team in TEAMS if team != team_on_v]
            cand1 = cover_teams[0]
            cand2 = cover_teams[1]
            
            # --- NOVEDAD: Lógica de desempate rotativo ---
            count1 = coverage_counts[cand1]
            count2 = coverage_counts[cand2]

            if count1 < count2:
                cover_team = cand1
            elif count2 < count1:
                cover_team = cand2
            else:
                # ¡EMPATE! Aplicar desempate rotativo
                # (A>B, B>C, C>A)
                if team_on_v == 'A': # Candidatos B, C. B gana el empate.
                    cover_team = 'B' if 'B' in cover_teams else 'C' 
                elif team_on_v == 'B': # Candidatos A, C. C gana el empate.
                    cover_team = 'C' if 'C' in cover_teams else 'A'
                elif team_on_v == 'C': # Candidatos A, B. A gana el empate.
                    cover_team = 'A' if 'A' in cover_teams else 'B'
            
            # Incrementar el contador T* TOTAL del que cubre
            coverage_counts[cover_team] += 1
            # --- FIN NOVEDAD ---

            final_schedule[cover_team][day_index] = f"T({team_on_v})" 
            
    total_covered = sum(coverage_counts.values())
    if total_covered != 39:
        st.error(f"Error de lógica interna: Se han cubierto {total_covered} días en vez de 39.")
        return None
    
    st.success(f"Cuadrante generado. Coberturas: A={coverage_counts['A']}T, B={coverage_counts['B']}T, C={coverage_counts['C']}T.")
    
    # Devuelve tanto el horario como los contadores
    return final_schedule, coverage_counts


# -------------------------------------------------------------------
# FUNCIÓN DE EXCEL (Sin cambios)
# -------------------------------------------------------------------

def create_calendar_xlsx_in_memory(schedule, names, year, coverage_counts):
    """
    Crea el archivo Excel en memoria con formato de matriz de calendario,
    colores y pestañas individuales Y hoja de estadísticas.
    """
    wb = Workbook()
    
    # --- Definición de Estilos ---
    STYLE_T = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Verde (Trabajo Base)
    STYLE_L = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # Gris (Libra)
    STYLE_V = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # Amarillo (Vacaciones)
    STYLE_EMPTY = PatternFill(start_color="808080", end_color="808080", fill_type="solid") # Gris oscuro
    STYLE_COVERAGE = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Rojo Claro
    FONT_COVERAGE = Font(color="9C0006", bold=True) # Fuente Rojo Oscuro
    FONT_BLACK = Font(color="000000")
    FONT_BOLD = Font(bold=True)
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
    
    styles = {'T': STYLE_T, 'L': STYLE_L, 'V': STYLE_V}

    def set_calendar_widths(ws):
        ws.column_dimensions['A'].width = 12
        for i in range(2, 33):
            ws.column_dimensions[get_column_letter(i)].width = 4

    def create_calendar_grid(ws, start_row, team_schedule, year):
        ws.cell(row=start_row, column=1, value="Mes").font = FONT_BOLD
        for d in range(1, 32):
            cell = ws.cell(row=start_row, column=d + 1, value=d)
            cell.font = FONT_BOLD
            cell.alignment = ALIGN_CENTER
        
        current_row = start_row + 1
        for m in range(1, 13):
            ws.cell(row=current_row, column=1, value=MESES[m-1]).font = FONT_BOLD
            num_days_in_month = calendar.monthrange(year, m)[1]
            
            for d in range(1, 32):
                cell = ws.cell(row=current_row, column=d + 1)
                
                if d <= num_days_in_month:
                    try:
                        day_of_year = datetime.date(year, m, d).timetuple().tm_yday
                        status = team_schedule[day_of_year - 1]
                        
                        if status in styles:
                            cell.value = status
                            cell.fill = styles.get(status)
                            cell.font = FONT_BLACK
                        elif status.startswith('T('):
                            covered_team = status[2] 
                            cell.value = covered_team 
                            cell.fill = STYLE_COVERAGE 
                            cell.font = FONT_COVERAGE  
                        else:
                            cell.value = "?" 
                            
                    except IndexError:
                        cell.value = "ERR"
                else:
                    cell.fill = STYLE_EMPTY
                
                cell.alignment = ALIGN_CENTER
            current_row += 1
        return current_row 

    # --- 1. Hoja "Cuadrante General" ---
    ws_general = wb.active
    ws_general.title = "Cuadrante General"
    current_row = 1
    for team_id in TEAMS: 
        team_name = names[team_id]
        team_schedule = schedule[team_id]
        
        cell_title = ws_general.cell(row=current_row, column=1, value=team_name)
        cell_title.font = Font(bold=True, size=14)
        ws_general.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=32)
        cell_title.alignment = ALIGN_CENTER
        
        current_row += 1 
        current_row = create_calendar_grid(ws_general, current_row, team_schedule, year)
        current_row += 2 
    set_calendar_widths(ws_general)

    # --- 2. Hojas Individuales ---
    for team_id in TEAMS: 
        team_name = names[team_id]
        team_schedule = schedule[team_id]
        safe_team_name = team_name[:31]
        ws_ind = wb.create_sheet(title=safe_team_name)
        create_calendar_grid(ws_ind, 1, team_schedule, year)
        set_calendar_widths(ws_ind)

    # --- 3. Hoja de Estadísticas ---
    ws_stats = wb.create_sheet(title="Estadísticas")
    ws_stats.column_dimensions['A'].width = 15
    ws_stats.column_dimensions['B'].width = 25
    ws_stats.column_dimensions['C'].width = 18
    ws_stats.column_dimensions['D'].width = 18
    ws_stats.column_dimensions['E'].width = 18
    ws_stats.column_dimensions['F'].width = 20
    ws_stats.column_dimensions['G'].width = 25

    # Cabecera
    headers = ["Turno", "Nombre", "Trabajo Base (T)", "Libra (L)", "Vacaciones (V)", "Cobertura (T*)", "Total Días Trabajo"]
    ws_stats.append(headers)
    for cell in ws_stats[1]:
        cell.font = FONT_BOLD
        cell.fill = STYLE_L # Fondo gris claro

    row = 2
    for team_id in TEAMS: # 'A', 'B', 'C'
        team_name = names[team_id]
        team_schedule = schedule[team_id]
        
        # Contar T, L, V (T* se pasa)
        t_base_count = team_schedule.count('T')
        l_count = team_schedule.count('L')
        v_count = team_schedule.count('V')
        
        # Obtener T* (cobertura) del diccionario
        t_cover_count = coverage_counts[team_id]
        
        # Calcular total T
        t_total = t_base_count + t_cover_count
        
        # Escribir fila
        data_row = [team_id, team_name, t_base_count, l_count, v_count, t_cover_count, t_total]
        ws_stats.append(data_row)
        
        # Estilo de la fila
        ws_stats.cell(row=row, column=1).font = FONT_BOLD
        ws_stats.cell(row=row, column=2).font = FONT_BOLD
        ws_stats.cell(row=row, column=7).font = FONT_BOLD
        row += 1

    # Añadir totales al final
    total_days = len(schedule['A']) # 365 o 366
    ws_stats.append([]) # Fila vacía
    ws_stats.append(["Total Días Año:", total_days])
    ws_stats.cell(row=row+1, column=1).font = FONT_BOLD
    ws_stats.cell(row=row+1, column=2).font = FONT_BOLD


    # --- Guardar en Memoria ---
    mem_file = io.BytesIO()
    wb.save(mem_file)
    mem_file.seek(0)
    return mem_file.getvalue()


# -------------------------------------------------------------------
# APLICACIÓN WEB (Streamlit) (SIN CAMBIOS)
# -------------------------------------------------------------------

st.set_page_config(layout="wide")
st.title("📅 Generador de Cuadrantes 3x3 (Modelo Intercambio)")
st.info("""
**Lógica de Vacaciones (Intercambio):**
1.  **Rotación Base (T-L-L):** La rotación principal de 1 día de trabajo y 2 libres nunca cambia.
2.  **13 Días de Vacaciones (T):** Cada persona elige 13 días "sueltos".
3.  **Validación:** Solo se pueden pedir vacaciones en un día que te tocaba trabajar (T).
4.  **Cobertura:** Cuando A pide (V), uno de los otros dos (B o C) que libraba (L) pasa a trabajar (T) para cubrirle. El sistema equilibra la carga automáticamente.
""")

# --- 1. Configuración ---
st.header("1. Año y Nombres")

col_cfg_1, col_cfg_2, col_cfg_3, col_cfg_4 = st.columns([1, 1, 1, 1])
with col_cfg_1:
    default_year = datetime.date.today().year
    selected_year = st.number_input("Año del cuadrante", min_value=2020, max_value=2050, value=default_year)
with col_cfg_2:
    name_a = st.text_input("Nombre Turno A", "Persona A")
with col_cfg_3:
    name_b = st.text_input("Nombre Turno B", "Persona B")
with col_cfg_4:
    name_c = st.text_input("Nombre Turno C", "Persona C")

names = {'A': name_a, 'B': name_b, 'C': name_c}


# --- 2. Cálculo Base (Automático y cacheado) ---
@st.cache_data
def get_base_data(year):
    is_leap = calendar.isleap(year)
    total_days = 366 if is_leap else 365
    all_dates = [datetime.date(year, 1, 1) + datetime.timedelta(days=i) for i in range(total_days)]
    base_schedule = generate_base_schedule(total_days)
    
    t_days = {team: [] for team in TEAMS}
    date_to_day_index_map = {date: idx for idx, date in enumerate(all_dates)}
    
    for team in TEAMS:
        t_days[team] = [date for date in all_dates if base_schedule[team][date_to_day_index_map[date]] == 'T']
        
    return t_days, total_days, base_schedule, all_dates

t_days, total_days, base_schedule, all_dates_in_year = get_base_data(selected_year)


# --- 3. Selección de Vacaciones (Dinámica) ---
st.header("2. Selección de 13 Días de Vacaciones")
st.warning("Las listas solo muestran los días (T) de cada turno. Debes seleccionar exactamente 13.", icon="⚠️")

if 'v_dates_a' not in st.session_state: st.session_state.v_dates_a = []
if 'v_dates_b' not in st.session_state: st.session_state.v_dates_b = []
if 'v_dates_c' not in st.session_state: st.session_state.v_dates_c = []

options_a = t_days['A']
options_b = t_days['B']
options_c = t_days['C']

col_a, col_b, col_c = st.columns(3)
with col_a:
    st.subheader(f"Días de {name_a}")
    v_dates_a = st.multiselect(f"Selecciona 13 días (T) para {name_a}", options_a, key='v_dates_a', format_func=lambda d: d.strftime("%d-%m-%Y (%a)"))
with col_b:
    st.subheader(f"Días de {name_b}")
    v_dates_b = st.multiselect(f"Selecciona 13 días (T) para {name_b}", options_b, key='v_dates_b', format_func=lambda d: d.strftime("%d-%m-%Y (%a)"))
with col_c:
    st.subheader(f"Días de {name_c}")
    v_dates_c = st.multiselect(f"Selecciona 13 días (T) para {name_c}", options_c, key='v_dates_c', format_func=lambda d: d.strftime("%d-%m-%Y (%a)"))


# --- 4. Validación (Dinámica) ---
st.divider()
st.subheader("Validación (13 días exactos)")
errors = False
all_counts = {}

for team, dates, name in [('A', v_dates_a, name_a), ('B', v_dates_b, name_b), ('C', v_dates_c, name_c)]:
    count = len(dates)
    all_counts[team] = count
    if count > 13:
        st.error(f"❌ {name} tiene {count}/13. Debes eliminar {count-13}.")
        errors = True
    elif count < 13:
        st.warning(f"⚠️ {name} tiene {count}/13. Faltan {13-count}.")
        errors = True
    else:
        st.success(f"✅ {name} tiene 13/13.")

if all_counts['A'] != 13 or all_counts['B'] != 13 or all_counts['C'] != 13:
    errors = True


# --- 5. Botón de Generar ---
st.divider()
submitted = st.button("Generar Cuadrante", type="primary", use_container_width=True, disabled=errors)

if submitted:
    try:
        all_selected_dates = v_dates_a + v_dates_b + v_dates_c
        if len(all_selected_dates) != len(set(all_selected_dates)):
             st.error("¡Error! Hay fechas de vacaciones solapadas. Esto no debería ocurrir si las listas de (T) son correctas.")
             sys.exit()

        vacation_requests_map = {} # Mapa de {dia_del_año: 'A'}
        
        for date_obj in v_dates_a: vacation_requests_map[date_obj.timetuple().tm_yday] = 'A'
        for date_obj in v_dates_b: vacation_requests_map[date_obj.timetuple().tm_yday] = 'B'
        for date_obj in v_dates_c: vacation_requests_map[date_obj.timetuple().tm_yday] = 'C'

        # Capturar los dos valores devueltos
        with st.spinner("Aplicando intercambios de vacaciones y generando Excel..."):
            final_schedule, coverage_counts = generate_schedule_interchange(base_schedule, vacation_requests_map, total_days)

            if final_schedule:
                # Pasar los contadores al Excel
                excel_data = create_calendar_xlsx_in_memory(final_schedule, names, selected_year, coverage_counts)
                
                st.download_button(
                    label="📥 Descargar Cuadrante (Formato Calendario)",
                    data=excel_data,
                    file_name=f"cuadrante_intercambio_{selected_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Ha ocurrido un error inesperado: {e}")