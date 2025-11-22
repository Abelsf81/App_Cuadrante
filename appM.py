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
                errors.append(f"{date_str}: {name_missing} no tiene cobertura válida (Regla Máx 2T).")
                continue
            
            def sort_key(cand_name):
                cand_turn = name_to_turn[cand_name]
                return (
                    turn_coverage_counters[cand_turn],
                    person_coverage_counters[cand_name],
                    random.random()
                )
            
            valid_candidates.sort(key=sort_key)
            chosen = valid_candidates[0]
            chosen_turn = name_to_turn[chosen]
            
            final_schedule[chosen][d] = f"T*({name_missing})"
            adjustments_log.append((d, chosen, name_missing))
            
            turn_coverage_counters[chosen_turn] += 1
            person_coverage_counters[chosen] += 1

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

    return final_schedule, errors, person_coverage_counters, fill_log, adjustments_log

# -------------------------------------------------------------------
# 3. GENERADOR DE INFORMES DE ERROR (NUEVO)
# -------------------------------------------------------------------
def generate_error_report(df_original, errors_dict):
    """
    Genera un Excel con los datos originales pero marcando en ROJO
    las filas que dieron error y añadiendo una columna con el motivo.
    """
    wb = Workbook()
    
    # Estilo Error
    fill_red = PatternFill("solid", fgColor="FFC7CE")
    font_red = Font(color="9C0006")
    
    # Hoja 1: Datos Marcados
    ws1 = wb.active
    ws1.title = "Datos con Errores"
    
    # Cabeceras
    headers = list(df_original.columns) + ["ERROR DETECTADO"]
    ws1.append(headers)
    
    # Escribir datos
    for idx, row in df_original.iterrows():
        # Convertir fila a lista
        row_data = row.tolist()
        
        # Si esta fila tiene error
        if idx in errors_dict:
            row_data.append(errors_dict[idx]) # Añadir mensaje
            ws1.append(row_data)
            # Pintar de rojo la fila actual (ws1.max_row)
            current_row = ws1.max_row
            for col in range(1, len(row_data) + 1):
                cell = ws1.cell(row=current_row, column=col)
                cell.fill = fill_red
                cell.font = font_red
        else:
            row_data.append("OK")
            ws1.append(row_data)

    # Hoja 2: Lista Limpia
    ws2 = wb.create_sheet("Log de Errores")
    ws2.append(["Fila Excel", "Descripción del Error"])
    for idx, msg in errors_dict.items():
        # idx + 2 porque Excel empieza en 1 y tiene cabecera
        ws2.append([f"Fila {idx + 2}", msg])
        
    # Guardar
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# -------------------------------------------------------------------
# INTERFAZ STREAMLIT
# -------------------------------------------------------------------

st.set_page_config(layout="wide", page_title="Gestor V5.3")

st.title("🚒 Gestor Integral V5.3 (Validador Excel)")

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
    with st.expander("🌑 Periodos Nocturnos", expanded=True):
        if 'nights' not in st.session_state: st.session_state.nights = []
        
        c_dn1, c_dn2 = st.columns(2)
        dn_start = c_dn1.date_input("Inicio", value=None, label_visibility="collapsed")
        dn_end = c_dn2.date_input("Fin", value=None, label_visibility="collapsed")
        if st.button("Añadir Periodo"):
            if dn_start and dn_end: st.session_state.nights.append((dn_start, dn_end))
        
        # --- IMPORTACIÓN NOCTURNAS VALIDADA ---
        uploaded_n = st.file_uploader("Sube Excel Nocturnas", type=['xlsx'], key="n_up", label_visibility="collapsed")
        if uploaded_n and st.button("Procesar Nocturnas"):
            try:
                df_n = pd.read_excel(uploaded_n)
                valid_periods = []
                errors_found = {}
                
                for idx, row in df_n.iterrows():
                    # Intentar leer
                    val_s = row.get('Inicio') if 'Inicio' in row else row.iloc[0]
                    val_e = row.get('Fin') if 'Fin' in row else row.iloc[1]
                    
                    if pd.isnull(val_s) or pd.isnull(val_e):
                        errors_found[idx] = "Fechas vacías"
                        continue
                    
                    try:
                        d_s = pd.to_datetime(val_s, dayfirst=True).date()
                        d_e = pd.to_datetime(val_e, dayfirst=True).date()
                        if d_s > d_e:
                            errors_found[idx] = "Fecha Fin anterior a Inicio"
                        else:
                            valid_periods.append((d_s, d_e))
                    except:
                        errors_found[idx] = "Formato de fecha inválido (Use DD/MM/YYYY)"

                if errors_found:
                    st.error(f"⛔ Se encontraron {len(errors_found)} errores en el Excel.")
                    err_file = generate_error_report(df_n, errors_found)
                    st.download_button("📥 Descargar Informe de Errores", err_file, "Errores_Nocturnas.xlsx")
                else:
                    st.session_state.nights.extend(valid_periods)
                    st.success(f"✅ Añadidos {len(valid_periods)} periodos correctamente.")
                    st.rerun()
                    
            except Exception as e: st.error(f"Error crítico leyendo Excel: {e}")
        # --------------------------------------

        with st.container(height=200):
            if st.session_state.nights:
                for i, (s, e) in enumerate(st.session_state.nights):
                    col_del, col_tx = st.columns([1,5])
                    if col_del.button("x", key=f"n_{i}"):
                        st.session_state.nights.pop(i)
                        st.rerun()
                    col_tx.caption(f"{s.strftime('%d/%m')} - {e.strftime('%d/%m')}")
            else:
                st.caption("Sin periodos.")

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
                errors_found = {} # {idx: msg}
                valid_requests = []
                
                for idx, row in df_upload.iterrows():
                    # Identificar trabajador
                    target_name = None
                    if 'ID_Puesto' in row and not pd.isnull(row['ID_Puesto']):
                        match = edited_df[edited_df['ID_Puesto'] == row['ID_Puesto']]
                        if not match.empty: target_name = match.iloc[0]['Nombre']
                    if not target_name and 'Nombre' in row:
                        if row['Nombre'] in names_list: target_name = row['Nombre']
                    
                    if not target_name:
                        errors_found[idx] = "Trabajador no encontrado en plantilla"
                        continue
                    
                    # Procesar periodos
                    row_has_error = False
                    row_error_msg = []
                    temp_reqs = []
                    
                    for i in range(1, 21):
                        col_start = f'Inicio {i}'
                        col_end = f'Fin {i}'
                        if col_start in row and col_end in row:
                            val_start = row[col_start]
                            val_end = row[col_end]
                            
                            if not pd.isnull(val_start) and not pd.isnull(val_end):
                                try:
                                    d_s = pd.to_datetime(val_start, dayfirst=True).date()
                                    d_e = pd.to_datetime(val_end, dayfirst=True).date()
                                    
                                    if d_s > d_e:
                                        row_has_error = True
                                        row_error_msg.append(f"Período {i}: Fin antes que inicio")
                                    elif is_night_restricted(d_s, st.session_state.nights) or is_night_restricted(d_e, st.session_state.nights):
                                        row_has_error = True
                                        row_error_msg.append(f"Período {i}: Choque con Nocturna")
                                    else:
                                        temp_reqs.append({
                                            "Nombre": target_name,
                                            "Inicio": d_s,
                                            "Fin": d_e
                                        })
                                except:
                                    row_has_error = True
                                    row_error_msg.append(f"Período {i}: Formato fecha inválido")
                    
                    if row_has_error:
                        errors_found[idx] = "; ".join(row_error_msg)
                    else:
                        valid_requests.extend(temp_reqs)
                        count += 1

                if errors_found:
                    st.error(f"⛔ Se encontraron errores en {len(errors_found)} filas.")
                    err_file = generate_error_report(df_upload, errors_found)
                    st.download_button("📥 Descargar Informe de Errores", err_file, "Errores_Vacaciones.xlsx")
                else:
                    st.session_state.requests.extend(valid_requests)
                    st.success(f"✅ Importación exitosa: {len(valid_requests)} periodos añadidos.")
                    st.rerun()
                    
            except Exception as e: st.error(f"Error crítico: {e}")

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

# --- LISTA LIMPIA CON ACORDEONES ---
with col_list:
    st.subheader("Listado Solicitudes")
    
    if st.session_state.requests:
        indexed_requests = []
        for i, r in enumerate(st.session_state.requests):
            r_with_index = r.copy()
            r_with_index['idx'] = i
            indexed_requests.append(r_with_index)
        indexed_requests.sort(key=lambda x: x['Nombre'])
        grouped_reqs = {}
        for key, group in groupby(indexed_requests, lambda x: x['Nombre']):
            grouped_reqs[key] = list(group)
        for name, reqs in grouped_reqs.items():
            with st.expander(f"{name} ({len(reqs)})"):
                for r in reqs:
                    c_txt, c_btn = st.columns([4, 1])
                    c_txt.caption(f"{r['Inicio'].strftime('%d/%m')} - {r['Fin'].strftime('%d/%m')}")
                    if c_btn.button("🗑️", key=f"del_{r['idx']}"):
                        st.session_state.requests.pop(r['idx'])
                        st.rerun()
    else:
        st.info("Sin solicitudes.")

    if st.button("🗑️ Borrar TODO", type="secondary"):
        st.session_state.requests = []
        st.rerun()

st.divider()
if st.button("🚀 Generar Excel Final", type="primary", use_container_width=True):
    if not st.session_state.requests:
        st.error("Faltan solicitudes.")
    else:
        final_sch, errs, counters, fill_log, adjustments_log = validate_and_generate(
            edited_df, st.session_state.requests, year_val, st.session_state.nights
        )
        if errs:
            st.error("❌ Conflictos:")
            for e in errs: st.write(f"- {e}")
        else:
            st.success("✅ Éxito")
            excel_data = create_excel(
                final_sch, edited_df, year_val, st.session_state.requests, fill_log, counters, st.session_state.nights, adjustments_log
            )
            st.download_button("📥 Descargar", excel_data, f"Cuadrante_V5.3_{year_val}.xlsx")
