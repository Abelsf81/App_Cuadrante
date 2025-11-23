import pandas as pd
import datetime
import random
import calendar
from datetime import timedelta

# --- 1. CONFIGURACIÓN ---
YEAR = 2026
FILENAME = "Carga_Vacaciones_PERFECTAS_13.xlsx"

# Periodos Nocturnos (Aproximados para evitar conflictos críticos en la generación)
# Formato: (Mes, Dia Fin) - Solo ponemos las fechas de FINAL de periodo que son las críticas
CRITICAL_NIGHT_ENDS = [
    datetime.date(YEAR, 1, 12), datetime.date(YEAR, 1, 27),
    datetime.date(YEAR, 2, 26), datetime.date(YEAR, 3, 28),
    datetime.date(YEAR, 4, 9),  datetime.date(YEAR, 4, 24),
    datetime.date(YEAR, 5, 12), datetime.date(YEAR, 5, 27),
    datetime.date(YEAR, 6, 14), datetime.date(YEAR, 6, 29),
    datetime.date(YEAR, 7, 29), datetime.date(YEAR, 8, 4),
    datetime.date(YEAR, 8, 28), datetime.date(YEAR, 9, 3),
    datetime.date(YEAR, 12, 31)
]

ROSTER = [
    # Jerarquía ordenada para el Draft
    {"ID": "Jefe A", "Nombre": "Jefe A", "Turno": "A", "Rol": "Jefe"},
    {"ID": "Jefe B", "Nombre": "Jefe B", "Turno": "B", "Rol": "Jefe"},
    {"ID": "Jefe C", "Nombre": "Jefe C", "Turno": "C", "Rol": "Jefe"},
    
    {"ID": "Subjefe A", "Nombre": "Subjefe A", "Turno": "A", "Rol": "Subjefe"},
    {"ID": "Subjefe B", "Nombre": "Subjefe B", "Turno": "B", "Rol": "Subjefe"},
    {"ID": "Subjefe C", "Nombre": "Subjefe C", "Turno": "C", "Rol": "Subjefe"},
    
    {"ID": "Cond A", "Nombre": "Cond A", "Turno": "A", "Rol": "Conductor"},
    {"ID": "Cond B", "Nombre": "Cond B", "Turno": "B", "Rol": "Conductor"},
    {"ID": "Cond C", "Nombre": "Cond C", "Turno": "C", "Rol": "Conductor"},
    
    {"ID": "Bombero A1", "Nombre": "Bombero A1", "Turno": "A", "Rol": "Bombero"},
    {"ID": "Bombero B1", "Nombre": "Bombero B1", "Turno": "B", "Rol": "Bombero"},
    {"ID": "Bombero C1", "Nombre": "Bombero C1", "Turno": "C", "Rol": "Bombero"},
    {"ID": "Bombero A2", "Nombre": "Bombero A2", "Turno": "A", "Rol": "Bombero"},
    {"ID": "Bombero B2", "Nombre": "Bombero B2", "Turno": "B", "Rol": "Bombero"},
    {"ID": "Bombero C2", "Nombre": "Bombero C2", "Turno": "C", "Rol": "Bombero"},
    {"ID": "Bombero A3", "Nombre": "Bombero A3", "Turno": "A", "Rol": "Bombero"},
    {"ID": "Bombero B3", "Nombre": "Bombero B3", "Turno": "B", "Rol": "Bombero"},
    {"ID": "Bombero C3", "Nombre": "Bombero C3", "Turno": "C", "Rol": "Bombero"},
]

# Mapa de ocupación global
occupation_map = {} # {dia_idx: [lista_personas]}

# --- 2. MOTOR LÓGICO ---

def generate_base_schedule():
    schedule = {'A': [], 'B': [], 'C': []}
    status = {'A': 0, 'B': 2, 'C': 1}
    days = 365 # 2026 no bisiesto
    for _ in range(days):
        for t in ['A', 'B', 'C']:
            schedule[t].append('T' if status[t] == 0 else 'L')
            status[t] = (status[t] + 1) % 3
    return schedule

BASE_SCH = generate_base_schedule()

def get_valid_blocks(turn, duration, required_credits):
    """
    Genera TODOS los bloques posibles en el año para un turno que cumplan:
    1. Duración exacta.
    2. Créditos (Guardias) exactos.
    3. No terminan en cambio de nocturna crítico.
    """
    valid_starts = []
    sch = BASE_SCH[turn]
    
    for d in range(0, 365 - duration):
        # Calcular créditos del bloque
        credits = 0
        for i in range(d, d + duration):
            if sch[i] == 'T': credits += 1
            
        # Check Nocturna (Final del bloque)
        end_date = datetime.date(YEAR, 1, 1) + timedelta(days=d + duration - 1)
        is_night_conflict = False
        # Si el último día del bloque es un día de transición de nocturna Y se trabaja
        if end_date in CRITICAL_NIGHT_ENDS:
             if sch[d + duration - 1] == 'T': is_night_conflict = True
        
        if credits == required_credits and not is_night_conflict:
            valid_starts.append(d)
            
    return valid_starts

def check_global_conflict(start_idx, duration, person):
    """Verifica si choca con las reglas globales de la unidad."""
    for i in range(start_idx, start_idx + duration):
        # Si el día no tiene ocupantes, libre
        if i not in occupation_map: continue
        
        occupants = occupation_map[i]
        
        # REGLA 1: MAX 2 PERSONAS
        if len(occupants) >= 2: return True
        
        for occ in occupants:
            # REGLA 2: MISMO TURNO (Estricta)
            if occ['Turno'] == person['Turno']: return True
            
            # REGLA 3: MISMA CATEGORÍA (Exc. Bomberos)
            if person['Rol'] != 'Bombero' and occ['Rol'] == person['Rol']: return True
            
    return False

def book_global(start_idx, duration, person):
    for i in range(start_idx, start_idx + duration):
        if i not in occupation_map: occupation_map[i] = []
        occupation_map[i].append(person)

# --- 3. EL ALGORITMO "CHEF" ---

final_rows = []
print("👨‍🍳 Cocinando menú de vacaciones perfecto (Objetivo: 13 Créditos)...")

# Definimos la "Dieta" de 13 Créditos:
# 2 Bloques de 4 créditos (10 días) -> Pata Negra
# 1 Bloque de 3 créditos (10 días)  -> Estándar
# 1 Bloque de 2 créditos (9 días)   -> Ligero
# Total: 4+4+3+2 = 13 Créditos.

RECIPE = [
    {"dur": 10, "cred": 4, "season": "summer"}, # Prioridad Verano
    {"dur": 10, "cred": 4, "season": "any"},
    {"dur": 10, "cred": 3, "season": "any"},
    {"dur": 9,  "cred": 2, "season": "any"}
]

for person in ROSTER:
    print(f"  > Asignando a: {person['Nombre']}...", end="")
    person_blocks = []
    
    # Obtener todas las opciones posibles para este turno
    options_4c = get_valid_blocks(person['Turno'], 10, 4)
    options_3c = get_valid_blocks(person['Turno'], 10, 3)
    options_2c = get_valid_blocks(person['Turno'], 9, 2)
    
    # Mezclar para aleatoriedad
    random.shuffle(options_4c)
    random.shuffle(options_3c)
    random.shuffle(options_2c)
    
    # Intentar llenar la receta
    my_slots = []
    
    # 1. Los dos bloques de 4 créditos (Pata Negra)
    count_4c = 0
    for start in options_4c:
        if count_4c >= 2: break
        # Verificar que no solape con los propios ya elegidos
        overlap_own = any(abs(start - s) < 11 for s in my_slots)
        if not overlap_own and not check_global_conflict(start, 10, person):
            my_slots.append((start, 10))
            book_global(start, 10, person)
            count_4c += 1
            
    # 2. El bloque de 3 créditos
    count_3c = 0
    for start in options_3c:
        if count_3c >= 1: break
        overlap_own = any(abs(start - s) < 11 for s in my_slots) # Margen de seguridad
        if not overlap_own and not check_global_conflict(start, 10, person):
            my_slots.append((start, 10))
            book_global(start, 10, person)
            count_3c += 1

    # 3. El bloque de 2 créditos (9 días)
    count_2c = 0
    for start in options_2c:
        if count_2c >= 1: break
        overlap_own = any(abs(start - s) < 11 for s in my_slots)
        if not overlap_own and not check_global_conflict(start, 9, person):
            my_slots.append((start, 9))
            book_global(start, 9, person)
            count_2c += 1
            
    # Formatear para Excel
    row = {"ID_Puesto": person['ID'], "Nombre": person['Nombre']}
    my_slots.sort(key=lambda x: x[0]) # Ordenar cronológicamente
    
    for i, (start, dur) in enumerate(my_slots):
        d_start = datetime.date(YEAR, 1, 1) + timedelta(days=start)
        d_end = d_start + timedelta(days=dur-1)
        row[f"Inicio {i+1}"] = d_start
        row[f"Fin {i+1}"] = d_end
        
    final_rows.append(row)
    print(f" Asignados {len(my_slots)}/4 periodos.")

# --- 4. GUARDAR ---
df = pd.DataFrame(final_rows)
# Ordenar columnas
cols = ["ID_Puesto", "Nombre", "Inicio 1", "Fin 1", "Inicio 2", "Fin 2", "Inicio 3", "Fin 3", "Inicio 4", "Fin 4"]
df = df.reindex(columns=cols)
df.to_excel(FILENAME, index=False)
print(f"\n✅ LISTO: {FILENAME}")
print("Sube este archivo a la App V10.2. Debería darte 13 créditos a casi todos.")
