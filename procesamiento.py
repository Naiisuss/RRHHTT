# procesamiento.py
import pandas as pd

def formato_hrs_min(minutos_float):
    """Convierte los minutos decimales en un texto limpio de Horas y Minutos"""
    minutos = int(round(minutos_float))
    if minutos <= 0:
        return "0m"
    elif minutos >= 60:
        horas = minutos // 60
        mins = minutos % 60
        if mins == 0:
            return f"{horas}hs"
        else:
            return f"{horas}hs {mins}m"
    else:
        return f"{minutos}m"

def procesar_tiempos(df_registros, df_horarios):
    df = df_registros.copy()
    horarios = df_horarios.copy()
    
    # 1. Normalizamos columnas
    df.columns = df.columns.str.strip().str.upper()
    horarios.columns = horarios.columns.str.strip().str.upper()
    
    # 🚫 FIX CLAVE: ELIMINAR REGISTROS DEL COMEDOR
    # Si existe la columna TERMINAL, filtramos y borramos todo lo que diga "COMEDOR"
    if 'TERMINAL' in df.columns:
        # Nos quedamos SOLO con las filas que NO (~) contienen la palabra COMEDOR
        df = df[~df['TERMINAL'].astype(str).str.upper().str.contains('COMEDOR', na=False)]

    # Forzamos DNI a número para un cruce perfecto
    if 'DNI' in horarios.columns:
        horarios['DNI'] = pd.to_numeric(horarios['DNI'], errors='coerce').fillna(0).astype('int64')
    if 'DNI' in df.columns:
        df['DNI'] = pd.to_numeric(df['DNI'], errors='coerce').fillna(0).astype('int64')

    if 'FECHA Y HORA' in df.columns:
        df['Fecha_Hora'] = pd.to_datetime(df['FECHA Y HORA'], errors='coerce')
    elif 'FECHA' in df.columns and 'HORA' in df.columns:
        df['Fecha_Hora'] = pd.to_datetime(df['FECHA'].astype(str) + ' ' + df['HORA'].astype(str), errors='coerce')
    else:
        raise Exception("No se encontró la columna de Fecha y Hora en el archivo de SPEC.")

    df['Fecha_Dia'] = df['Fecha_Hora'].dt.date

    if 'CÓDIGO' in df.columns and 'LEGAJO' not in df.columns: df['LEGAJO'] = df['CÓDIGO']
    if 'APELLIDOS' in df.columns and 'NOMBRE' in df.columns: df['APELLIDO Y NOMBRE'] = df['APELLIDOS'].astype(str) + ', ' + df['NOMBRE'].astype(str)
    if 'ZONA' in df.columns and 'REGISTRO' not in df.columns: df['REGISTRO'] = df['ZONA']

    # Cruzamos los archivos
    if 'DNI' in df.columns and 'DNI' in horarios.columns:
        cols_horarios = ['DNI', 'TURNO', 'FINTURNO']
        df = pd.merge(df, horarios[cols_horarios], on='DNI', how='left')
    else:
        df['TURNO'] = 'Sin Turno'
        df['FINTURNO'] = pd.NaT

    # Limpiador de Turnos para evitar la "Ñ" mutante
    def limpiar_turno(t):
        t_str = str(t).upper()
        if 'MA' in t_str and 'NA' in t_str: 
            return 'Mañana'
        elif 'TARDE' in t_str: 
            return 'Tarde'
        elif 'NOCHE' in t_str: 
            return 'Noche'
        elif t_str in ['NAN', 'NAT', 'NONE', '']:
            return 'Sin Turno'
        return str(t).title() 
        
    df['TURNO'] = df['TURNO'].apply(limpiar_turno)

    # Ordenamos a los empleados cronológicamente
    df = df.sort_values(['LEGAJO', 'Fecha_Hora']).reset_index(drop=True)
    resumen = []
    
    for (legajo, fecha_dia), grupo in df.groupby(['LEGAJO', 'Fecha_Dia']):
        grupo = grupo.sort_values('Fecha_Hora')
        
        nombre = grupo['APELLIDO Y NOMBRE'].iloc[0] if 'APELLIDO Y NOMBRE' in grupo.columns else 'S/D'
        turno = grupo['TURNO'].iloc[0] if 'TURNO' in grupo.columns else 'Sin Turno'
        
        hora_inicio_real = grupo['Fecha_Hora'].iloc[0]
        hora_fin_real = grupo['Fecha_Hora'].iloc[-1]
        
        tiempo_muerto_intermedio = pd.Timedelta(seconds=0)
        salida_anticipada = pd.Timedelta(seconds=0)
        
        # A) Cálculo de Ocio Intermedio
        ultima_salida = None
        for index, row in grupo.iterrows():
            evento = str(row['REGISTRO']).upper() if 'REGISTRO' in grupo.columns else ""
            if 'SALIDA' in evento:
                ultima_salida = row['Fecha_Hora']
            elif 'INGRESO' in evento and ultima_salida is not None:
                tiempo_muerto_intermedio += (row['Fecha_Hora'] - ultima_salida)
                ultima_salida = None
                
        # B) Cálculo de Salida Anticipada
        fin_turno_teorico = grupo['FINTURNO'].iloc[0] if 'FINTURNO' in grupo.columns else None
        
        if pd.notna(fin_turno_teorico) and str(fin_turno_teorico).strip() not in ['', 'nan', 'NaT']:
            try:
                fin_str = str(fin_turno_teorico).strip()
                partes = fin_str.split(':')
                hora_td = int(partes[0])
                min_td = int(partes[1])
                
                fin_jornada_td = pd.Timedelta(hours=hora_td, minutes=min_td)
                fecha_base = hora_inicio_real.date()
                
                fecha_fin_teorica = pd.to_datetime(f"{fecha_base} 00:00:00") + fin_jornada_td
                
                if fecha_fin_teorica < hora_inicio_real:
                    fecha_fin_teorica += pd.Timedelta(days=1)

                if hora_fin_real < fecha_fin_teorica:
                    diff_segundos = (fecha_fin_teorica - hora_fin_real).total_seconds()
                    salida_anticipada = pd.Timedelta(seconds=diff_segundos)
            except Exception:
                pass 

        tiempo_muerto_total = tiempo_muerto_intermedio + salida_anticipada
        
        ocio_min = max(0, tiempo_muerto_intermedio.total_seconds() / 60)
        anticipada_min = max(0, salida_anticipada.total_seconds() / 60)
        total_min = max(0, tiempo_muerto_total.total_seconds() / 60)
        
        resumen.append({
            'Fecha': fecha_dia,
            'Legajo': legajo,
            'Nombre': nombre,
            'Turno Asignado': turno,
            'Primer Ingreso': hora_inicio_real.strftime('%H:%M:%S'),
            'Última Salida': hora_fin_real.strftime('%H:%M:%S'),
            'Tiempo Ocioso': formato_hrs_min(ocio_min),
            'Se retiró antes': formato_hrs_min(anticipada_min),
            'Tiempo Muerto TOTAL': formato_hrs_min(total_min),
            'Total_Num': int(round(total_min))
        })

    df_resumen = pd.DataFrame(resumen)
    return df_resumen, df