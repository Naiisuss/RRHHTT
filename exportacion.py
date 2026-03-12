# exportacion.py
import pandas as pd
import io

def generar_excel_con_semaforo(df_resumen, df_detalle):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. CREAMOS LAS HOJAS
        df_resumen.to_excel(writer, sheet_name='Resumen Gerencial', index=False)
        df_detalle.to_excel(writer, sheet_name='Registros Detallados', index=False)
        
        workbook  = writer.book
        ws_resumen = writer.sheets['Resumen Gerencial']
        ws_detalle = writer.sheets['Registros Detallados']
        
        # Agregamos la hoja de KPIs al principio
        ws_kpi = workbook.add_worksheet('Dashboard KPIs')
        workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop(workbook.worksheets_objs.index(ws_kpi)))

        # -----------------------------------------------------------------
        # COLORES CORPORATIVOS (CRF) Y FORMATOS "PREMIUM"
        # -----------------------------------------------------------------
        crf_azul = '#004b87'  # Azul corporativo
        crf_rojo = '#df0b25'  # Rojo corporativo
        crf_gris = '#f4f4f4'  # Gris claro para fondos limpios
        
        # Formatos para el Dashboard
        formato_titulo = workbook.add_format({'bold': True, 'font_size': 22, 'font_color': 'white', 'bg_color': crf_azul, 'align': 'center', 'valign': 'vcenter'})
        formato_subtitulo = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': 'white', 'bg_color': crf_rojo, 'align': 'center', 'valign': 'vcenter'})
        
        # Formatos para las tablas
        formato_header_tabla = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': crf_azul, 
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        formato_celda_centro = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_celda_izq = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        
        # Semáforo
        formato_verde = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_amarillo = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_rojo = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'valign': 'vcenter'})

        # -----------------------------------------------------------------
        # PESTAÑA 2: RESUMEN GERENCIAL (Súper espaciado)
        # -----------------------------------------------------------------
        ws_resumen.set_row(0, 35) # Fila de títulos más alta para que respire
        for col_num, value in enumerate(df_resumen.columns):
            ws_resumen.write(0, col_num, value, formato_header_tabla)
            
        # Ajustamos anchos y alineaciones
        ws_resumen.set_column('A:B', 14, formato_celda_centro)
        ws_resumen.set_column('C:C', 35, formato_celda_izq) 
        ws_resumen.set_column('D:D', 18, formato_celda_centro)
        ws_resumen.set_column('E:F', 15, formato_celda_centro)
        ws_resumen.set_column('G:I', 22, formato_celda_centro)
        ws_resumen.set_column('J:J', None, None, {'hidden': True}) 
        
        ws_resumen.autofilter(f'A1:I{len(df_resumen)}')
        ws_resumen.freeze_panes(1, 0)
        
        rango_semaforo = f'I2:I{len(df_resumen) + 1}'
        ws_resumen.conditional_format(rango_semaforo, {'type': 'formula', 'criteria': '=$J2>30', 'format': formato_rojo})
        ws_resumen.conditional_format(rango_semaforo, {'type': 'formula', 'criteria': '=AND($J2>=10, $J2<=30)', 'format': formato_amarillo})
        ws_resumen.conditional_format(rango_semaforo, {'type': 'formula', 'criteria': '=$J2<10', 'format': formato_verde})

        # -----------------------------------------------------------------
        # PESTAÑA 3: REGISTROS DETALLADOS (Para que no desentone)
        # -----------------------------------------------------------------
        ws_detalle.set_row(0, 30)
        for col_num, value in enumerate(df_detalle.columns):
            ws_detalle.write(0, col_num, value, formato_header_tabla)
        ws_detalle.set_column('A:Z', 18, formato_celda_centro)
        ws_detalle.freeze_panes(1, 0)

        # -----------------------------------------------------------------
        # PESTAÑA 1: DASHBOARD KPIs (Nivel Multinacional)
        # -----------------------------------------------------------------
        ws_kpi.hide_gridlines(2)
        ws_kpi.set_column('A:A', 3) # Margen izquierdo
        ws_kpi.set_column('B:D', 22) # Columnas de KPI
        ws_kpi.set_column('E:L', 15) # Columnas de Gráficos
        
        # Banner Corporativo Superior
        ws_kpi.set_row(1, 40) # Fila gigante
        ws_kpi.merge_range('B2:K2', 'CONTROL DE TIEMPOS Y ASISTENCIA', formato_titulo)
        ws_kpi.set_row(2, 20)
        ws_kpi.merge_range('B3:K3', 'REPORTE GERENCIAL - RECURSOS HUMANOS', formato_subtitulo)
        
        # Variables matemáticas
        total_empleados = len(df_resumen)
        criticos = len(df_resumen[df_resumen['Total_Num'] > 30])
        moderados = len(df_resumen[(df_resumen['Total_Num'] >= 10) & (df_resumen['Total_Num'] <= 30)])
        optimos = len(df_resumen[df_resumen['Total_Num'] < 10])
        tiempo_promedio = int(df_resumen['Total_Num'].mean()) if total_empleados > 0 else 0
        
        # Cajas de KPIs
        f_kpi_tit = workbook.add_format({'bold': True, 'font_size': 10, 'font_color': crf_azul, 'bg_color': crf_gris, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        f_kpi_val = workbook.add_format({'bold': True, 'font_size': 28, 'font_color': crf_azul, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        f_kpi_val_rojo = workbook.add_format({'bold': True, 'font_size': 28, 'font_color': crf_rojo, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        
        ws_kpi.set_row(4, 25) # Altura títulos KPI
        ws_kpi.write('B5', 'TOTAL EVALUADOS', f_kpi_tit)
        ws_kpi.write('C5', 'TIEMPO PROM. (MIN)', f_kpi_tit)
        ws_kpi.write('D5', 'CASOS CRÍTICOS', f_kpi_tit)
        
        ws_kpi.set_row(5, 55) # Altura números gigantes
        ws_kpi.write('B6', total_empleados, f_kpi_val)
        ws_kpi.write('C6', tiempo_promedio, f_kpi_val)
        ws_kpi.write('D6', criticos, f_kpi_val_rojo)
        
        # --- TABLAS OCULTAS ---
        ws_kpi.write('N3', 'Estado')
        ws_kpi.write('O3', 'Cantidad')
        ws_kpi.write('N4', 'Óptimo (<10m)')
        ws_kpi.write('O4', optimos)
        ws_kpi.write('N5', 'Alerta (10-30m)')
        ws_kpi.write('O5', moderados)
        ws_kpi.write('N6', 'Crítico (>30m)')
        ws_kpi.write('O6', criticos)
        
        # Gráfico 1: Dona (Mucho más elegante que la torta)
        chart_pie = workbook.add_chart({'type': 'doughnut'})
        chart_pie.add_series({
            'name': 'Distribución',
            'categories': "='Dashboard KPIs'!$N$4:$N$6",
            'values':     "='Dashboard KPIs'!$O$4:$O$6",
            'points': [{'fill': {'color': '#00b050'}}, {'fill': {'color': '#ffc000'}}, {'fill': {'color': crf_rojo}}],
            'data_labels': {'percentage': True, 'font': {'size': 11, 'bold': True, 'color': 'white'}}
        })
        chart_pie.set_title({'name': 'Estado del Personal', 'name_font': {'size': 13, 'color': crf_azul, 'bold': True}})
        chart_pie.set_size({'width': 350, 'height': 280})
        chart_pie.set_chartarea({'border': {'none': True}}) # Le quitamos el marco al gráfico para que se funda con la hoja
        ws_kpi.insert_chart('B8', chart_pie)

        # Gráfico 2: Barras Top 10
        top_10 = df_resumen.sort_values('Total_Num', ascending=False).head(10)
        row_oculta = 10
        for idx, data in top_10.iterrows():
            ws_kpi.write(row_oculta, 13, data['Nombre'])  
            ws_kpi.write(row_oculta, 14, data['Total_Num']) 
            row_oculta += 1
            
        if len(top_10) > 0:
            chart_bar = workbook.add_chart({'type': 'bar'})
            chart_bar.add_series({
                'name': 'Minutos Perdidos',
                'categories': f"='Dashboard KPIs'!$N$11:$N${row_oculta}",
                'values':     f"='Dashboard KPIs'!$O$11:$O${row_oculta}",
                'fill': {'color': crf_rojo},
                'data_labels': {'value': True, 'font': {'bold': True}} # Muestra el número exacto al final de la barra
            })
            chart_bar.set_title({'name': 'TOP 10: Mayor Tiempo Muerto', 'name_font': {'size': 13, 'color': crf_azul, 'bold': True}})
            chart_bar.set_y_axis({'reverse': True, 'major_gridlines': {'visible': False}}) 
            chart_bar.set_x_axis({'visible': False}) # Ocultamos los números de abajo para que quede más limpio
            chart_bar.set_legend({'none': True}) 
            chart_bar.set_size({'width': 550, 'height': 350})
            chart_bar.set_chartarea({'border': {'none': True}})
            ws_kpi.insert_chart('F8', chart_bar)

    procesado = output.getvalue()
    return procesado