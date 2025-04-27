import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
import io
import base64

# Importar funciones directamente del archivo route_planner.py
from route_planner import (read_excel_data, generate_routes, 
                           create_excel_report, print_summary, 
                           MAX_WEEKS, Task, Operario, RouteDayWeek)import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Planificador de Rutas - Eix Ambiental",
    page_icon="🗺️",
    layout="wide"
)

# Título y descripción
st.title("Planificador de Rutas - Eix Ambiental")
st.markdown("""
Esta aplicación genera una planificación optimizada de rutas para operarios de Eix Ambiental.
Las tareas se distribuyen de lunes a jueves, hasta un máximo de 4 semanas.
""")

# Barra lateral para la configuración
with st.sidebar:
    st.header("Configuración")
    
    # Carga de archivo
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx", "xls"])
    
    # Número de operarios
    num_operarios = st.radio("Número de operarios", [1, 2, 3], horizontal=True)
    
    # Información
    st.info(f"Las tareas se planificarán de lunes a jueves, hasta un máximo de {MAX_WEEKS} semanas.")
    
    # Botón para generar
    generate_button = st.button("Generar Planificación", type="primary", disabled=not uploaded_file)

# Área principal para resultados
if uploaded_file and generate_button:
    with st.spinner("Procesando el archivo Excel..."):
        # Leer datos
        df = pd.read_excel(uploaded_file, header=None)
        tasks = read_excel_data(df)
        
        if tasks:
            # Mostrar información de tareas cargadas
            st.success(f"Se han cargado {len(tasks)} tareas válidas")
            
            # Generar rutas
            with st.spinner(f"Generando planificación para {num_operarios} operarios..."):
                operarios = generate_routes(tasks, num_operarios)
            
            # Mostrar resumen
            st.subheader("Resumen de Planificación")
            
            # Crear columnas para mostrar estadísticas
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Total de tareas", len(tasks))
                st.metric("Poblaciones", len(set(task.poblacion for task in tasks)))
                
            with col2:
                total_minutes = sum(task.duracion for task in tasks)
                hours = total_minutes // 60
                minutes = total_minutes % 60
                st.metric("Duración total", f"{hours}h {minutes}min")
                st.metric("Días de trabajo", "Lunes a Jueves")
            
            # Crear y descargar Excel
            with st.spinner("Generando informe Excel..."):
                # Crear Excel en memoria
                output = io.BytesIO()
                success = create_excel_report(operarios, output)
                
                if success:
                    # Botón para descargar el Excel
                    output.seek(0)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    excel_filename = f"Planificacion_Rutas_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="📥 Descargar Planificación Excel",
                        data=output,
                        file_name=excel_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            
            # Mostrar distribución por operario
            st.subheader("Distribución por Operario")
            
            for operario in operarios:
                with st.expander(f"Operario {operario.operario_id}", expanded=True):
                    # Crear una tabla para este operario
                    op_data = []
                    total_tasks = 0
                    
                    # Obtener datos por semana
                    for week in range(1, MAX_WEEKS + 1):
                        week_tasks = sum(len(day.tasks) for day in operario.weeks[week].values())
                        week_time = sum(day.total_time for day in operario.weeks[week].values())
                        
                        if week_tasks > 0:
                            hours = week_time // 60
                            minutes = week_time % 60
                            op_data.append({
                                "Semana": f"Semana {week}",
                                "Tareas": week_tasks,
                                "Tiempo Total": f"{hours}h {minutes}min"
                            })
                            total_tasks += week_tasks
                    
                    # Añadir fila de total
                    if op_data:
                        op_df = pd.DataFrame(op_data)
                        st.table(op_df)
                        
                        # Detalle por día
                        st.write("Detalle por día:")
                        for week in range(1, MAX_WEEKS + 1):
                            for day_name in ["Lunes", "Martes", "Miércoles", "Jueves"]:
                                day_route = operario.weeks[week][day_name]
                                if day_route.tasks:
                                    with st.expander(f"{day_name} (Semana {week} - {day_route.date_str})"):
                                        day_data = []
                                        for task in day_route.tasks:
                                            day_data.append({
                                                "Hora": task.start_time.strftime("%H:%M"),
                                                "Cliente": task.nombre_cliente,
                                                "Población": task.poblacion,
                                                "Tarea": task.observaciones,
                                                "Duración": f"{task.duracion} min"
                                            })
                                        if day_data:
                                            st.table(pd.DataFrame(day_data))
                    else:
                        st.write("No hay tareas asignadas para este operario.")
            
            # Verificar si quedaron tareas sin asignar
            unassigned = [task for task in tasks if not task.assigned]
            if unassigned:
                st.warning(f"{len(unassigned)} tareas no pudieron ser asignadas")
                with st.expander("Ver tareas no asignadas"):
                    unassigned_data = []
                    for task in unassigned:
                        unassigned_data.append({
                            "Cliente": task.nombre_cliente,
                            "Población": task.poblacion,
                            "Dirección": task.direccion,
                            "Tarea": task.observaciones,
                            "Duración": f"{task.duracion} min"
                        })
                    st.table(pd.DataFrame(unassigned_data))
        else:
            st.error("No se pudieron cargar tareas del archivo Excel. Verifique el formato.")

# Footer
st.markdown("---")
st.caption("Planificador de Rutas - Eix Ambiental © 2025")
