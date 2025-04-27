import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
import io

# Constantes globales
ORIGIN_LOCATION = "Vic"  # Ubicación de la empresa
WORK_HOURS = 8  # Horas por jornada
WORK_DAYS = ["Lunes", "Martes", "Miércoles", "Jueves"]  # Días laborables
MAX_WEEKS = 4  # Máximo número de semanas para planificar
LUNCH_DURATION = 30  # Minutos para comer
START_HOUR = 8  # Hora de inicio de la jornada
START_MINUTE = 0  # Minuto de inicio de la jornada

class Task:
    """Clase para representar una tarea con todos sus atributos."""
    
    def __init__(self, row=None):
        """Inicializa una tarea a partir de una fila del Excel."""
        if row is not None:
            # Mapeo de las columnas del Excel
            self.mantenimiento = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            self.cod_cliente = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ""
            self.nombre_cliente = str(row.iloc[3]) if not pd.isna(row.iloc[3]) else ""
            self.direccion = str(row.iloc[4]) if not pd.isna(row.iloc[4]) else ""
            self.alias = str(row.iloc[5]) if not pd.isna(row.iloc[5]) else ""
            self.poblacion = str(row.iloc[6]) if not pd.isna(row.iloc[6]) else ""
            self.observaciones = str(row.iloc[11]) if not pd.isna(row.iloc[11]) else ""
        else:
            self.mantenimiento = ""
            self.cod_cliente = ""
            self.nombre_cliente = ""
            self.direccion = ""
            self.alias = ""
            self.poblacion = ""
            self.observaciones = ""
        
        # Variables para la planificación
        self.duracion = self.calcular_duracion() if hasattr(self, 'observaciones') else 0
        self.start_time = None
        self.travel_time = 0
        self.assigned = False  # Flag para saber si la tarea ya ha sido asignada
        
    def calcular_duracion(self):
        """Calcula la duración de la tarea basada en la descripción."""
        descripcion = self.observaciones.lower()
        duracion = 0
        
        # Contar legios (muestras de legionela)
        legio_match = re.search(r'(\d+)\s*legio', descripcion)
        if legio_match:
            num_legios = int(legio_match.group(1))
        elif 'legio' in descripcion:
            num_legios = 1
        else:
            num_legios = 0
        
        # Aplicar duración según número de legios
        if num_legios == 1:
            duracion = 45  # 45 minutos
        elif 2 <= num_legios <= 3:
            duracion = 60  # 1 hora
        elif 4 <= num_legios <= 5:
            duracion = 90  # 1.5 horas
        elif 6 <= num_legios <= 7:
            duracion = 120  # 2 horas
        elif 7 <= num_legios <= 9:
            duracion = 150  # 2.5 horas
        elif 9 <= num_legios <= 11:
            duracion = 180  # 3 horas
        
        # Añadir tiempo si incluye revisión
        if 'revisió' in descripcion or 'revisio' in descripcion:
            duracion += 45  # +45 minutos
        
        return duracion
    
    def is_valid(self):
        """Verifica si la tarea tiene datos válidos."""
        return bool(self.nombre_cliente and self.poblacion)
    
    def __str__(self):
        """Representación en texto de la tarea."""
        return f"{self.nombre_cliente} - {self.poblacion} - {self.observaciones} ({self.duracion} min)"


class RouteDayWeek:
    """Clase para representar un día de ruta de un operario en una semana específica."""
    
    def __init__(self, day_name, week_number):
        """Inicializa un día de ruta con su semana."""
        self.day_name = day_name
        self.week_number = week_number  # 1, 2, 3 o 4
        self.tasks = []
        self.total_time = 0
        self.start_location = ORIGIN_LOCATION
        self.end_location = ORIGIN_LOCATION
        self.return_travel_time = 0
        
        # Calcular la fecha real basada en la semana
        now = datetime.now()
        first_day_of_month = datetime(now.year, now.month, 1)
        # Encontrar el primer lunes del mes
        while first_day_of_month.weekday() != 0:  # 0 = Lunes
            first_day_of_month += timedelta(days=1)
        
        # Calcular el día correspondiente
        day_offset = WORK_DAYS.index(day_name)
        week_offset = (week_number - 1) * 7
        self.date = first_day_of_month + timedelta(days=day_offset + week_offset)
        
        # Formatear la fecha
        self.date_str = self.date.strftime("%d/%m/%Y")
    
    def add_task(self, task, travel_time):
        """Añade una tarea al día y actualiza los tiempos."""
        # Calcular la hora de inicio de la tarea
        start_time = self.calculate_start_time(travel_time)
        
        # Clonar la tarea y añadir datos de tiempo
        task_copy = Task()
        task_copy.__dict__.update(task.__dict__)
        task_copy.start_time = start_time
        task_copy.travel_time = travel_time
        
        # Añadir a la lista de tareas
        self.tasks.append(task_copy)
        
        # Actualizar tiempos y ubicación
        self.total_time += task_copy.duracion + travel_time
        self.end_location = task_copy.poblacion
        
        return task_copy
    
    def finalize_day(self):
        """Finaliza el día añadiendo el tiempo de viaje de regreso."""
        if self.tasks:
            # Añadir tiempo de viaje de regreso a Vic
            return_travel = estimate_travel_time(self.end_location, ORIGIN_LOCATION)
            self.return_travel_time = return_travel
            self.total_time += return_travel
            return True
        return False
    
    def calculate_start_time(self, travel_time):
        """Calcula la hora de inicio para una tarea."""
        if not self.tasks:
            # Primera tarea del día
            minutes = travel_time
            hours = START_HOUR
            mins = START_MINUTE
        else:
            last_task = self.tasks[-1]
            last_end_time = last_task.start_time + timedelta(minutes=last_task.duracion)
            current_start_time = last_end_time + timedelta(minutes=travel_time)
            
            # Comprobar si hay que añadir pausa para comer
            lunch_start = datetime.combine(datetime.today(), datetime.min.time()) + timedelta(hours=13)
            
            if (last_end_time.hour < 13 and current_start_time.hour >= 13 and 
                current_start_time.time() > lunch_start.time()):
                current_start_time += timedelta(minutes=LUNCH_DURATION)
            
            return current_start_time
        
        # Para la primera tarea
        start_datetime = datetime.combine(datetime.today(), datetime.min.time())
        start_datetime += timedelta(hours=hours, minutes=mins + minutes)
        return start_datetime
    
    def has_capacity_for(self, task_duration, travel_time):
        """Verifica si hay capacidad para añadir una tarea más."""
        new_total_time = self.total_time + task_duration + travel_time
        return new_total_time <= (WORK_HOURS * 60) - LUNCH_DURATION
    
    def get_full_day_name(self):
        """Devuelve el nombre completo del día incluyendo la semana."""
        return f"{self.day_name} (Semana {self.week_number} - {self.date_str})"


class Operario:
    """Clase para representar un operario con sus rutas asignadas en múltiples semanas."""
    
    def __init__(self, operario_id):
        """Inicializa un operario con ID y días de trabajo para múltiples semanas."""
        self.operario_id = operario_id
        self.weeks = {}
        
        # Inicializar las semanas (hasta 4)
        for week in range(1, MAX_WEEKS + 1):
            self.weeks[week] = {
                day: RouteDayWeek(day, week) for day in WORK_DAYS
            }
    
    def get_route_day(self, day_name, week_number):
        """Obtiene un día de ruta específico en una semana específica."""
        return self.weeks[week_number][day_name]


def estimate_travel_time(origin, destination):
    """Estima el tiempo de viaje entre dos ubicaciones (simplificado)."""
    # En una versión más avanzada, esto podría usar una API de mapas
    if origin == destination:
        return 5  # 5 minutos entre ubicaciones cercanas
    return 30  # 30 minutos entre ubicaciones diferentes


def format_time(datetime_obj):
    """Formatea un objeto datetime a string HH:MM."""
    return datetime_obj.strftime("%H:%M")


def format_minutes(minutes):
    """Formatea minutos como horas y minutos en texto."""
    hours = minutes // 60
    mins = minutes % 60
    if hours > 0:
        return f"{hours}h {mins}min" if mins > 0 else f"{hours}h"
    return f"{mins}min"


def read_excel_data(df):
    """Lee el DataFrame y extrae las tareas."""
    try:
        # Convertir filas a tareas
        tasks = []
        for _, row in df.iterrows():
            task = Task(row)
            if task.is_valid():
                tasks.append(task)
        
        return tasks
    
    except Exception as e:
        print(f"Error al procesar los datos del Excel: {e}")
        return []


def generate_routes(tasks, num_operarios):
    """Genera rutas optimizadas para los operarios en múltiples semanas."""
    # Lista de operarios
    operarios = [Operario(i+1) for i in range(num_operarios)]
    
    # Agrupar tareas por población
    tasks_by_location = {}
    for task in tasks:
        if task.poblacion not in tasks_by_location:
            tasks_by_location[task.poblacion] = []
        tasks_by_location[task.poblacion].append(task)
    
    # Ordenar poblaciones (podría mejorarse con distancias reales)
    locations = sorted(tasks_by_location.keys())
    
    # Variables para distribución
    current_operario = 0
    current_week = 1
    current_day_index = 0
    
    # Para cada población
    for location in locations:
        location_tasks = tasks_by_location[location]
        
        # Ordenar tareas por duración (primero las más largas)
        location_tasks.sort(key=lambda x: x.duracion, reverse=True)
        
        # Asignar cada tarea
        for task in location_tasks:
            # Saltear tareas ya asignadas
            if task.assigned:
                continue
            
            # Intentar asignar la tarea
            assigned = False
            
            # Buscar un día, semana y operario que tenga espacio
            for attempt in range(num_operarios * len(WORK_DAYS) * MAX_WEEKS):
                op_index = current_operario % num_operarios
                week_num = current_week
                day_index = current_day_index % len(WORK_DAYS)
                day_name = WORK_DAYS[day_index]
                
                # Obtener el día de ruta para esta semana
                route_day = operarios[op_index].get_route_day(day_name, week_num)
                
                # Estimar tiempo de viaje
                origin = route_day.end_location
                travel_time = estimate_travel_time(origin, task.poblacion)
                
                # Comprobar si cabe en la jornada
                if route_day.has_capacity_for(task.duracion, travel_time):
                    # Asignar la tarea
                    route_day.add_task(task, travel_time)
                    task.assigned = True
                    assigned = True
                    break
                
                # Probar con el siguiente día/operario/semana
                current_operario = (current_operario + 1) % num_operarios
                if current_operario == 0:
                    current_day_index = (current_day_index + 1) % len(WORK_DAYS)
                    if current_day_index == 0:
                        current_week = (current_week % MAX_WEEKS) + 1
            
            if not assigned:
                print(f"ADVERTENCIA: No se pudo asignar la tarea: {task}")
            
            # Mover al siguiente operario para la próxima tarea
            current_operario = (current_operario + 1) % num_operarios
            if current_operario == 0:
                current_day_index = (current_day_index + 1) % len(WORK_DAYS)
                if current_day_index == 0:
                    current_week = (current_week % MAX_WEEKS) + 1
    
    # Finalizar rutas (añadir viaje de vuelta)
    for operario in operarios:
        for week in range(1, MAX_WEEKS + 1):
            for day_name in WORK_DAYS:
                operario.weeks[week][day_name].finalize_day()
    
    return operarios


def create_excel_report(operarios, output_buffer):
    """Crea un informe Excel con las rutas generadas."""
    try:
        # Crear un ExcelWriter
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # Para cada operario, crear una hoja
            for operario in operarios:
                # Crear DataFrame para este operario
                data = []
                
                # Para cada semana y día
                for week in range(1, MAX_WEEKS + 1):
                    week_has_tasks = False
                    
                    # Verificar si esta semana tiene tareas
                    for day_name in WORK_DAYS:
                        if operario.weeks[week][day_name].tasks:
                            week_has_tasks = True
                            break
                    
                    if not week_has_tasks:
                        continue  # Saltar semanas sin tareas
                    
                    # Añadir encabezado de semana
                    data.append({
                        'Semana': f"SEMANA {week}",
                        'Dia': "",
                        'Fecha': "",
                        'Hora': "",
                        'Cliente': "",
                        'Poblacion': "",
                        'Direccion': "",
                        'Tarea': "",
                        'Duracion': "",
                        'Tiempo_Viaje': ""
                    })
                    
                    # Para cada día de esta semana
                    for day_name in WORK_DAYS:
                        day_route = operario.weeks[week][day_name]
                        
                        if day_route.tasks:
                            # Añadir encabezado de día
                            data.append({
                                'Semana': "",
                                'Dia': day_name,
                                'Fecha': day_route.date_str,
                                'Hora': "",
                                'Cliente': "",
                                'Poblacion': "",
                                'Direccion': "",
                                'Tarea': "",
                                'Duracion': "",
                                'Tiempo_Viaje': ""
                            })
                            
                            # Salida desde Vic
                            data.append({
                                'Semana': "",
                                'Dia': "",
                                'Fecha': "",
                                'Hora': '08:00',
                                'Cliente': 'Eix Ambiental',
                                'Poblacion': 'Vic',
                                'Direccion': '-',
                                'Tarea': 'Sortida',
                                'Duracion': '-',
                                'Tiempo_Viaje': '-'
                            })
                            
                            # Tareas del día
                            last_end_time = None
                            for task in day_route.tasks:
                                task_start_time = task.start_time
                                task_end_time = task_start_time + timedelta(minutes=task.duracion)
                                
                                # Comprobar si hay que insertar pausa para comer
                                if last_end_time:
                                    lunch_time = datetime.combine(datetime.today(), datetime.min.time()) + timedelta(hours=13)
                                    
                                    if (last_end_time.hour < 13 and task_start_time.hour >= 13 and 
                                        task_start_time.time() > lunch_time.time()):
                                        # Insertar pausa para comer
                                        data.append({
                                            'Semana': "",
                                            'Dia': "",
                                            'Fecha': "",
                                            'Hora': '13:00',
                                            'Cliente': '-',
                                            'Poblacion': '-',
                                            'Direccion': '-',
                                            'Tarea': 'Pausa per dinar',
                                            'Duracion': '30 min',
                                            'Tiempo_Viaje': '-'
                                        })
                                
                                # Añadir la tarea
                                data.append({
                                    'Semana': "",
                                    'Dia': "",
                                    'Fecha': "",
                                    'Hora': format_time(task_start_time),
                                    'Cliente': task.nombre_cliente,
                                    'Poblacion': task.poblacion,
                                    'Direccion': task.direccion,
                                    'Tarea': task.observaciones,
                                    'Duracion': format_minutes(task.duracion),
                                    'Tiempo_Viaje': format_minutes(task.travel_time)
                                })
                                
                                last_end_time = task_end_time
                            
                            # Regreso a Vic
                            if last_end_time:
                                data.append({
                                    'Semana': "",
                                    'Dia': "",
                                    'Fecha': "",
                                    'Hora': format_time(last_end_time),
                                    'Cliente': 'Eix Ambiental',
                                    'Poblacion': 'Vic',
                                    'Direccion': '-',
                                    'Tarea': 'Tornada',
                                    'Duracion': '-',
                                    'Tiempo_Viaje': format_minutes(day_route.return_travel_time)
                                })
                
                # Crear DataFrame y escribir a Excel
                if data:
                    df = pd.DataFrame(data)
                    df.to_excel(writer, sheet_name=f'Operario {operario.operario_id}', index=False)
        
        return True
    
    except Exception as e:
        print(f"Error al crear el informe Excel: {e}")
        return False


def print_summary(tasks, operarios):
    """Imprime un resumen de la planificación."""
    summary = []
    summary.append("\n===== RESUMEN DE PLANIFICACIÓN =====")
    summary.append(f"Total de tareas: {len(tasks)}")
    summary.append(f"Poblaciones: {len(set(task.poblacion for task in tasks))}")
    summary.append(f"Días de trabajo: {', '.join(WORK_DAYS)} (hasta {MAX_WEEKS} semanas)")
    
    total_minutes = sum(task.duracion for task in tasks)
    summary.append(f"Duración total de tareas: {format_minutes(total_minutes)} (sin contar desplazamientos)")
    
    summary.append("\n--- Distribución por operario ---")
    for operario in operarios:
        total_tasks = 0
        total_time = 0
        
        # Contar por semanas
        for week in range(1, MAX_WEEKS + 1):
            week_tasks = sum(len(day.tasks) for day in operario.weeks[week].values())
            week_time = sum(day.total_time for day in operario.weeks[week].values())
            
            if week_tasks > 0:
                summary.append(f"  Operario {operario.operario_id} - Semana {week}: {week_tasks} tareas, {format_minutes(week_time)}")
                total_tasks += week_tasks
                total_time += week_time
        
        summary.append(f"  Operario {operario.operario_id} - TOTAL: {total_tasks} tareas, {format_minutes(total_time)}")
    
    # Verificar si quedaron tareas sin asignar
    unassigned = [task for task in tasks if not task.assigned]
    if unassigned:
        summary.append(f"\nAdvertencia: {len(unassigned)} tareas no pudieron ser asignadas")
    
    return "\n".join(summary)
