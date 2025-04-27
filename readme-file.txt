# Planificador de Rutas - Eix Ambiental

Aplicación web para planificar rutas optimizadas para operarios de Eix Ambiental. La aplicación distribuye tareas entre operarios en jornadas de lunes a jueves, con planificación hasta 4 semanas.

## Características

- Procesamiento de archivos Excel con tareas mensuales
- Distribución optimizada por ubicación geográfica
- Planificación multi-semana (hasta 4 semanas)
- Cálculo automático de tiempos basado en tipo de tarea
- Exportación a Excel de la planificación

## Demo

La aplicación está desplegada en Streamlit Cloud: [Ver Demo](https://link-to-your-app.streamlit.app)

## Estructura de los Archivos Excel

La aplicación está diseñada para procesar archivos Excel con la siguiente estructura de columnas:
- Columna B: Mantenimiento
- Columna C: Código cliente
- Columna D: Nombre cliente
- Columna E: Dirección
- Columna F: Alias
- Columna G: Población
- Columna L: Observaciones/Tareas (donde se especifican los legios y revisiones)

## Uso Local

### Requisitos

- Python 3.7+
- Dependencias en `requirements.txt`

### Instalación

1. Clonar el repositorio:
```
git clone https://github.com/yourusername/planificador-rutas.git
cd planificador-rutas
```

2. Instalar dependencias:
```
pip install -r requirements.txt
```

3. Ejecutar la aplicación:
```
streamlit run app.py
```

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue para discutir los cambios importantes antes de enviar un pull request.

## Licencia

Este proyecto está licenciado bajo la licencia MIT.
