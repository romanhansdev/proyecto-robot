from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
from datetime import datetime

# --- CONFIGURACIÓN ---
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
SERVICE_ACCOUNT_FILE = 'path/to/your/service-account-file.json'  # Cambiar por la ruta del archivo JSON
CALENDAR_ID = 'your-calendar-id@group.calendar.google.com'  # Cambiar por el ID del calendario
EXCEL_FILE = 'reporte_tickets.xlsx'

# --- FUNCIONES DEL BOT ---

def obtener_eventos():
    """Obtiene los eventos del calendario de Google."""
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('calendar', 'v3', credentials=credentials)

    now = datetime.utcnow().isoformat() + 'Z'  # Fecha y hora actual en formato RFC3339
    print('Obteniendo eventos desde:', now)

    events_result = service.events().list(
        calendarId=CALENDAR_ID, timeMin=now, singleEvents=True,
        orderBy='startTime').execute()
    events = events_result.get('items', [])

    return events

def procesar_eventos(events):
    """Procesa los eventos para extraer información relevante."""
    datos = []
    total_horas_por_actividad = {}  # Diccionario para acumular horas por actividad
    for event in events:
        # Extraer información básica
        titulo = event.get('summary', 'Sin título')
        inicio = event['start'].get('dateTime', event['start'].get('date'))
        fin = event['end'].get('dateTime', event['end'].get('date'))
        participantes = event.get('attendees', [])

        # Calcular duración
        inicio_dt = datetime.fromisoformat(inicio)
        fin_dt = datetime.fromisoformat(fin)
        duracion = (fin_dt - inicio_dt).total_seconds() / 3600  # Duración en horas

        # Acumular duración por actividad
        if titulo not in total_horas_por_actividad:
            total_horas_por_actividad[titulo] = 0
        total_horas_por_actividad[titulo] += duracion

        # Filtrar participantes con correos específicos
        # Puedes modificar el dominio aquí para cambiar los correos que se filtran
        participantes_filtrados = [
            p['email'] for p in participantes if '@kbeli.cl' in p.get('email', '')
        ]

        # Extraer número de ticket del título (si existe)
        numero_ticket = None
        if "TICKET-" in titulo:
            numero_ticket = titulo.split("TICKET-")[1].split()[0]

        # Agregar datos procesados
        datos.append({
            'Fecha': inicio_dt.date(),
            'Número de Ticket': numero_ticket,
            'Duración (horas)': duracion,
            'Participantes': ', '.join(participantes_filtrados),
            'Título': titulo
        })

    # Agregar totales por actividad al final del reporte
    for actividad, total_horas in total_horas_por_actividad.items():
        datos.append({
            'Fecha': 'TOTAL',
            'Número de Ticket': '',
            'Duración (horas)': total_horas,
            'Participantes': '',
            'Título': f'Total de horas para "{actividad}"'
        })

    return datos

def generar_excel(datos):
    """Genera un archivo Excel con los datos procesados."""
    # Crear un DataFrame con las columnas necesarias
    columnas = ['Fecha', 'Número de Ticket', 'Duración (horas)', 'Participantes', 'Título']
    df = pd.DataFrame(datos, columns=columnas)
    
    # Guardar el DataFrame en un archivo Excel
    df.to_excel(EXCEL_FILE, index=False, sheet_name="Reporte de Tickets")
    print(f"✅ Archivo Excel generado: {EXCEL_FILE}")

# --- EJECUCIÓN DEL BOT ---

if __name__ == "__main__":
    try:
        eventos = obtener_eventos()
        if not eventos:
            print("No se encontraron eventos en el calendario.")
        else:
            datos_procesados = procesar_eventos(eventos)
            generar_excel(datos_procesados)
    except Exception as e:
        print(f"❌ Error al ejecutar el bot: {e}")