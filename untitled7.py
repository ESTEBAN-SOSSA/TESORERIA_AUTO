# Librerias

import time
from pynput.mouse import Button, Controller
import numpy as np
import re
import pyautogui
import pyperclip
from datetime import datetime
import pyodbc
import ctypes 
import pandas as pd
import smtplib
from email.message import EmailMessage

########################################################
## Exactraccion de infromación de la base datos de WM ##
########################################################

# Crear el DataFrame con los datos proporcionados
data = {
    'NIT': [
        '890319193', '890319193', '811019984', '800147520', '800147520', '800147520', '800147520',
        '860034944', '860034944', '860034944', '860034944', '860034944', '860034944', '860034944',
        '890311875', '890311875', '890311875', '890311875', '900479968', '900479968', '900479968', 
		'900879149', '900879149', '890921335', '890921335', '890921335', '890921335', '890319193', 
		'890319193', '860034944', '800147520', '800147520', '901094947', '901094947', '901094947', 
		'890319193', '860034944', '800147520'
    ],
    'FAC': [
        '03IE-00053405-00', '03IE-00054242-00', 'FEVP-00056516-00', 'FE-00059078-00', 'FE-00060437-00', 'FE-00060495-00', 'FE-00060656-00',
        'AJC-00003696-00', 'BT-00470736-00', 'BT-00470737-00', 'BT-00470701-00', 'BT-00472831-00', 'BT-00471814-00', 'BT-00472663-00',
		'FV30-48091088-00', 'FV30-48091147-00', 'FV30-48091330-00', 'FV30-48092134-00', '20200806-00000001-00', 'NIC-00042617-00', 'NIC-00043424-00', 
		'NIC-00047682-00', 'NIC-00049322-00', '-25022022-00', 'NIC-00036642-00', 'LC-00000818-00', 'LYFA-00004550-00', 'LYFA-00006810-00', 
		'LYFA-00010246-00', 'FAC-00NOEN-00', 'NO-ENCONTRADA', 'PRUEBA-2', 'NO-INF', 'TEST-00-00', '99-TYYY', '-00043958-00', '-00043957-00', '41845'
    ]
}

df_final = pd.DataFrame(data)

# Guardar el DataFrame filtrado en un archivo Excel
output_file = r"C:\Users\usuario\Documents\AUTOMATIZACION_PAGO_PROVEEDORES\Resultados\Todas_Facturas.xlsx"
df_final.to_excel(output_file, index=False)

###############################################
## Limpiar Excel base de facturas no validas ##
###############################################

# Ruta del archivo Excel
ruta_archivo = r"C:\Users\usuario\Documents\AUTOMATIZACION_PAGO_PROVEEDORES\Resultados\facturas_no_validadas.xlsx"

# Leer el archivo Excel para obtener los encabezados
df = pd.read_excel(ruta_archivo)

# Crear un DataFrame vacío con los mismos encabezados
df_limpio = pd.DataFrame(columns=df.columns)

# Sobrescribir el archivo original con el DataFrame vacío (solo encabezados)
df_limpio.to_excel(ruta_archivo, index=False)

print("El archivo ha sido limpiado y solo contiene los encabezados.")


#######################
## Ejecución del RPA ##
#######################

# Inicializar el controlador del ratón
mouse = Controller()

def click_raton_posicion(x, y, delay=1):
    """Mueve el ratón a la posición (x, y) y hace clic izquierdo."""
    mouse.position = (x, y)
    mouse.click(Button.left, 1)
    time.sleep(delay)
    
def click_izquierdo(x, y, delay=1):
    """Mueve el ratón a la posición (x, y) y hace clic izquierdo."""
    mouse.position = (x, y)
    mouse.click(Button.right, 1)
    time.sleep(delay)
    
def escribir_texto(texto):
    """Escribe el texto proporcionado usando pyautogui."""
    pyautogui.write(texto)
    time.sleep(1)

def borrar_caracteres(num_caracteres):
    """Borra una cantidad específica de caracteres."""
    for _ in range(num_caracteres):
        pyautogui.press("backspace")
    time.sleep(1)

def obtener_texto_del_portapapeles(intentos=5, espera=0.5, tiempo_limpiar=1.0):
    """
    Obtiene el texto del portapapeles, espera un tiempo y lo limpia.

    Args:
        intentos (int): Número máximo de intentos para acceder al portapapeles.
        espera (float): Tiempo de espera entre intentos en segundos.
        tiempo_limpiar (float): Tiempo en segundos que se espera antes de limpiar el portapapeles.

    Returns:
        str: Texto obtenido del portapapeles.

    Raises:
        RuntimeError: Si no se puede acceder al portapapeles después de varios intentos.
    """
    for intento in range(1, intentos + 1):
        try:
            # Intentar abrir el portapapeles
            texto = pyperclip.paste()
            print(f"Texto obtenido del portapapeles: {texto}")
            time.sleep(tiempo_limpiar)  # Esperar antes de limpiar
            pyperclip.copy("")  # Limpiar el portapapeles
            print("El portapapeles ha sido limpiado.")
            return texto
        except pyperclip.PyperclipException as e:
            print(f"Intento {intento} de {intentos} fallido: {e}")
            time.sleep(espera)
        except Exception as e:
            print(f"Error inesperado en el intento {intento}: {e}")
            time.sleep(espera)
    
    raise RuntimeError("No se pudo acceder al portapapeles después de varios intentos.")

def realizar_automatizacion(df_final):
    no_validadas = []

    try:
	# Secuencia inicial de clics
        click_raton_posicion(180, 751, delay=2)   # Icono del remoto
        click_izquierdo(408, 242, delay=0.3)  # Clic izquierdo sobre el nombre del servidor donde va acceder
        click_raton_posicion(456, 305, delay=0.3)  # Clic en copiar
        nom_ser = obtener_texto_del_portapapeles() #Nombre de servidor
        print(f"El nombre sel servidor es: {nom_ser}")
        
        # Verificación del nombre del servidor
        if nom_ser == "osiris":
            # Si el nombre del servidor es "osiris", sigue el código normalmente
            print("El servidor es osiris, continuando con el código...")
        else:
            # Si no es "osiris", ingresamos "osiris" en el campo
            print("El servidor no es osiris, ingresando el nombre osiris...")
            click_raton_posicion(504, 241, delay=0.3)  # Clic en el campo donde se debe ingresar el nombre
            borrar_caracteres(15) #Borrar información
            escribir_texto("osiris")  # Ingresa "osiris"
            print("El nombre 'osiris' ha sido ingresado.")
            
        click_raton_posicion(568, 346, delay=30)  # Botón de conectar
        click_raton_posicion(268, 752, delay=0.2)  # Icono de Siesa
        click_raton_posicion(268, 752, delay=30)  # Icono de Siesa segunda vez
        
        # Función para verificar el estado de Caps Lock
        def is_capslock_active():
            # GetKeyState devuelve un valor diferente de 0 si está activado
            return ctypes.windll.user32.GetKeyState(0x14) != 0

        # Desactivar Caps Lock si está activado
        if is_capslock_active():
            pyautogui.press('capslock')  # Desactiva Caps Lock si está activado
            
        click_raton_posicion(562, 365, delay=5)    # Apartado de usuario
        borrar_caracteres(15)
        escribir_texto("automateso")  # Usuario
        click_raton_posicion(566, 391)  # Apartado de contraseña
        escribir_texto("edemco2024")  # Contraseña
        click_raton_posicion(566, 418, delay=30)  # Ingresar       

        # Navegación en el menú
        click_raton_posicion(663, 32, delay=5)    # Cuentas por pagar
        click_raton_posicion(664, 86, delay=5)   # Programaciones de pagos
        click_raton_posicion(193, 190, delay=5)  # Hacer clic en el campo de fecha
        
        # Función para verificar el estado de Num Lock
        def is_numlock_active():
            # GetKeyState devuelve 0 si está desactivado y 1 si está activado
            return ctypes.windll.user32.GetKeyState(0x90) != 0

        # Desactivar Num Lock si está activado
        if is_numlock_active():
            pyautogui.press('numlock')  # Desactiva el teclado numérico si está activado

        # Usar las teclas de flecha de navegación normales
        pyautogui.press('up')   # Presionar flecha hacia arriba
        
        click_raton_posicion(490, 602, delay=5)   # Monto máximo
        click_raton_posicion(241, 269, delay=5)   # Apartado del proveedor

        grouped_df = df_final.groupby('NIT')  # Agrupamos el DataFrame por NIT

        for nit_value, group in grouped_df:
            # Inicialización de conjuntos y contadores
            facturas_temporales = set(group['FAC'])  # Conjunto de facturas para el NIT actual
            factura_counts = {factura: 0 for factura in facturas_temporales}  # Contador para cada factura
            facturas_encontradas = {factura: False for factura in facturas_temporales}  # Seguimiento de facturas encontradas
            portapapeles_vacio_contador = 0  # Contador para textos vacíos del portapapeles

            # Ingreso del NIT en la interfaz
            escribir_texto(nit_value)
            click_raton_posicion(241, 269, delay=8)
            pyautogui.press('enter')
            click_raton_posicion(601, 635, delay=5)
            click_raton_posicion(704, 195, delay=7)

            # Validación por columnas
            def validar_factura_en_columna(x, y, x_copiar, y_copiar, x_apartado, y_apartado):
                """Valida facturas dentro de una columna específica."""
                nonlocal portapapeles_vacio_contador

                # Si las coordenadas son None, solo clic en apartado
                if x is None or y is None:
                    click_raton_posicion(x_apartado, y_apartado, delay=0.6)
                    return

                # Proceso de validación de facturas
                click_raton_posicion(x, y, delay=0.7)
                click_izquierdo(x, y, delay=0.3)
                click_raton_posicion(x_copiar, y_copiar, delay=0.7)
                factura = obtener_texto_del_portapapeles()
                print(f"Factura leída: {factura}")

                # Si el portapapeles está vacío, incrementar el contador
                if not factura.strip():
                    portapapeles_vacio_contador += 1
                    print(f"Portapapeles vacío. Contador: {portapapeles_vacio_contador}")
                    return

                # Reiniciar el contador si se encuentra una factura válida
                portapapeles_vacio_contador = 0

                if factura in facturas_temporales:
                    if factura_counts[factura] < 3:
                        click_raton_posicion(x_apartado, y_apartado, delay=0.4)
                        factura_counts[factura] += 1
                        facturas_encontradas[factura] = True
                        print(f"Factura {factura} marcada. Contador: {factura_counts[factura]}")

                    if factura_counts[factura] == 3:
                        print(f"Factura {factura} alcanzó el límite y será eliminada.")
                        facturas_temporales.discard(factura)

            # Coordenadas de las columnas
            columnas = [
                (260, 432, 310, 496, 74, 430),  # Columna 1
                (260, 448, 310, 512, 74, 447),  # Columna 2    
                (260, 464, 310, 526, 74, 463),  # Columna 3
                (260, 480, 310, 543, 74, 480),  # Columna 4                        
                (260, 496, 310, 562, 74, 494),  # Columna 5
                (260, 512, 310, 353, 74, 511),  # Columna 6                       
                (260, 528, 310, 366, 74, 527),  # Columna 7
                (260, 544, 310, 383, 74, 543),  # Columna 8                        
                (260, 560, 310, 401, 74, 559),  # Columna 9
                (260, 576, 315, 415, 74, 575),  # Columna 10                    
                (260, 592, 310, 433, 74, 592),  # Columna 11
                (260, 608, 310, 448, 74, 607),  # Columna 12                       
                (260, 624, 310, 462, 74, 623),  # Columna 13
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 14
                (260, 512, 310, 353, 74, 511),  # Columna 15                      
                (260, 528, 310, 366, 74, 527),  # Columna 16
                (260, 544, 310, 383, 74, 543),  # Columna 17                       
                (260, 560, 310, 401, 74, 559),  # Columna 18
                (260, 576, 315, 415, 74, 575),  # Columna 19                    
                (260, 592, 310, 433, 74, 592),  # Columna 20
                (260, 608, 310, 448, 74, 607),  # Columna 21                       
                (260, 624, 310, 462, 74, 623),  # Columna 22
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 23
                (260, 512, 310, 353, 74, 511),  # Columna 24                      
                (260, 528, 310, 366, 74, 527),  # Columna 25
                (260, 544, 310, 383, 74, 543),  # Columna 26                       
                (260, 560, 310, 401, 74, 559),  # Columna 27
                (260, 576, 315, 415, 74, 575),  # Columna 28                    
                (260, 592, 310, 433, 74, 592),  # Columna 29
                (260, 608, 310, 448, 74, 607),  # Columna 30                       
                (260, 624, 310, 462, 74, 623),  # Columna 31
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado                       
                (260, 496, 310, 562, 74, 494),  # Columna 33
                (260, 512, 310, 353, 74, 511),  # Columna 34                      
                (260, 528, 310, 366, 74, 527),  # Columna 35
                (260, 544, 310, 383, 74, 543),  # Columna 36                       
                (260, 560, 310, 401, 74, 559),  # Columna 37
                (260, 576, 315, 415, 74, 575),  # Columna 38                    
                (260, 592, 310, 433, 74, 592),  # Columna 39
                (260, 608, 310, 448, 74, 607),  # Columna 40                       
                (260, 624, 310, 462, 74, 623),  # Columna 41
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 42
                (260, 512, 310, 353, 74, 511),  # Columna 43                      
                (260, 528, 310, 366, 74, 527),  # Columna 44
                (260, 544, 310, 383, 74, 543),  # Columna 45                       
                (260, 560, 310, 401, 74, 559),  # Columna 46
                (260, 576, 315, 415, 74, 575),  # Columna 47                    
                (260, 592, 310, 433, 74, 592),  # Columna 48
                (260, 608, 310, 448, 74, 607),  # Columna 49                       
                (260, 624, 310, 462, 74, 623),  # Columna 50
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 51
                (260, 512, 310, 353, 74, 511),  # Columna 52                      
                (260, 528, 310, 366, 74, 527),  # Columna 53
                (260, 544, 310, 383, 74, 543),  # Columna 54                       
                (260, 560, 310, 401, 74, 559),  # Columna 55
                (260, 576, 315, 415, 74, 575),  # Columna 56                    
                (260, 592, 310, 433, 74, 592),  # Columna 57
                (260, 608, 310, 448, 74, 607),  # Columna 58                       
                (260, 624, 310, 462, 74, 623),  # Columna 59
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 60
                (260, 512, 310, 353, 74, 511),  # Columna 61                      
                (260, 528, 310, 366, 74, 527),  # Columna 62
                (260, 544, 310, 383, 74, 543),  # Columna 63                       
                (260, 560, 310, 401, 74, 559),  # Columna 64
                (260, 576, 315, 415, 74, 575),  # Columna 65                    
                (260, 592, 310, 433, 74, 592),  # Columna 66
                (260, 608, 310, 448, 74, 607),  # Columna 67                       
                (260, 624, 310, 462, 74, 623),  # Columna 68
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 69
                (260, 512, 310, 353, 74, 511),  # Columna 70                      
                (260, 528, 310, 366, 74, 527),  # Columna 71
                (260, 544, 310, 383, 74, 543),  # Columna 72                       
                (260, 560, 310, 401, 74, 559),  # Columna 73
                (260, 576, 315, 415, 74, 575),  # Columna 74                    
                (260, 592, 310, 433, 74, 592),  # Columna 75
                (260, 608, 310, 448, 74, 607),  # Columna 76                       
                (260, 624, 310, 462, 74, 623),  # Columna 77
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 78
                (260, 512, 310, 353, 74, 511),  # Columna 79                      
                (260, 528, 310, 366, 74, 527),  # Columna 80
                (260, 544, 310, 383, 74, 543),  # Columna 81                       
                (260, 560, 310, 401, 74, 559),  # Columna 82
                (260, 576, 315, 415, 74, 575),  # Columna 83                    
                (260, 592, 310, 433, 74, 592),  # Columna 84
                (260, 608, 310, 448, 74, 607),  # Columna 85                       
                (260, 624, 310, 462, 74, 623),  # Columna 86
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 87
                (260, 512, 310, 353, 74, 511),  # Columna 88                      
                (260, 528, 310, 366, 74, 527),  # Columna 89
                (260, 544, 310, 383, 74, 543),  # Columna 90                       
                (260, 560, 310, 401, 74, 559),  # Columna 91
                (260, 576, 315, 415, 74, 575),  # Columna 92                    
                (260, 592, 310, 433, 74, 592),  # Columna 93
                (260, 608, 310, 448, 74, 607),  # Columna 94                       
                (260, 624, 310, 462, 74, 623),  # Columna 95
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (None, None, None, None, 980, 647),  # Solo clic de apartado
                (260, 496, 310, 562, 74, 494),  # Columna 96
                (260, 512, 310, 353, 74, 511),  # Columna 97                      
                (260, 528, 310, 366, 74, 527),  # Columna 98
                (260, 544, 310, 383, 74, 543),  # Columna 99                       
                (260, 560, 310, 401, 74, 559),  # Columna 100
                (260, 576, 315, 415, 74, 575),  # Columna 101                   
                (260, 592, 310, 433, 74, 592),  # Columna 102
                (260, 608, 310, 448, 74, 607),  # Columna 103                      
                (260, 624, 310, 462, 74, 623),  # Columna 104
            ]

            # Iterar por columnas y validar facturas
            for x, y, x_copiar, y_copiar, x_apartado, y_apartado in columnas:
                if not facturas_temporales:
                    break

                # Validamos la factura de la columna actual
                validar_factura_en_columna(x, y, x_copiar, y_copiar, x_apartado, y_apartado)

                # Si el portapapeles está vacío más de 1 veces, pasar al siguiente NIT
                if portapapeles_vacio_contador > 1:
                    print(f"Portapapeles vacío más de 1 veces. Pasando al siguiente NIT: {nit_value}")
                    break

            # Registrar facturas no validadas
            for fac_valor in facturas_temporales:
                if not facturas_encontradas[fac_valor]:
                    no_validadas.append({"NIT": nit_value, "FAC": fac_valor})

            # Reinicio de interfaz para el siguiente NIT
            click_raton_posicion(222, 63, delay=2)
            pyautogui.press('enter')
            click_raton_posicion(24, 156, delay=1)
            pyautogui.press('enter')

            # Validación de fecha
            click_izquierdo(193, 190, delay=0.3)
            click_raton_posicion(233, 249, delay=0.3)
            nom_fe = obtener_texto_del_portapapeles()
            print(f"La fecha es: {nom_fe}")

            if nom_fe != "2026":
                click_raton_posicion(193, 190, delay=0.3)
                borrar_caracteres(4)
                escribir_texto("2026")
                print("El nombre '2026' ha sido ingresado.")

            click_raton_posicion(241, 269, delay=1)
            borrar_caracteres(10)
            
         # Almacenar las facturas no validadas en un archivo Excel
        if no_validadas:
              df_no_validadas = pd.DataFrame(no_validadas)
              df_no_validadas.to_excel(
                  r"C:\Users\usuario\Documents\AUTOMATIZACION_PAGO_PROVEEDORES\Resultados\facturas_no_validadas.xlsx",
                  index=False,
                  engine='openpyxl'
              )
              print(f"Se han guardado {len(no_validadas)} facturas no validadas en 'facturas_no_validadas.xlsx'.")
        
        #################################
        ## Envio de Correo electronico ##
        #################################
        
        # Definir información del correo
        sender_email = "matic@edemco.co"
        receiver_emails = ["juan.sossa@edemco.co", "brayan.rebolledo@edemco.co"]
        subject = "ARCHIVO DE PROGRAMACIÓN DEL DIA DE HOY"
        body = body = '''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Notificación Edemco</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    background-color: #f4f4f4;
                }
                .container {
                    max-width: 600px;
                    margin: 20px auto;
                    background-color: #fff;
                    border: 2px solid #4CAF50;
                    border-radius: 10px;
                    padding: 20px;
                    position: relative;
                    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                }
                .logo {
                    position: absolute;
                    top: 10px;
                    right: 10px;
                    max-width: 100px;
                }
                .content {
                    text-align: center;
                    margin-top: 30px;
                }
                .content p {
                    color: #333;
                    font-size: 16px;
                    line-height: 1.5;
                    margin: 15px 0;
                }
                .content a {
                    color: #fff;
                    background-color: #4CAF50;
                    text-decoration: none;
                    padding: 10px 20px;
                    border-radius: 5px;
                    display: inline-block;
                    margin-top: 10px;
                    font-weight: bold;
                }
                .content a:hover {
                    background-color: #45a049;
                }
                footer {
                    margin-top: 20px;
                    text-align: center;
                    font-size: 14px;
                    color: #777;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <img src="https://www.edemco.co/uploads/files/logo-edemco.png" alt="Logo Edemco" class="logo">
                <div class="content">
                    <p style="font-weight: bold;">Cordial saludo.</p>
                    <p>
                        A continuación encontrarás el archivo con las facturas no encontradas durante la ejecución de la programación de pagos en Siesa correspondiente al día de hoy para su revisión.
                    </p>
                    <p>
                        Por favor, cualquier novedad relacionada con la información adjunta, les solicitamos escalarla mediante el buzón de soporte:
                    </p>
                    <a href="mailto:support@edemco.co">Contactar soporte</a>
                    <p>
                        Seguimos mejorando con buena energía.<br>
                        ¡Que tengas un excelente día!
                    </p>
                </div>
                <footer>
                    © 2025 Edemco. Todos los derechos reservados.
                </footer>
            </div>
        </body>
        </html>


        '''
        # Crear un objeto EmailMessage
        message = EmailMessage()
        message.set_content(body)
        message["Subject"] = subject
        message["From"] = sender_email
        message["To"] = ", ".join(receiver_emails)  # Unir múltiples destinatarios con una coma

        # Establecer el contenido HTML
        message.add_alternative(body, subtype='html')

        # Adjuntar el primer archivo
        with open(r"C:\Users\usuario\Documents\AUTOMATIZACION_PAGO_PROVEEDORES\Resultados\Facturas_no_validadas.xlsx", "rb") as attachment1:
            message.add_attachment(attachment1.read(), maintype="application", subtype="octet-stream", filename="Facturas_no_validadas.xlsx")

        # Adjuntar el segundo archivo
        with open(r"C:\Users\usuario\Documents\AUTOMATIZACION_PAGO_PROVEEDORES\Resultados\Todas_Facturas.xlsx", "rb") as attachment2:
            message.add_attachment(attachment2.read(), maintype="application", subtype="octet-stream", filename="Todas_Facturas.xlsx")

        # Conectar al servidor SMTP (Microsoft 365)
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()  # Iniciar una conexión segura
            server.login(sender_email, "")
            server.send_message(message)

        print("Correo enviado con éxito.")        
        
    except Exception as e:
        print(f"Error durante la automatización: {str(e)}")

# Llamar a la función de automatización
realizar_automatizacion(df_final)



