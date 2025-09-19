from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pynput import keyboard, mouse
import threading
import time
import os
import pandas as pd
import datetime

# --- Parámetros ---
shortcut_path = r'.\RSIRAT\Actualiza RSIRAT.lnk'  # Ahora dentro de la carpeta RSIRAT
interferencia_detectada = False

# --- Función para bloquear pantalla ---
def bloquear():
    global interferencia_detectada
    if not interferencia_detectada:
        interferencia_detectada = True
        print("Interferencia detectada. Bloqueando pantalla...")
        os.system("rundll32.exe user32.dll,LockWorkStation")

# --- Listeners de teclado y mouse ---
def iniciar_monitoreo():
    def on_key(key):
        bloquear()

    def on_move(x, y):
        bloquear()

    def on_click(x, y, button, pressed):
        if pressed:
            bloquear()

    keyboard_listener = keyboard.Listener(on_press=on_key)
    mouse_listener = mouse.Listener(on_move=on_move, on_click=on_click)

    keyboard_listener.start()
    mouse_listener.start()

# --- Ejecutar RSIRAT ---
try:
    os.startfile(shortcut_path)
    print("RSIRAT iniciado desde acceso directo.")
except Exception as e:
    print(f"Error al iniciar RSIRAT: {e}")
    exit()

# --- Esperar a que cargue ---
time.sleep(15)  # Puedes ajustar el tiempo

# --- Conectar con la aplicación ---
try:
    app = Application(backend="uia").connect(title="SIRAT")
    windows = app.windows()
    if not windows:
        print("No se encontró ninguna ventana de SIRAT.")
        exit()
    target_window = app.window(handle=windows[-1].handle)
    print("Conectado a la ventana de RSIRAT.")
except Exception as e:
    print(f"Error al conectar con RSIRAT: {e}")
    exit()

# --- Iniciar monitoreo en segundo plano (después de conectar RSIRAT) ---
monitoreo_thread = threading.Thread(target=iniciar_monitoreo, daemon=True)
monitoreo_thread.start()

# --- Leer contraseña ---
with open(r'CREDENCIAL\contrasena.txt', 'r', encoding='utf-8') as f:
    password = f.read().strip()

# --- Leer y procesar Excel ---
df = pd.read_excel(r'DATA\EXPEDIENTES.xlsx')

# Normalizar dependencias
def map_dependencia(dep):
    dep_str = str(dep).zfill(4)
    if dep_str.endswith('0021') or dep_str.endswith('21'):
        return '0021 I.R. Lima - PRICO'
    elif dep_str.endswith('0023') or dep_str.endswith('23'):
        return '0023 I.R. Lima - MEPECO'
    else:
        return dep  # O manejar otros casos

df['DEPENDENCIA_NORM'] = df['DEPENDENCIA'].apply(map_dependencia)

# Agrupar y ordenar por dependencia normalizada
dependencias_ordenadas = df['DEPENDENCIA_NORM'].drop_duplicates().tolist()

for dep in dependencias_ordenadas:
    grupo = df[df['DEPENDENCIA_NORM'] == dep]
    print(f"Procesando {len(grupo)} expedientes para dependencia: {dep}")

    # --- Paso 1: Escribir la dependencia ---
    try:
        combo_edit = target_window.child_window(control_type="Edit", found_index=0)
        combo_edit.set_edit_text(dep)
        print(f"Dependencia escrita: {dep}")
    except Exception as e:
        print("No se pudo escribir la dependencia:", e)

    # --- Paso 2: Ingresar la contraseña ---
    try:
        edits = target_window.descendants(control_type="Edit")
        password_field = next((e for e in edits if e.element_info.automation_id == "1005"), edits[-1])
        password_field.set_focus()
        password_field.type_keys(password, with_spaces=True)
        print("✅ Contraseña ingresada.")
    except Exception as e:
        print("❌ No se pudo ingresar la contraseña:", e)

    # --- Paso 3: Activar "Aceptar" usando Alt + A ---
    try:
        send_keys('%a')  # Alt + A
        print("✅ Se hizo clic en 'Aceptar' usando Alt+A.")
    except Exception as e:
        print("❌ No se pudo activar 'Aceptar' con Alt+A:", e)

    time.sleep(5)  # Espera a que cargue el menú

    for idx, row in grupo.iterrows():
        expediente = str(row['EXPEDIENTE'])

        # --- Navegar Menú: SIRAT > Cobranza Coactiva ---
        try:
            # Menú de Opciones > SIRAT
            menu = target_window.child_window(title="Menú de Opciones", control_type="Menu")
            sirat_menu = menu.child_window(title="SIRAT", control_type="MenuItem")
            sirat_menu.click_input()
            # Cobranza Coactiva
            cobranza = sirat_menu.child_window(title="Cobranza Coactiva", control_type="MenuItem")
            cobranza.click_input()
            # Exp. Cob. Coactiva - Individual (doble clic)
            exp_ind = cobranza.child_window(title="Exp. Cob. Coactiva - Individual", control_type="MenuItem")
            exp_ind.double_click_input()
            print("✅ Menú navegado correctamente.")
        except Exception as e:
            print("❌ Error navegando el menú:", e)
            continue

        time.sleep(2)  # Espera a que cargue la ventana

        # --- Ventana Selección de Expediente Coactivo ---
        try:
            sel_exp = target_window.child_window(title="Selección de Expediente Coactivo", control_type="Window")
            exp_edit = sel_exp.child_window(control_type="Edit", found_index=0)
            exp_edit.set_edit_text(expediente)
            send_keys('{TAB}')
            print(f"✅ Expediente digitado: {expediente}")
            time.sleep(1)

            # Verificar campo "Grupo"
            grupo_edit = sel_exp.child_window(title="Grupo", control_type="Edit")
            grupo_val = grupo_edit.get_value() if hasattr(grupo_edit, 'get_value') else grupo_edit.window_text()
            if not grupo_val.strip():
                df.at[idx, 'RESULTADO'] = "Sin ejecutor asignado"
                print(f"❌ Expediente {expediente}: Sin ejecutor asignado.")
                # Cerrar ventana si es necesario
                sel_exp.close()
                continue

            # Confirmar con Alt+A
            send_keys('%a')
            print("✅ Expediente confirmado con Alt+A.")
        except Exception as e:
            print(f"❌ Error en selección de expediente {expediente}:", e)
            continue

        time.sleep(2)  # Espera a que regrese al menú

        # --- Menú: Proceso de Embargo > Trabar Embargo ---
        try:
            proceso_embargo = cobranza.child_window(title="Proceso de Embargo", control_type="MenuItem")
            proceso_embargo.click_input()
            trabar_embargo = proceso_embargo.child_window(title="Trabar Embargo", control_type="MenuItem")
            trabar_embargo.click_input()
            print("✅ Navegación hasta Trabar Embargo completada.")
        except Exception as e:
            print("❌ Error navegando hasta Trabar Embargo:", e)
            continue

# Eliminar columna auxiliar antes de guardar
if 'DEPENDENCIA_NORM' in df.columns:
    df.drop(columns=['DEPENDENCIA_NORM'], inplace=True)

# Asegurar que RUC y EXPEDIENTE sean texto
df['RUC'] = df['RUC'].astype(str)
df['EXPEDIENTE'] = df['EXPEDIENTE'].astype(str)

# Guardar resultados en una nueva carpeta con nombre personalizado
now = datetime.datetime.now()
folder_name = f"RESULTADO_RSIRAT_PRAC-AREYESS_NBH738_{now.day:02d}D_{now.month:02d}M_{now.year}_\
{now.hour:02d}H_{now.minute:02d}M_{now.second:02d}S"
output_dir = os.path.join(os.getcwd(), folder_name)
os.makedirs(output_dir, exist_ok=True)
output_excel = os.path.join(output_dir, "EXPEDIENTES_RESULTADO.xlsx")
df.to_excel(output_excel, index=False)
print(f"🏁 Script finalizado. Archivo generado en: {output_excel}\nMonitoreo de seguridad sigue activo.")