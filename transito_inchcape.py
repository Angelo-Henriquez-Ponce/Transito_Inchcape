# %%
import win32com.client
import getpass
from pathlib import Path
import pandas as pd
import time
from datetime import datetime, timedelta
import numpy as np
import os
from glob import glob
import subprocess


import getpass
from pathlib import Path

# 🔒 Usuario actual
usuario = getpass.getuser()
print(f"👤 Usuario detectado: {usuario}")

# ⚠️ Tkinter no se recomienda en Jupyter, pero esto lo forzará a intentarlo
try:
    from tkinter import Tk
    from tkinter.filedialog import askdirectory

    # Ocultar ventana raíz de Tkinter
    root = Tk()
    root.withdraw()

    # Mostrar selector de carpeta
    ruta_seleccionada = askdirectory(title="Selecciona carpeta donde se guardarán los archivos")
    root.destroy()

    if not ruta_seleccionada:
        print("❌ No se seleccionó ninguna carpeta. Proceso detenido.")
        raise Exception("Proceso detenido por usuario")

    # Convertir a Path
    ruta_base = Path(ruta_seleccionada)
    print(f"📂 Carpeta seleccionada: {ruta_base}")

except Exception as e:
    print(f"⚠️ Error o cancelación: {e}")


# %%
# Inicializar SAP GUI
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
session.findById("wnd[0]").maximize()

# %%


# ✅ Mensaje para el usuario
print("📥 Iniciando descarga de datos desde SAP (ME2N)...")

# === ME2N ===
session.findById("wnd[0]/tbar[0]/okcd").text = "ME2N"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV"





# %%
#  Abrir selección múltiple para clase de documento (SELPA)
session.findById("wnd[0]/usr/btn%_SELPA_%_APP_%-VALU_PUSH").press()

# Agregar valores WE101 y WE104
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "WE101"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "WE104"

session.findById("wnd[1]").sendVKey(8)

# %%
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1305"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "1344"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1335"
# session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "3261"
session.findById("wnd[1]").sendVKey(8)
session.findById("wnd[0]").sendVKey(8)
session.findById("wnd[0]/tbar[1]/btn[23]").press

session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").select()


# %%
# === Exportar ALV usando la carpeta seleccionada previamente ===
try:
    # Abrir menú de exportación
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    # Aceptar popup intermedio (si aparece)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    # Guardar en wnd[1]
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(ruta_base)
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME2N.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(str(ruta_base))
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

except Exception as e:
    print(f"❌ Error al exportar ME2N: {e}")

# %%
# Intentar cerrar Excel automáticamente si se abrió tras la exportación desde SAP
subprocess.call("taskkill /f /im excel.exe", shell=True)

# %%

# === Volver a SAP Easy Access (robusto) ==
import time as _time

# 1) Cerrar pop-ups si quedó alguno abierto (Cancel)
for _ in range(2):
    try:
        session.findById("wnd[1]/tbar[0]/btn[12]").press()  # Cancel
        _time.sleep(0.2)
    except:
        break

# 2) Volver atrás con F3 un par de veces
for _ in range(3):
    try:
        session.findById("wnd[0]").sendVKey(3)  # F3 Back
        _time.sleep(0.2)
    except:
        break


# %%
### === LIMPIEZA PREVIA A ME80FN === ###
# Esperar un segundo a que el archivo esté disponible
time.sleep(1)

# Leer archivo recién guardado
me2n_tmp = pd.read_excel(ruta_base / 'ME2N.XLSX', dtype=str, engine='openpyxl', skiprows=[1])




# %%
# Eliminar filas nulas o vacías
me2n_tmp = me2n_tmp[me2n_tmp['Documento compras'].notna() & (me2n_tmp['Documento compras'].str.strip() != '')]

# Eliminar clases de pedido no deseadas
clases_a_eliminar = ['Z250', 'Z210', 'Z520', 'Z420']
me2n_tmp = me2n_tmp[~me2n_tmp['Cl.documento compras'].isin(clases_a_eliminar)]

# Extraer órdenes de compra válidas
ocs_validas = me2n_tmp['Documento compras'].dropna().drop_duplicates()
ocs_validas = ocs_validas[ocs_validas.str.strip() != '']

if ocs_validas.empty:
    print("⚠️ No hay OCs válidas. Proceso detenido.")
    exit(1)


# %%
# Guardar archivo temporal
archivo_tmp = ruta_base / "ordenes_tmp.txt"
ocs_validas.to_csv(archivo_tmp, index=False, header=False)
ocs_validas.to_clipboard(index=False, header=False)
print(f"📋 {len(ocs_validas)} OCs válidas copiadas al portapapeles.")

# %%

# %%
# === ME80FN con filtros ===
print("🚀 Ejecutando ME80FN con filtros desde ME2N...")

session.findById("wnd[0]/tbar[0]/okcd").text = "ME80FN"
session.findById("wnd[0]").sendVKey(0)

# Filtro por OC desde portapapeles
session.findById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()  # Seleccionar todo
session.findById("wnd[1]/tbar[0]/btn[0]").press()   # Transferir selección
session.findById("wnd[1]/tbar[0]/btn[8]").press()   # Confirmar

# Filtro por tipo de documento (si aplica)
session.findById("wnd[0]/usr/ctxtSP$00011-LOW").setFocus()
session.findById("wnd[0]/usr/btn%_SP$00011_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[16]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()

# Ejecutar búsqueda
session.findById("wnd[0]").sendVKey(8)


# %%
# Esperar un momento para cargar el grid
import time
time.sleep(2)

# Exportar
# Esperar que el grid esté disponible (espera activa con timeout)
print("⏳ Esperando que el grid de resultados de ME80FN esté disponible...")

grid_ok = False
for i in range(10):  # Hasta 10 intentos, espera de 1s entre cada uno
    try:
        grid = session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell")
        grid_ok = True
        break
    except:
        time.sleep(1)

if not grid_ok:
    print("❌ No se encontró el grid de resultados. Verifica si hubo datos.")
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    exit(1)



# Continuar solo si el grid fue encontrado
print("✅ Grid cargado correctamente. Accediendo al historial...")
grid.pressToolbarContextButton("DETAIL_MENU")
grid.selectContextMenuItem("TO_HIST")

# 🧩 Selección del layout desde menú contextual
try:
    print("📦 Abriendo menú contextual para cargar layout...")
    grid = session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell")
    grid.pressToolbarContextButton("&MB_VARIANT")  # Abre menú contextual de variantes
    grid.selectContextMenuItem("&LOAD")  # Selecciona opción 'Cargar'

    print("🎯 Seleccionando primer layout disponible...")
    layout_table = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
    layout_table.currentCellColumn = "TEXT"
    layout_table.selectedRows = "0"
    layout_table.clickCurrentCell()

except Exception as e:
    print(f"❌ Error al seleccionar el layout: {e}")
    exit(1)


session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
session.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_HIST/shellcont/shell").selectContextMenuItem("&XXL")

# Guardar el archivo exportado
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(ruta_base)
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "me80fn.XLSX"
session.findById("wnd[1]/tbar[0]/btn[11]").press()

# Cerrar ventanas
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)

# Intentar cerrar Excel automáticamente si se abrió tras la exportación desde SAP
subprocess.call("taskkill /f /im excel.exe", shell=True)

# %%
#Ocupando las mismas OC

ocs_validas.to_csv(archivo_tmp, index=False, header=False)
ocs_validas.to_clipboard(index=False, header=False)
print(f"📋 {len(ocs_validas)} OCs válidas copiadas al portapapeles.")

# %%
# ============================================================
# VL06IF: Lista entregas  ejecutar y exportar
# ============================================================

import time as _time
import subprocess
import pandas as pd
from pathlib import Path

archivo_tmp = ruta_base / "ordenes_tmp.txt"   # generado previamente
salida_vl06if = ruta_base / "VL06IF.xlsx"

print("🚚 Abriendo VL06IF…")
session.findById("wnd[0]/tbar[0]/okcd").text = "VL06IF"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").maximize()

# 1) Limpiar fechas (si existen en tu layout)
for fid in ["wnd[0]/usr/ctxtIT_LFDAT-LOW", "wnd[0]/usr/ctxtIT_LFDAT-HIGH"]:
    try:
        session.findById(fid).text = ""
    except:
        pass

# 2) Preparar portapapeles (respaldo) con las OCs
if archivo_tmp.exists():
    try:
        ocs_clip = pd.read_csv(archivo_tmp, header=None, dtype=str, names=["OC"])
        ocs_clip["OC"].to_clipboard(index=False, header=False)
        print(f"📋 {len(ocs_clip)} OCs copiadas al portapapeles.")
    except Exception as e:
        print(f"⚠️ No se pudo preparar el portapapeles: {e}")
else:
    print("ℹ️ No existe ordenes_tmp.txt; continuarás sin filtro por OC.")

# 3) Abrir selección múltiple para EBELN y cargar OCs
try:
    session.findById("wnd[0]/usr/btn%_IT_EBELN_%_APP_%-VALU_PUSH").press()
    # Preferir pegar desde portapapeles (simple y rápido)
    try:
        session.findById("wnd[1]/tbar[0]/btn[24]").press()   # Pegar
        print("📋 OCs pegadas desde portapapeles.")
    except:
        # Alternativa: importar por archivo si tu popup lo permite
        try:
            session.findById("wnd[1]/tbar[0]/btn[23]").press()   # Importar archivo
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = str(archivo_tmp.parent)
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = archivo_tmp.name
            session.findById("wnd[2]/tbar[0]/btn[11]").press()
            print("📂 OCs importadas desde archivo.")
        except Exception as e:
            print(f"❌ No se pudo importar/pegar OCs: {e}")
            # Cerrar popup y seguir sin filtro
            try: session.findById("wnd[1]/tbar[0]/btn[12]").press()
            except: pass
    # Transferir y confirmar
    try:
        session.findById("wnd[1]/tbar[0]/btn[8]").press()    # Transferir
    except:
        pass
except:
    print("ℹ️ No se encontró el botón de selección múltiple para EBELN; se ejecutará sin filtro.")

# 4) Ejecutar (F8)
print("▶️ Ejecutando consulta…")
try:
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
except:
    session.findById("wnd[0]").sendVKey(8)



# %%

_time.sleep(1.0)

print("✅ Marcando todo (F5)…")
try:
    session.findById("wnd[0]/tbar[1]/btn[5]").press()  # botón F5 clásico
except:
    try:
        session.findById("wnd[0]").sendVKey(5)         # F5
    except:
        print("⚠️ No se pudo marcar todo. Continúo…")

print("🔄 Cambiando a vista de posición (Shift+F6)…")
cambio_ok = False
for btn_idx in [18, 19, 21]:                           # índices frecuentes
    try:
        session.findById(f"wnd[0]/tbar[1]/btn[{btn_idx}]").press()
        cambio_ok = True
        break
    except:
        pass
if not cambio_ok:
    try:
        session.findById("wnd[0]").sendVKey(18)         # Shift+F6 (muchos layouts)
        cambio_ok = True
    except:
        print("⚠️ No se pudo cambiar a vista de posición. Intento exportar igual…")

# %%
# === Abrir selector de disposición (btn[33]) y aceptar con sendVKey(2) ===
import time as _time

# 1) Abrir el popup de disposición / layout
session.findById("wnd[0]/tbar[1]/btn[33]").press()

# 2) Esperar a que aparezca la ventana wnd[1]
for _ in range(25):
    try:
        session.findById("wnd[1]")
        break
    except:
        _time.sleep(0.1)
else:
    raise RuntimeError("❌ No apareció el popup de disposición (wnd[1]).")

# 3) Intentar enfocar la celda/label (coordenadas típicas)
focado = False
for lbl in [
    "wnd[1]/usr/lbl[1,3]",  # la que mostraste
    "wnd[1]/usr/lbl[1,4]",  # alternativa frecuente
    "wnd[1]/usr/lbl[0,3]",  # por si el índice empieza en 0
]:
    try:
        obj = session.findById(lbl)
        obj.setFocus()
        try:
            obj.caretPosition = 1  # no siempre aplica en lbl, pero no hace daño
        except:
            pass
        focado = True
        break
    except:
        pass

# 4) Confirmar selección (equivalente a tu sendVKey 2). Fallback: botón OK o Enter.
if focado:
    try:
        session.findById("wnd[1]").sendVKey(2)  # mismo que tu VB
    except:
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK
        except:
            session.findById("wnd[1]").sendVKey(0)             # Enter
else:
    # Si no pudimos enfocar el label, intentamos directamente OK / Enter
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        session.findById("wnd[1]").sendVKey(0)

print("✅ Disposición aplicada desde el popup (btn[33] + sendVKey(2)).")



# %%
# === VL06IF: Lista → Exportar → Hoja de cálculo, guardar en ruta_base ===
import subprocess

nombre_salida = "vl06i.XLSX"  # ajusta si quieres otro nombre

# Abrir: Lista → Exportar → Hoja de cálculo
session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").select()
session.findById("wnd[1]/tbar[0]/btn[0]").press()  # "Hoja de cálculo"

# Asignar carpeta y nombre
session.findById("wnd[1]/usr/ctxtDY_PATH").text = str(ruta_base)
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_salida

# (Opcional) foco/caret como en VB
try:
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(str(ruta_base))
except:
    pass

# Confirmar guardado (algunos sistemas usan btn[11], otros btn[0])
try:
    session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Guardar
except:
    session.findById("wnd[1]/tbar[0]/btn[0]").press()   # OK

# Cerrar Excel “zombie” por si se abre al exportar
subprocess.call("taskkill /f /im excel.exe", shell=True)

print(f"✅ Archivo exportado en: {ruta_base / nombre_salida}")

# %%
# === Volver a SAP Easy Access (robusto) ===
import time as _time

# 1) Cerrar pop-ups si quedó alguno abierto (Cancel)
for _ in range(2):
    try:
        session.findById("wnd[1]/tbar[0]/btn[12]").press()  # Cancel
        _time.sleep(0.2)
    except:
        break

# 2) Volver atrás con F3 un par de veces
for _ in range(3):
    try:
        session.findById("wnd[0]").sendVKey(3)  # F3 Back
        _time.sleep(0.2)
    except:
        break



# %%
# 🛑 Verificar si la carpeta seleccionada está en OneDrive o SharePoint
ruta_base_str = str(ruta_base).lower()

if "onedrive" in ruta_base_str:
    print("⚠️ Advertencia: estás usando una carpeta sincronizada con OneDrive.")
    print("⏳ Esperando 60 segundos para permitir la sincronización de archivos...")
    time.sleep(60)

elif "sharepoint" in ruta_base_str or "inchcape" in ruta_base_str:
    print("⚠️ Advertencia: estás usando una carpeta compartida en SharePoint o red de empresa.")
    print("⏳ Esperando 60 segundos por posibles demoras en la sincronización...")
    time.sleep(60)

else:
    print("📁 Carpeta local detectada. Continuando sin espera...")

# %%
### === PROCESAMIENTO ===  NUEVO###

# 📂 Archivos fijos
archivo_me2n = ruta_base / 'ME2N.XLSX'
archivo_me80fn = ruta_base / 'me80fn.XLSX'
archivo_salida = ruta_base / 'Transito_Base.xlsx'


# 📁 Ruta
ruta = ruta_base

# %%
# 🔄 Cargar y limpiar archivo ME80FN
me80fn = pd.read_excel(ruta / 'me80fn.XLSX', 
                        dtype=str, 
                        engine='openpyxl')

# %%
# Una columna
me80fn = me80fn.drop(columns=['Cantidad'])

# %%
me80fn.rename(columns={
    'Cantidad.1': 'Cantidad'
}, inplace=True)

# %%


# 🔧 Limpiar columna 'Cantidad'
me80fn['Cantidad'] = (
    me80fn['Cantidad']
    .astype(str)
    .str.strip()
    .str.replace(r'\s+', '', regex=True)      # elimina espacios internos
    .str.replace('.', '', regex=False)        # elimina puntos de miles
    .str.replace(',', '.', regex=False)       # cambia coma decimal por punto
    .str.replace(r'[^\d\.-]', '', regex=True) # conserva solo dígitos, punto y guion
)

# 🔢 Convertir a numérico
me80fn['Cantidad'] = pd.to_numeric(me80fn['Cantidad'], errors='coerce')

# %%
import pandas as pd
import numpy as np
from collections import defaultdict

# ================ CONFIG ================
USE_POSICION = True     # True: clave = OC|Material|Posición. Si tu anulación NO distingue posición, pon False.
ADD_CENTRO_TO_KEY = False  # pon True si también quieres separar por Centro en la clave.
AMT_DECIMALS = 6        # redondeo para comparar montos (evita ruido float)

# ============== 1) Copia y normalización básica ==============
df = me80fn.copy()

# Asegura columnas clave
needed = ['Documento compras','Material','Cantidad','Tipo de historial de pedido']
missing = [c for c in needed if c not in df.columns]
if missing:
    raise KeyError(f"Faltan columnas en me80fn: {missing}")

# Tipos / trims
for c in ['Documento compras','Material','Posición','Documento material',
          'Tipo de historial de pedido','Centro','Indicador Debe/Haber']:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip()

# Fecha
if 'Fe.contabilización' in df.columns:
    df['Fe.contabilización'] = pd.to_datetime(df['Fe.contabilización'], errors='coerce', dayfirst=True)

# Cantidad numérica
df['Cantidad_num'] = pd.to_numeric(df['Cantidad'], errors='coerce')

# Normaliza Posición (evita 00010 vs 10)
if 'Posición' in df.columns:
    pos_num = pd.to_numeric(df['Posición'], errors='coerce')
    df['Posición'] = np.where(pos_num.notna(), pos_num.astype('Int64').astype(str),
                              df['Posición'].astype(str))

# ============== 2) Trabajamos SOLO con Q ==============
tipo = df['Tipo de historial de pedido'].str.upper()
q = df.loc[tipo.eq('Q')].copy()
otros = df.loc[~tipo.eq('Q')].copy()    # por si luego quieres recombinar

# ============== 3) Clave AUX ==============
key_parts = ['Documento compras','Material']
if USE_POSICION and 'Posición' in q.columns:
    key_parts.append('Posición')
if ADD_CENTRO_TO_KEY and 'Centro' in q.columns:
    key_parts.append('Centro')

# construye AUX con seguridad
for k in key_parts:
    q[k] = q[k].astype(str).str.strip().str.upper() if k == 'Material' else q[k].astype(str).str.strip()
q['AUX'] = q[key_parts].agg('|'.join, axis=1)

# ============== 4) Orden estable: S (factura) antes que H (anulación) ==============
# Usamos Indicador Debe/Haber si existe; si no, fallback al signo de la cantidad.
if 'Indicador Debe/Haber' in q.columns:
    q['_orden_anul'] = (q['Indicador Debe/Haber'].astype(str).str.upper() == 'H').astype(int)
else:
    q['_orden_anul'] = (q['Cantidad_num'] < 0).astype(int)   # 0=factura, 1=anulacion

# Parsear Fecha de documento si existe
if 'Fecha de documento' in q.columns:
    q['_fecha_doc'] = pd.to_datetime(q['Fecha de documento'], errors='coerce', dayfirst=True)
else:
    q['_fecha_doc'] = pd.NaT

# Orden: primero signo (positivos/facturas antes que negativos/anulaciones),
# luego fecha. Asi todas las facturas entran al stack antes de que llegue
# cualquier anulacion, independientemente de sus fechas de contabilizacion.
by = ['AUX', '_orden_anul', '_fecha_doc']
if 'Fe.contabilización' in q.columns:
    by.append('Fe.contabilización')
if 'Documento material' in q.columns:
    by.append('Documento material')

q = (q.sort_values(by=by, ascending=[True]*len(by), na_position='last', kind='mergesort')
       .reset_index(drop=True))



# ============== 5) Empareje LIFO por monto dentro de cada AUX ==============
# Regla:
#   - Indicador 'S' (o cantidad > 0 si no hay indicador) -> push al stack
#   - Indicador 'H' (o cantidad < 0)                     -> pop LIFO y eliminar ambos
#   - Si una H no encuentra par, NO se elimina (se conserva por trazabilidad)
to_drop = set()
has_indicador = 'Indicador Debe/Haber' in q.columns

for aux, g in q.groupby('AUX', sort=False):
    stacks = defaultdict(list)  # clave_monto -> stack de índices de facturas (LIFO)
    for i in g.index:
        m = q.at[i, 'Cantidad_num']
        if pd.isna(m) or m == 0:
            continue
        k = round(abs(float(m)), AMT_DECIMALS)

        if has_indicador:
            ind = str(q.at[i, 'Indicador Debe/Haber']).upper()
            es_factura  = (ind == 'S')
            es_anulacion = (ind == 'H')
        else:
            es_factura  = (m > 0)
            es_anulacion = (m < 0)

        if es_factura:
            stacks[k].append(i)            # push factura
        elif es_anulacion:
            if stacks[k]:
                j = stacks[k].pop()        # pop factura más reciente (LIFO)
                to_drop.update([i, j])
            # si no hay match, dejamos la anulación tal cual

# Aplica eliminación de pares
q_dep = q.drop(index=list(to_drop)).reset_index(drop=True)

# ============== 6) Resultado ==============
# Solo FACTURAS (Q) depuradas: ésta es la que usarás para tus cruces/agrupaciones
me80fn_q = q_dep.drop(columns=['_orden_anul','AUX'], errors='ignore').copy()

# (Opcional) Con otros tipos recombinados
me80fn_full = (pd.concat([me80fn_q, otros], ignore_index=True)
                 .sort_values(
                     by=[c for c in ['Documento compras','Material','Posición',
                                     'Fe.contabilización','Documento material']
                         if c in me80fn_q.columns or c in otros.columns],
                     kind='mergesort')
                 .reset_index(drop=True))

print("✅ Q depuradas (me80fn_q):", me80fn_q.shape)
print("ℹ️  Con otros tipos (me80fn_full):", me80fn_full.shape)

# ===== Chequeo opcional de duplicados en Q por clave y monto =====
chk = (me80fn_q
       .assign(monto_key = me80fn_q['Cantidad_num'].abs().round(AMT_DECIMALS))
       .groupby(key_parts + ['monto_key'], dropna=False)
       .size().reset_index(name='n').query('n > 1'))
print("🔎 Duplicados Q restantes (mismo AUX y monto):", chk.shape[0])
# print(chk.head(10))


# %%
me80fn = me80fn_q.copy()

# %%
#exportar para revision 

me80fn.to_excel(ruta_base / "me80fn_anu.xlsx", index=False)
print(f"✅ Archivo exportado en: {ruta_base / 'me80fn_anu.xlsx'}")

# %%


# ✅ Filtrar tipo 'Q' y cantidad positiva
me80fn = me80fn[
    (me80fn['Tipo de historial de pedido'].str.strip() == 'Q') &
    (me80fn['Cantidad'] > 0)
]

# 🔄 Cargar archivo ME2N
me2n = pd.read_excel(ruta / 'ME2N.XLSX', 
                    dtype=str, 
                    engine='openpyxl', 
                    skiprows=[1])

# %%
# 📋 Diagnóstico: Ver columnas disponibles antes de seleccionar
print("\n📋 Columnas disponibles en ME2N:")
for i, col in enumerate(me2n.columns):
    print(f"{i}: {col}")

# 🛡️ Verificación segura de columna 'Material'
if 'Material' in me2n.columns:
    col_material_correcta = 'Material'
else:
    raise ValueError("❌ No se encontró la columna 'Material' en ME2N. Revisa el layout exportado.")



# %%
# 📋 Definir columnas esperadas
columnas_deseadas = [
    'Documento compras',
    'Cl.documento compras',
    'Por entregar (cantidad)',
    'Cantidad de pedido',
    'Posición',
    col_material_correcta,
    'Fecha de entrega',
    'Fecha documento'
]

# %%
# 🧹 Filtrar solo columnas existentes (evita error si falta alguna)
me2n = me2n[[col for col in columnas_deseadas if col in me2n.columns]]

# ✅ Asegurar nombre estándar
if col_material_correcta != 'Material':
    me2n = me2n.rename(columns={col_material_correcta: 'Material'})


# Eliminar filas donde 'Documento compras' está vacío o nulo
me2n = me2n[me2n['Documento compras'].notna() & (me2n['Documento compras'].str.strip() != '')]

print("🗑️ Eliminando clases de pedido no deseadas en ME2N...")
clases_a_eliminar = ['Z250', 'Z210', 'Z520', 'Z420']
me2n = me2n[~me2n['Cl.documento compras'].isin(clases_a_eliminar)]

# Extraer órdenes válidas después de eliminar la fila y las clases no deseadas
ocs_validas = me2n['Documento compras'].dropna().drop_duplicates()
ocs_validas = ocs_validas[ocs_validas.str.strip() != '']

# Guardar en archivo temporal para carga en SAP
archivo_tmp = ruta_base / "ordenes_tmp.txt"
ocs_validas.to_csv(archivo_tmp, index=False, header=False)

# Copiar al portapapeles (opcional si haces el copy-paste manual)
ocs_validas.to_clipboard(index=False, header=False)
print("📋 OCs copiadas al portapapeles.")


# %%
print("🔧 Limpiando columnas de ME2N...")
col_material_correcta = me2n.columns[5]  # Columna F
me2n = me2n[['Documento compras', 'Cl.documento compras', 'Por entregar (cantidad)', 'Cantidad de pedido', 'Posición', col_material_correcta, 'Fecha de entrega', 'Fecha documento']]
me2n = me2n.rename(columns={col_material_correcta: 'Material'})

print("🗑️ Eliminando clases de pedido no deseadas en ME2N...")
clases_a_eliminar = ['Z250', 'Z210', 'Z520', 'Z420']
me2n = me2n[~me2n['Cl.documento compras'].isin(clases_a_eliminar)]

print("🚛 Agregando columna 'VIA' según 'Cl.documento compras'...")
mapeo_via = {
    'Z300': 'Aereo',
    'Z290': 'Aereo',
    'Z280': 'Aereo',
    'Z270': 'Maritimo',
    'Z260': 'Aereo',
    'Z241': 'Aereo',
    'Z240': 'Aereo',
    'Z220': 'Maritimo',
    'Z200': 'Maritimo',
    'Z100': 'Nacional'
}

# %%
me2n['VIA'] = me2n['Cl.documento compras'].map(mapeo_via).fillna('')
me80fn = me80fn[['Documento compras', 'Posición', 'Cantidad', 'Material',  'Tipo de historial de pedido', 'Fe.contabilización']]
print("🧹 Filtrando ME80FN...")

# %%
print("🛠️ Creando claves AUX en ME2N y ME80FN...")
me80fn['AUX'] = me80fn['Documento compras'].str.strip() + me80fn['Material'].str.strip() + me80fn['Posición'].str.strip()
me2n['AUX'] = me2n['Documento compras'].str.strip() + me2n['Material'].str.strip() + me2n['Posición'].str.strip()

print("🔄 Realizando merge ME2N con ME80FN...")
resultado = pd.merge(
    me2n,
    me80fn[['AUX', 'Cantidad', 'Tipo de historial de pedido', 'Fe.contabilización']],
    on='AUX',
    how='left'
)


# %%
from pathlib import Path
import pandas as pd
import numpy as np
import re

# 1) Localizar y cargar
posibles = [ruta_base / "vl06i.XLSX", ruta_base / "VL06I.xlsx", ruta_base / "VL06IF.xlsx"]
vl_path = next((p for p in posibles if p.exists()), None)
if vl_path is None:
    raise FileNotFoundError("❌ No encontré vl06i.xlsx / VL06I.xlsx / VL06IF.xlsx en la carpeta seleccionada.")

vl = pd.read_excel(vl_path, dtype=str, engine="openpyxl")
print("📄 VL06I cargado:", vl_path.name, "| filas:", len(vl))
print("🔤 Columnas VL06I:", list(vl.columns)[:12], "...")

# 2) Detectar la columna de fecha
col_fecha = None
for c in vl.columns:
    if c.strip().lower() in {"entregas (de/a)","entrega (de/a)","entregas de/a"}:
        col_fecha = c
        break
if col_fecha is None:
    posibles_fechas = [c for c in vl.columns if "entreg" in c.lower()]
    raise KeyError(f"❌ No encontré 'Entregas (de/a)'. Detectadas: {posibles_fechas}")

print("📌 Columna de fecha en VL06I:", col_fecha)
print("👀 Muestra de valores de fecha:", vl[col_fecha].dropna().head(10).to_list())

# %%
# Asegurarse de que la columna 'Entregas (de/a)' sea datetime
vl['Entregas (de/a)'] = pd.to_datetime(vl['Entregas (de/a)'], errors='coerce')



# %%
vl.rename(columns={
    'Documento compras': 'OC_VL06',
    'Entregas (de/a)': 'FEC.EST.NAVEG.'
}, inplace=True)                   # ✅


# %%
import pandas as pd
import numpy as np

# ========= 0) Normalización mínima y chequeos =========
resultado = resultado.copy()
vl = vl.copy()

resultado.columns = resultado.columns.str.strip()
vl.columns = vl.columns.str.strip()

req_res = {'Documento compras','Material','Cantidad','Tipo de historial de pedido'}
req_vl  = {'OC_VL06','Material','Cantidad entrega','FEC.EST.NAVEG.'}
miss_res = req_res - set(resultado.columns)
miss_vl  = req_vl  - set(vl.columns)
if miss_res or miss_vl:
    raise KeyError(f"Faltan columnas -> resultado:{miss_res}  vl:{miss_vl}")

# ========= 1) Limpieza de claves y tipos =========
resultado['Documento compras'] = resultado['Documento compras'].astype(str).str.strip()
vl['OC_VL06']                  = vl['OC_VL06'].astype(str).str.strip()

resultado['Material'] = resultado['Material'].astype(str).str.upper().str.strip()
vl['Material']        = vl['Material'].astype(str).str.upper().str.strip()

to_num = lambda s: pd.to_numeric(s, errors='coerce')

resultado['Cantidad']        = to_num(resultado['Cantidad'])
vl['Cantidad entrega']       = to_num(vl['Cantidad entrega'])

vl['FEC.EST.NAVEG.'] = pd.to_datetime(vl['FEC.EST.NAVEG.'], errors='coerce', dayfirst=True)
if 'Fe.contabilización' in resultado.columns:
    resultado['Fe.contabilización'] = pd.to_datetime(resultado['Fe.contabilización'], errors='coerce', dayfirst=True)

# ========= 2) Facturados (regla: contiene 'Q' en Tipo de historial de pedido) =========
mask_fact = resultado['Tipo de historial de pedido'].astype(str).str.upper().str.contains('Q', na=False)
fact = resultado.loc[mask_fact].copy()

# Entregas con nombre de OC coherente
ent = vl.rename(columns={'OC_VL06': 'Documento compras'}).copy()

# ========= 2.1) Centro opcional como parte de la llave =========
keys = ['Documento compras', 'Material']
if 'Centro' in ent.columns and 'Centro' in fact.columns:
    fact['Centro'] = fact['Centro'].astype(str).str.strip()
    ent['Centro']  = ent['Centro'].astype(str).str.strip()
    keys.append('Centro')

# ========= 3) Cantidades como unidades enteras para expansión =========
# Si manejas decimales, cambia factor=1000 y multiplica antes de .round()
factor = 1
fact['_qty_u'] = (fact['Cantidad'].fillna(0) * factor).round().astype(int).clip(lower=0)
ent['_qty_u']  = (ent['Cantidad entrega'].fillna(0) * factor).round().astype(int).clip(lower=0)

# ========= 4) Orden determinístico (estable) =========
# Facturas: por llaves (+ fecha contab si existe)
if 'Fe.contabilización' in fact.columns:
    fact = fact.sort_values(keys + ['Fe.contabilización'], kind='mergesort')
else:
    fact = fact.sort_values(keys, kind='mergesort')

# Entregas: por llaves + fecha DESC (más reciente primero)
bys = keys + ['FEC.EST.NAVEG.']
asc = [True]*len(keys) + [False]  # fecha descendente

# Desempate opcional por 'Entrega' DESC si existe (para misma fecha)
if 'Entrega' in ent.columns:
    bys = bys + ['Entrega']
    asc = asc + [False]

ent = ent.sort_values(bys, ascending=asc, kind='mergesort')

# ========= 5) Expansión por unidades =========
fact['_row_id'] = fact.index
ent['_ent_id']  = ent.index

fact_rep = fact.loc[fact.index.repeat(fact['_qty_u'])].copy()
ent_rep  = ent.loc[ent.index.repeat(ent['_qty_u'])].copy()

if len(fact_rep) == 0 or len(ent_rep) == 0:
    resultado['FEC_EST_NAVEG_MATCH'] = pd.NaT
else:
    # Slots por llaves
    fact_rep['_slot'] = fact_rep.groupby(keys).cumcount()
    ent_rep['_slot']  = ent_rep.groupby(keys).cumcount()

    # Merge 1–a–1 por (llaves + slot)
    on_cols = keys + ['_slot']
    match = fact_rep.merge(
        ent_rep[on_cols + ['FEC.EST.NAVEG.']],
        on=on_cols,
        how='left',
        validate='one_to_one'
    )

    # Colapsa a la fila original de factura (si hay >1 unidad, toma la más reciente de las asignadas)
    agg_dates = (match.groupby('_row_id', as_index=False)['FEC.EST.NAVEG.']
                      .agg(FEC_EST_NAVEG_MATCH='max'))

    resultado = resultado.join(agg_dates.set_index('_row_id'), how='left')

# (Opcional) diagnóstico
# print('Facturas con fecha asignada:', resultado['FEC_EST_NAVEG_MATCH'].notna().sum(), '/', len(resultado))


# %%

#exportar para revision 

resultado.to_excel(ruta_base / "revfecha.xlsx", index=False)
print(f"✅ Archivo exportado en: {ruta_base / 'revfecha.xlsx'}")



# %%
print("🔧 Generando columna 'estatus ft'...")
resultado['estatus ft'] = resultado['Cantidad'].apply(lambda x: 'Facturado' if pd.notna(x) else 'No Facturado')

# %%
# Calcular días entre fecha de factura y Fecha de entrega
import numpy as np

print("📆 Calculando días en factura y creando 'status2'...")

# Asegurar formato de fechas
resultado['Fe.contabilización'] = pd.to_datetime(resultado['Fe.contabilización'], errors='coerce')
resultado['Fecha documento'] = pd.to_datetime(resultado['Fecha documento'], errors='coerce')

# Calcular diferencia de días
resultado['dias_en_factura'] = (resultado['Fe.contabilización'] - resultado['Fecha documento']).dt.days

# Aplicar condiciones según VIA y días 
conditions = [
    (resultado['estatus ft'] == 'Facturado') & (resultado['VIA'] == 'Maritimo') & (resultado['dias_en_factura'] <= 60),
    (resultado['estatus ft'] == 'Facturado') & (resultado['VIA'] == 'Maritimo') & (resultado['dias_en_factura'] > 60),
    (resultado['estatus ft'] == 'Facturado') & (resultado['VIA'] == 'Aereo') & (resultado['dias_en_factura'] <= 30),
    (resultado['estatus ft'] == 'Facturado') & (resultado['VIA'] == 'Aereo') & (resultado['dias_en_factura'] > 30),
]

choices = [
    'Facturado No Vencido',
    'Facturado Vencido',
    'Facturado No Vencido',
    'Facturado Vencido'
]

resultado['status2'] = np.select(conditions, choices, default='No Facturado')

# %%
hoy = pd.to_datetime(datetime.today().strftime('%Y-%m-%d'))

# %%
resultado.rename(columns={
    'FEC_EST_NAVEG_MATCH': 'FEC.EST.NAVEG.'
}, inplace=True)                   



# %%
# Asegurarse de que las fechas estén en formato datetime
resultado['Fecha documento'] = pd.to_datetime(resultado['Fecha documento'], errors='coerce')
resultado['' \
''] = pd.to_datetime(resultado['FEC.EST.NAVEG.'], errors='coerce')


# Calcular Lead Time objetivo según la vía
resultado['LT Objetivo'] = np.where(resultado['VIA'].isin(["Maritimo", "Terrestre"]), 45,
                        np.where(resultado['VIA'] == "Aereo", 15,
                        np.where(resultado['VIA'] == "Courier", 7, 0)))

# Calcular días de desfase respecto al objetivo
resultado['Dias LT BO'] = (hoy - resultado['Fecha documento']).dt.days - resultado['LT Objetivo']
resultado['Dias LT BO'] = resultado['Dias LT BO'].fillna(0)

# %%
# Categorización de LT por rangos
conditions = [
    (resultado['VIA'].isin(["Maritimo", "Terrestre"])) & (resultado['Dias LT BO'].between(1, 30)),
    (resultado['VIA'].isin(["Maritimo", "Terrestre"])) & (resultado['Dias LT BO'].between(31, 60)),
    (resultado['VIA'].isin(["Maritimo", "Terrestre"])) & (resultado['Dias LT BO'].between(61, 90)),
    (resultado['VIA'].isin(["Maritimo", "Terrestre"])) & (resultado['Dias LT BO'] > 90),
    (resultado['VIA'] == "Aereo") & (resultado['Dias LT BO'].between(1, 15)),
    (resultado['VIA'] == "Aereo") & (resultado['Dias LT BO'].between(16, 30)),
    (resultado['VIA'] == "Aereo") & (resultado['Dias LT BO'].between(31, 45)),
    (resultado['VIA'] == "Courier") & (resultado['Dias LT BO'].between(1, 7)),
    (resultado['VIA'] == "Courier") & (resultado['Dias LT BO'].between(8, 15)),
    (resultado['VIA'] == "Courier") & (resultado['Dias LT BO'].between(16, 21)),
    (resultado['VIA'] == "Courier") & (resultado['Dias LT BO'] > 21),
    (resultado['VIA'] == "Aereo") & (resultado['Dias LT BO'] > 45)
]
choices = [
    "De 1 a 30 días",
    "De 31 a 60 días",
    "De 61 a 90 días",
    "De 91 días a más",
    "De 1 a 15 días",
    "De 16 a 30 días",
    "De 31 a 45 días",
    "De 1 a 7 días",
    "De 8 a 15 días",
    "De 16 a 21 días",
    "De 21 días a más",
    "De 45 días a más"
]
resultado['Categoría LT'] = np.select(conditions, choices, default="Dentro de LT")

# %%
# UF .... ME COSTO

from datetime import datetime
import numpy as np
import pandas as pd

# Asegura tipo fecha
resultado['Fecha de entrega'] = pd.to_datetime(resultado['Fecha de entrega'], errors='coerce')

# Fecha actual como datetime (sin hora)
hoy = pd.to_datetime(datetime.now().date())

# === Calcular status3 para No Facturado basado en Fecha de entrega ===
limite = hoy + pd.Timedelta(days=31)

cond_teorica = (resultado['Fecha de entrega'] > limite)
cond_replanificar = resultado['Fecha de entrega'].between(hoy, limite, inclusive='both')
cond_tr_vencido = (resultado['Fecha de entrega'] < hoy)

resultado['status2'] = np.select(
    [
        (resultado['estatus ft'] == 'No Facturado') & cond_teorica,
        (resultado['estatus ft'] == 'No Facturado') & cond_replanificar,
        (resultado['estatus ft'] == 'No Facturado') & cond_tr_vencido
    ],
    
    [
        'Fecha Teórica',
        'Replanificar+45',
        'TR Vencido'
    ],
    default= resultado['status2'] 
)



# %%
print("🔧 Renombrando columna 'Cantidad' a 'Cantidad Facturada'...")
resultado = resultado.rename(columns={'Cantidad': 'Cantidad Facturada'})

# %%
import pandas as pd
import numpy as np

llave      = ['Documento compras', 'Material', 'Posición']
col_pedido = 'Cantidad de pedido'
col_fact   = 'Cantidad Facturada'
col_pe     = 'Por entregar (cantidad)'  # existe en tu base

# --- Tipos numéricos básicos
for c in [col_pedido, col_fact, col_pe]:
    if c in resultado.columns:
        resultado[c] = pd.to_numeric(resultado[c], errors='coerce').fillna(0)

# --- Estatus normalizado
estatus = resultado['estatus ft'].astype(str).str.upper().str.strip()
es_fact = estatus.eq('FACTURADO')
es_no_f = estatus.eq('NO FACTURADO')

# --- Métricas por grupo
fact_sum   = resultado.groupby(llave)[col_fact].transform('sum')
pedido_ref = resultado.groupby(llave)[col_pedido].transform('first')
pendiente  = (pedido_ref - fact_sum).clip(lower=0)

# --- Detectar grupos SIN NO FACTURADO y con pendiente > 0
tiene_nf = resultado.groupby(llave)['estatus ft'] \
    .transform(lambda s: (s.astype(str).str.upper().str.strip() == 'NO FACTURADO').any())

falta_nf_y_pend = (~tiene_nf) & (pendiente > 0)

# --- Construir DF de claves a crear
claves_crear = resultado.loc[falta_nf_y_pend, llave].drop_duplicates()

# --- Crear filas sintéticas NO FACTURADO por cada clave faltante
filas_nuevas = []
for _, key in claves_crear.iterrows():
    mask_grp = (resultado[llave] == key.values).all(axis=1)
    fila_base = resultado.loc[mask_grp].iloc[0].copy()

    fila_base['estatus ft']     = 'NO FACTURADO'
    fila_base[col_fact]         = 0
    if col_pe in resultado.columns:
        fila_base[col_pe]       = 0   # opcional: podrías poner el pendiente del grupo
    fila_base[col_pedido]       = 0   # evita inflar sumas ingenuas
    fila_base['fila_sintetica'] = True

    filas_nuevas.append(fila_base)

if filas_nuevas:
    resultado = pd.concat([resultado, pd.DataFrame(filas_nuevas)], ignore_index=True)

# --- Recalcular estatus / métricas tras insertar
estatus = resultado['estatus ft'].astype(str).str.upper().str.strip()
es_fact = estatus.eq('FACTURADO')
es_no_f = estatus.eq('NO FACTURADO')

fact_sum   = resultado.groupby(llave)[col_fact].transform('sum')
pedido_ref = resultado.groupby(llave)[col_pedido].transform('first')
pendiente  = (pedido_ref - fact_sum).clip(lower=0)

# --- Marcar 1ª NO FACTURADO de cada grupo
resultado['_nf'] = es_no_f
first_no_fact = es_no_f & (resultado.groupby(llave)['_nf'].cumsum() == 1)

# --- Construir Cantidad real (sin doble conteo)
resultado['Cantidad real'] = 0.0
resultado.loc[es_fact, 'Cantidad real'] = resultado.loc[es_fact, col_fact]
resultado.loc[first_no_fact, 'Cantidad real'] = pendiente.loc[first_no_fact]

# Limpieza
resultado.drop(columns=['_nf'], inplace=True, errors='ignore')




# %%
import pandas as pd
import numpy as np

llave      = ['Documento compras', 'Material', 'Posición']
col_pedido = 'Cantidad de pedido'
col_fact   = 'Cantidad Facturada'
col_pe     = 'Por entregar (cantidad)'

# 1) Tipos numéricos
for c in [col_pedido, col_fact, col_pe]:
    if c in resultado.columns:
        resultado[c] = pd.to_numeric(resultado[c], errors='coerce').fillna(0)

# 2) Estatus
estatus = resultado['estatus ft'].astype(str).str.upper().str.strip()
es_fact = estatus.eq('FACTURADO')
es_no_f = estatus.eq('NO FACTURADO')

# 3) Pendiente por grupo
fact_sum   = resultado.groupby(llave)[col_fact].transform('sum')
pedido_ref = resultado.groupby(llave)[col_pedido].transform('first')
pendiente  = (pedido_ref - fact_sum).clip(lower=0)

# 4) Detectar grupos SIN NO FACTURADO y con pendiente > 0
tiene_nf = es_no_f.groupby([resultado[c] for c in llave]).transform('any')
claves_crear = resultado.loc[(~tiene_nf) & (pendiente > 0), llave].drop_duplicates()

# 5) Crear filas sintéticas NO FACTURADO por clave faltante
filas_nuevas = []
for _, key in claves_crear.iterrows():
    mask_grp = (resultado[llave] == key.values).all(axis=1)
    fila_base = resultado.loc[mask_grp].iloc[0].copy()

    fila_base['estatus ft']     = 'NO FACTURADO'
    fila_base[col_fact]         = 0
    fila_base[col_pedido]       = 0         # evita inflar sumas ingenuas
    if col_pe in resultado.columns:
        fila_base[col_pe]       = 0         # se usará solo como ponderador
    fila_base['fila_sintetica'] = True

    filas_nuevas.append(fila_base)

if filas_nuevas:
    resultado = pd.concat([resultado, pd.DataFrame(filas_nuevas)], ignore_index=True)

# 6) Recalcular métricas tras insertar
estatus = resultado['estatus ft'].astype(str).str.upper().str.strip()
es_fact = estatus.eq('FACTURADO')
es_no_f = estatus.eq('NO FACTURADO')

fact_sum   = resultado.groupby(llave)[col_fact].transform('sum')
pedido_ref = resultado.groupby(llave)[col_pedido].transform('first')
pendiente  = (pedido_ref - fact_sum).clip(lower=0)

# === Construir Cantidad real con reparto robusto ===
# 1) FACTURADO: lo facturado de la fila
resultado['Cantidad real'] = 0.0
resultado.loc[es_fact, 'Cantidad real'] = resultado.loc[es_fact, col_fact]

# 2) Pesos NO FACTURADO
resultado['_peso_no_f'] = np.where(es_no_f, resultado[col_pe], 0.0)

# Suma de pesos por grupo (solo NO FACTURADO)
peso_sum = resultado.groupby(llave)['_peso_no_f'].transform('sum')

# ---- FIX: máscara de filas sintéticas (por-fila, no por-columnas)
mask_sint = resultado['fila_sintetica'] if 'fila_sintetica' in resultado.columns else pd.Series(False, index=resultado.index)

# Si el grupo no tiene base de reparto (peso_sum == 0) y hay fila_sintetica,
# forzar que esa fila reciba todo el pendiente.
sin_base = es_no_f & peso_sum.eq(0)
resultado.loc[sin_base & mask_sint, '_peso_no_f'] = 1.0

# (Opcional, solo visual): refleja el pendiente en Por Entregar de la fila creada
# para que en Excel veas el "40" también allí.
if col_pe in resultado.columns:
    resultado.loc[sin_base & mask_sint, col_pe] = pendiente.loc[sin_base & mask_sint]

# Recalcular suma de pesos por grupo tras el ajuste
peso_sum = resultado.groupby(llave)['_peso_no_f'].transform('sum')

# 3) Reparto proporcional del 'pendiente' SOLO entre NO FACTURADO
prop = np.divide(
    np.where(es_no_f, resultado['_peso_no_f'], 0.0),
    peso_sum.where(peso_sum > 0, 1.0),
    out=np.zeros_like(resultado['_peso_no_f'], dtype=float),
    where=(peso_sum.values > 0)
)

resultado['Cantidad real'] += (prop * pendiente)

# Limpieza
resultado.drop(columns=['_peso_no_f'], inplace=True, errors='ignore')



# %%
mask_nf = resultado['estatus ft'].astype(str).str.upper().str.strip().eq('NO FACTURADO')
resultado.loc[mask_nf, 'status2'] = 'No Facturado'

# %%
# UF .... ME COSTO 2

from datetime import datetime
import numpy as np
import pandas as pd

# Asegura tipo fecha
resultado['Fecha de entrega'] = pd.to_datetime(resultado['Fecha de entrega'], errors='coerce')

# Fecha actual (sin hora) y límite fijo de 31 días
hoy    = pd.Timestamp.today().normalize()
limite = hoy + pd.Timedelta(days=31)

# Solo filas donde status2 == "No Facturado" (robusto a mayúsculas/espacios)
mask_nf2 = resultado['status2'].astype(str).str.strip().str.upper().eq('NO FACTURADO')

# Condiciones de fecha (se aplican SOLO sobre esas filas)
cond_teorica      = mask_nf2 & (resultado['Fecha de entrega'] > limite)
cond_replanificar = mask_nf2 & resultado['Fecha de entrega'].between(hoy, limite, inclusive='both')
cond_tr_vencido   = mask_nf2 & (resultado['Fecha de entrega'] < hoy)

# Actualiza status2; si la fecha es NaT se mantiene "No Facturado"
resultado['status2'] = np.select(
    [cond_teorica,            cond_replanificar,        cond_tr_vencido],
    ['Fecha Teórica',         'Replanificar+45',        'TR Vencido'],
    default=resultado['status2']
)


# %%
print("📅 Calculando 'Fecha final'...")

# Inicializar columna
resultado['Fecha final'] = pd.NaT

# %%
print("🔧 Renombrando columna 'status2' a 'estado_final'...")
resultado = resultado.rename(columns={'status2': 'estado_final'})

# %%


# 1. Facturado:  Fe.contabilización + LT
cond_facturado = resultado['estado_final'].isin(['Facturado Vencido', 'Facturado No Vencido'])


resultado.loc[
    cond_facturado & resultado['FEC.EST.NAVEG.'].notna(),
    'Fecha final'
] = resultado['FEC.EST.NAVEG.']

resultado.loc[
    cond_facturado & resultado['FEC.EST.NAVEG.'].isna(),
    'Fecha final'
] = 'revisar'



# %%


# 1) Condición para Fecha Teórica
cond_teorica = resultado['estado_final'].eq('Fecha Teórica')

# 2) Asignar siempre 'Fecha de entrega' a 'Fecha final' cuando sea Fecha Teórica
resultado.loc[cond_teorica, 'Fecha final'] = resultado['Fecha de entrega']



# %%

# 3. Otros estados: lógica anterior
cond_otros = resultado['Fecha final'].isna()

# Subconjunto temporal de esas filas
otros_df = resultado.loc[cond_otros].copy()


# Aqui va la regla, agregue mas, y ojo con aereo

# Crear condiciones SOLO para esas filas
cond1 = otros_df['estado_final'] == "Replanificar+45"
cond2 = (otros_df['estado_final'] == "TR Vencido") & ((otros_df['Categoría LT'] == 0) | (otros_df['Categoría LT'] == "Dentro de LT")) # Para aereo y maritimo que NO estan atrasados pero esta bien segun su categoria 
cond3 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 31 a 60 días") # Aqui cae el maritimo y terrestre 
cond4 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 61 a 90 días")  # Aqui cae el maritimo y terrestre 
cond5 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 91 días a más") # Aqui cae el maritimo y terrestre 
cond6 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 31 a 45 días") # Aqui cae el Aereo  
cond7 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 45 días a más") # Aqui cae el Aereo  
cond8 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 1 a 30 días") #  Aqui cae el maritimo y terrestre  
cond9 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 1 a 15 días") # Aqui cae el Aereo  
cond10 = (otros_df['estado_final'] == "TR Vencido") & (otros_df['Categoría LT'] == "De 16 a 30 días") # Aqui cae el Aereo  


# Aplicar np.select sobre ese subset
otros_fechas = np.select(
    [cond1, cond2, cond3, cond4, cond5, cond6, cond7, cond8, cond9, cond10],
    [
        hoy + timedelta(days=45), # "Replanificar+45"
        hoy + timedelta(days=30), # "Dentro de LT"
        hoy + timedelta(days=60), # Maritimo y terrestre
        hoy + timedelta(days=90), # Maritimo y terrestre
        hoy + timedelta(days=120), # Maritimo y terrestre
        hoy + timedelta(days=60), # Aereo 
        hoy + timedelta(days=120), # Aereo   
        hoy + timedelta(days=30),  # Maritimo y terrestre  
        hoy + timedelta(days=30), # Aereo
        hoy + timedelta(days=30) # Aereo
    ], 
    default=hoy + timedelta(days=120)
)

# Asignar resultados al DataFrame original
resultado.loc[cond_otros, 'Fecha final'] = otros_fechas



# %%


# 🧼 Traer columnas adicionales desde ME2N original (para merge)
me2n_extra_cols = pd.read_excel(
    archivo_me2n,
    dtype=str,
    engine='openpyxl',
    skiprows=[1],
    usecols=['Documento compras', 'Posición', 'Material', 'Texto breve', 'Centro', 'Nombre del proveedor', 'Grupo de compras']
)

# %%
# Eliminar duplicados por clave para no traer repetidos
me2n_extra_cols = me2n_extra_cols.drop_duplicates(subset=['Documento compras', 'Posición', 'Material'])

# Crear clave de unión
me2n_extra_cols['AUX'] = me2n_extra_cols['Documento compras'].str.strip() + me2n_extra_cols['Material'].str.strip() + me2n_extra_cols['Posición'].str.strip()
resultado['AUX'] = resultado['Documento compras'].str.strip() + resultado['Material'].str.strip() + resultado['Posición'].str.strip()

# Hacer el merge con los datos extra
resultado = resultado.merge(
    me2n_extra_cols[['AUX', 'Texto breve', 'Centro', 'Nombre del proveedor', 'Grupo de compras']],
    on='AUX',
    how='left'
)

# %%
# 🧹 Eliminar filas con Documento compras vacío por seguridad
resultado = resultado[resultado['Documento compras'].notna() & (resultado['Documento compras'].str.strip() != '')]


# 🧹 Eliminar materiales específicos antes de guardar
materiales_a_excluir = ['100042753', '100045025', '100045026', '100045522']
resultado = resultado[~resultado['Material'].isin(materiales_a_excluir)]
print(f"🧽 Filtrado de materiales específicos completado. Total registros ahora: {resultado.shape[0]}")


# %%

# 💾 Guardar archivo 
resultado.to_excel(archivo_salida, index=False)

# %%

# 🧾 Dejar columnas en el orden solicitado
columnas_finales = [
    
    'Documento compras',
    'estado_final',
    'Material',
    'Texto breve',
    'Cantidad real',
    'Centro',
    'Cl.documento compras',
    'Fecha final',
    'Nombre del proveedor',
    'Grupo de compras',
    'Posición',

]

resultado = resultado[columnas_finales]




# %%
# === RESUMEN MOVIMIENTOS 103 ===
print("📦 Generando resumen de movimientos con clase 103...")

# Cargar archivo original ME80FN si no se ha hecho aún
me80fn_completo = pd.read_excel(archivo_me80fn, dtype=str, engine='openpyxl')

# Asegurarnos de que las columnas necesarias existen
columnas_necesarias = ['Documento compras', 'Material', 'Posición', 'Clase de movimiento', 'St.bloq.EM UMP', 'Tipo de historial de pedido']
faltantes = [col for col in columnas_necesarias if col not in me80fn_completo.columns]

if faltantes:
    print(f"❌ No se pueden generar los movimientos 103. Faltan columnas: {faltantes}")
else:
    # Filtrar clase de movimiento 103
    movimientos_103 = me80fn_completo[me80fn_completo['Clase de movimiento'].astype(str).str.strip() == '103'].copy()

    # Convertir 'St.bloq.EM UMP' a numérico
    movimientos_103['St.bloq.EM UMP'] = pd.to_numeric(movimientos_103['St.bloq.EM UMP'], errors='coerce').fillna(0)

    # ✅ Agrupar sin 'Tipo de historial de pedido'
    resumen_103 = (
        movimientos_103
        .groupby(['Documento compras', 'Material', 'Posición'], as_index=False)
        .agg({'St.bloq.EM UMP': 'sum'})
    )



    if resumen_103.empty:
        print("⚠️ No se encontraron movimientos 103 para exportar.")
    else:
        archivo_103 = ruta / 'Movimientos_103.xlsx'
        resumen_103.to_excel(archivo_103, index=False)
        print(f"✅ Archivo Movimientos_103.xlsx generado correctamente con {len(resumen_103)} registros.")


resultado.to_excel(archivo_salida, index=False)
print(f"✅ Archivo generado correctamente en: {archivo_salida}")

# %%
# === RESUMEN MOVIMIENTOS 103 MENOS 105 ===
print("📦 Generando resumen de movimientos 103 menos movimientos 105...")

# Cargar archivo original ME80FN
me80fn_completo = pd.read_excel(archivo_me80fn, dtype=str, engine='openpyxl')

# Columnas necesarias
columnas_necesarias = [
    'Documento compras',
    'Material',
    'Posición',
    'Clase de movimiento',
    'St.bloq.EM UMP',
    'Tipo de historial de pedido'
]

faltantes = [col for col in columnas_necesarias if col not in me80fn_completo.columns]

if faltantes:
    print(f"❌ No se pueden generar los movimientos 103/105. Faltan columnas: {faltantes}")

    # Para evitar error más abajo si resumen_103 no existe
    resumen_103 = pd.DataFrame(columns=[
        'Documento compras',
        'Material',
        'Posición',
        'St.bloq.EM UMP'
    ])

else:
    # =========================
    # 1) Normalizar columnas clave
    # =========================
    for c in ['Documento compras', 'Material', 'Posición', 'Clase de movimiento']:
        me80fn_completo[c] = me80fn_completo[c].astype(str).str.strip()

    # =========================
    # 2) Convertir St.bloq.EM UMP a número
    # =========================
    me80fn_completo['St.bloq.EM UMP'] = (
        me80fn_completo['St.bloq.EM UMP']
        .astype(str)
        .str.strip()
        .str.replace(r'\s+', '', regex=True)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False)
    )

    me80fn_completo['St.bloq.EM UMP'] = pd.to_numeric(
        me80fn_completo['St.bloq.EM UMP'],
        errors='coerce'
    ).fillna(0)

    # =========================
    # 3) Filtrar movimientos 103 y 105
    # =========================
    movimientos_103 = me80fn_completo[
        me80fn_completo['Clase de movimiento'].eq('103')
    ].copy()

    movimientos_105 = me80fn_completo[
        me80fn_completo['Clase de movimiento'].eq('105')
    ].copy()

    # Importante:
    # Dejamos las cantidades como positivas, porque el 105 se va a RESTAR.
    movimientos_103['St.bloq.EM UMP'] = movimientos_103['St.bloq.EM UMP'].abs()
    movimientos_105['St.bloq.EM UMP'] = movimientos_105['St.bloq.EM UMP'].abs()

    # =========================
    # 4) Agrupar 103
    # =========================
    resumen_103_base = (
        movimientos_103
        .groupby(['Documento compras', 'Material', 'Posición'], as_index=False)
        .agg({'St.bloq.EM UMP': 'sum'})
        .rename(columns={'St.bloq.EM UMP': 'Cantidad_103'})
    )

    # =========================
    # 5) Agrupar 105
    # =========================
    resumen_105 = (
        movimientos_105
        .groupby(['Documento compras', 'Material', 'Posición'], as_index=False)
        .agg({'St.bloq.EM UMP': 'sum'})
        .rename(columns={'St.bloq.EM UMP': 'Cantidad_105'})
    )

    # =========================
    # 6) Exportar detalle/resumen 105
    # =========================
    archivo_105 = ruta / 'Movimientos_105.xlsx'

    if resumen_105.empty:
        print("⚠️ No se encontraron movimientos 105 para exportar.")
    else:
        resumen_105.to_excel(archivo_105, index=False)
        print(f"✅ Archivo Movimientos_105.xlsx generado correctamente con {len(resumen_105)} registros.")

    # =========================
    # 7) Cruzar 103 con 105 y calcular neto
    # =========================
    resumen_103_neto = resumen_103_base.merge(
        resumen_105,
        on=['Documento compras', 'Material', 'Posición'],
        how='left'
    )

    resumen_103_neto['Cantidad_105'] = resumen_103_neto['Cantidad_105'].fillna(0)

    resumen_103_neto['Cantidad_103_neta'] = (
        resumen_103_neto['Cantidad_103'] - resumen_103_neto['Cantidad_105']
    )

    # Evitar cantidades negativas
    resumen_103_neto['Cantidad_103_neta'] = resumen_103_neto['Cantidad_103_neta'].clip(lower=0)

    # =========================
    # 8) Dejar resumen_103 listo para tu bloque proporcional actual
    # =========================
    # Mantener Cantidad_103 (original) y Cantidad_103_neta (neto tras 105)
    # para poder identificar exactamente qué fila eliminar/ajustar en resultado
    resumen_103 = resumen_103_neto[
        ['Documento compras', 'Material', 'Posición', 'Cantidad_103', 'Cantidad_103_neta']
    ].copy()

    # =========================
    # 9) Exportar respaldo del neto 103 - 105
    # =========================
    archivo_103_neto = ruta / 'Movimientos_103_menos_105.xlsx'
    resumen_103_neto.to_excel(archivo_103_neto, index=False)

    print(f"✅ Archivo Movimientos_103_menos_105.xlsx generado correctamente.")
    print(f"📊 Movimientos 103 encontrados: {len(resumen_103_base)}")
    print(f"📊 Movimientos 105 encontrados: {len(resumen_105)}")
    print(f"📊 Movimientos 103 netos a descontar: {len(resumen_103)}")

# Guardar resultado antes del ajuste proporcional
resultado.to_excel(archivo_salida, index=False)
print(f"✅ Archivo generado correctamente en: {archivo_salida}")

# %%
print("🧩 Columnas disponibles en resultado:", resultado.columns.tolist())

# %%
print("🔁 Aplicando lógica: Cantidad = Cantidad_103_Neta (consolidando en una fila por grupo)...")

# ============================================================
# 1) Normalización robusta de claves
# ============================================================
def _norm_key(s):
    """Convierte a str, limpia espacios, decimales tipo '51.0' y ceros a la izquierda."""
    s = s.astype(str).str.strip()
    # Quitar '.0' al final (ej: '51.0' → '51')
    s = s.str.replace(r'\.0+$', '', regex=True)
    # Quitar ceros a la izquierda solo si es numérico (ej: '00051' → '51')
    s = s.str.replace(r'^0+(?=\d)', '', regex=True)
    return s

for df in [resultado, resumen_103]:
    df['Documento compras'] = _norm_key(df['Documento compras'])
    df['Material']          = _norm_key(df['Material'])
    df['Posición']          = _norm_key(df['Posición'])

# ============================================================
# 2) Cantidad real a numérico
# ============================================================
resultado['Cantidad real'] = (
    pd.to_numeric(resultado['Cantidad real'], errors='coerce')
      .fillna(0)
      .clip(lower=0)
      .round(2)
)

# ============================================================
# 3) Detectar la columna neta en resumen_103
# ============================================================
resumen_103 = resumen_103.copy()

if 'Cantidad_103_neta' in resumen_103.columns:
    resumen_103 = resumen_103.rename(columns={'Cantidad_103_neta': 'Cantidad_103_Neta_Final'})
elif 'Cantidad_103_Neta' in resumen_103.columns:
    resumen_103 = resumen_103.rename(columns={'Cantidad_103_Neta': 'Cantidad_103_Neta_Final'})
elif 'St.bloq.EM UMP' in resumen_103.columns:
    resumen_103 = resumen_103.rename(columns={'St.bloq.EM UMP': 'Cantidad_103_Neta_Final'})
elif 'Cantidad_103' in resumen_103.columns:
    resumen_103 = resumen_103.rename(columns={'Cantidad_103': 'Cantidad_103_Neta_Final'})
else:
    raise KeyError("❌ No encontré la columna neta de movimientos 103 - 105 en resumen_103.")

resumen_103['Cantidad_103_Neta_Final'] = (
    pd.to_numeric(resumen_103['Cantidad_103_Neta_Final'], errors='coerce')
      .fillna(0)
      .clip(lower=0)
      .round(2)
)

# Normalizar Cantidad_103 (original) si viene del bloque 103-105
if 'Cantidad_103' in resumen_103.columns:
    resumen_103['Cantidad_103'] = (
        pd.to_numeric(resumen_103['Cantidad_103'], errors='coerce')
          .fillna(0)
          .clip(lower=0)
          .round(2)
    )

# Agrupar por si hay duplicados en resumen_103
agg_cols = {'Cantidad_103_Neta_Final': 'sum'}
if 'Cantidad_103' in resumen_103.columns:
    agg_cols['Cantidad_103'] = 'sum'

resumen_103 = (
    resumen_103
    .groupby(['Documento compras', 'Material', 'Posición'], as_index=False)
    .agg(agg_cols)
)

# ============================================================
# 4) Merge (left → mantiene materiales sin 103/105)
# ============================================================
merged = (
    resultado
      .reset_index(names='__orden__')
      .merge(
          resumen_103,
          on=['Documento compras', 'Material', 'Posición'],
          how='left'
      )
)

EPS = 1e-6

# ============================================================
# 5) Función de consolidación por grupo
# ============================================================
def _consolidar_grupo(g):
    """
    Ajusta solo la fila que corresponde al bloqueo 103, identificándola por
    su coincidencia con Cantidad_103 (original, antes de restar 105).

    - NaN en neta  → sin movimiento 103/105 → dejar todas las filas.
    - neta = 0     → eliminar la fila cuya Cantidad real ≈ Cantidad_103 original.
    - neta > 0     → ajustar esa misma fila a qty_neta; el resto sin cambio.
    """
    g = g.sort_values('__orden__').copy()
    qty_neta = g['Cantidad_103_Neta_Final'].iloc[0]

    # Caso 1: sin movimientos 103/105
    if pd.isna(qty_neta):
        return g

    qty_neta = float(qty_neta)

    # Cantidad original del bloqueo 103 (antes de restar 105)
    qty_original = (
        float(g['Cantidad_103'].iloc[0])
        if 'Cantidad_103' in g.columns and not pd.isna(g['Cantidad_103'].iloc[0])
        else qty_neta  # fallback: asumir que la neta es la referencia
    )

    # Identificar la fila del bloqueo: la que su Cantidad real ≈ Cantidad_103 original
    mask_103 = (g['Cantidad real'] - qty_original).abs() < EPS

    if qty_neta <= EPS:
        # Neto = 0: eliminar SOLO la fila del bloqueo; las demás quedan intactas
        resultado_g = g[~mask_103]
        if resultado_g.empty:
            # Solo había una fila y era el bloqueo → eliminar todo el grupo
            return g.iloc[0:0]
        return resultado_g
    else:
        # Neto > 0: ajustar la fila del bloqueo a qty_neta
        if mask_103.any():
            g.loc[mask_103, 'Cantidad real'] = round(qty_neta, 2)
            return g
        else:
            # No se identificó la fila exacta → consolidar todo en la primera fila
            primera = g.iloc[[0]].copy()
            primera['Cantidad real'] = round(qty_neta, 2)
            return primera

# ============================================================
# 6) Aplicar consolidación
# ============================================================
resultado_final = (
    merged
      .groupby(['Documento compras', 'Material', 'Posición'], group_keys=False)
      .apply(_consolidar_grupo)
      .reset_index(drop=True)
)

# ============================================================
# 7) Métricas
# ============================================================
qty_antes = float(merged['Cantidad real'].sum())
qty_despues = float(resultado_final['Cantidad real'].sum())
filas_antes = len(merged)
filas_despues = len(resultado_final)

print(f"🧹 Filas antes: {filas_antes} | después: {filas_despues} | eliminadas: {filas_antes - filas_despues}")
print(f"🔢 Cantidad total antes: {qty_antes:.2f} | después: {qty_despues:.2f} | diferencia: {qty_antes - qty_despues:.2f}")

# ============================================================
# 8) Limpiar columnas auxiliares
# ============================================================
resultado_final.drop(
    columns=['Cantidad_103_Neta_Final', 'Cantidad_103', '__orden__'],
    inplace=True,
    errors='ignore'
)

# ============================================================
# 9) Numeración
# ============================================================
resultado_final.insert(len(resultado_final.columns), 'N°', range(1, len(resultado_final) + 1))

# ============================================================
# 10) Renombrar Cantidad real → Cantidad
# ============================================================
if 'Cantidad real' in resultado_final.columns:
    resultado_final = resultado_final.rename(columns={'Cantidad real': 'Cantidad'})

# ============================================================
# 11) Reordenar columnas finales
# ============================================================
columnas_deseadas = [
    'N°',
    'Documento compras',
    'estado_final',
    'Material',
    'Texto breve',
    'Cantidad',
    'Centro',
    'Cl.documento compras',
    'Fecha final',
    'Nombre del proveedor',
    'Grupo de compras',
    'Posición'
]

resultado_final = resultado_final[[c for c in columnas_deseadas if c in resultado_final.columns]]

# ============================================================
# 12) Guardar archivo final
# ============================================================
archivo_ajustado = ruta / 'Transito_Base_final.xlsx'
resultado_final.to_excel(archivo_ajustado, index=False)

print(f"✅ Archivo ajustado guardado como: {archivo_ajustado}")


