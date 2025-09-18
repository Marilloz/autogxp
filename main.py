"""
 * AutoCalculadorDeTramos
 * Copyright © 2023-2025  Marcos Martín Sandeogracias
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.

/* SPDX-License-Identifier: GPL-3.0 https://www.gnu.org/licenses/licenses/license-object.html*/
"""

import gpxpy
import math
import xlwings as xw
from xlwings.constants import AutoFillType
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
import os, sys, pathlib

help_texts = [
    "Longitud mínima (en metros) para que un tramo sea considerado. \nEj, si el valor es 100, todos los tramos menores a 100 metros serán juntados con siguiente tramo.\nPara eliminar esta variable, hay que ponerla a 0",
    "Diferencia mínima de elevación (en metros) que un tramo sea considerado. \nEj, si el valor es 10, todos los tramos con menos elevación que 10 metros serán juntados con el siguiente tramo.\nPara eliminar esta variable, hay que ponerla a 0",
    "Valor máximo permitido de pendiente (como decimal, ej. 0.6 = 60%).\nCualquier otro valor se considera un error y por tanto el tramo es juntado con el siguiente tramo.\nPara eliminar esta variable, hay que ponerla a 1",
    "Distancia mínima horizontal (en metros) para filtrar tramos.\nEste valor se usa en conjunto con la Elevación mínima asociada\nSi la distancia y la elevación del tramo son ambas menores a estas variables, el tramo será juntado con el siguiente tramo.\nPara eliminar esta variable, hay que ponerla a 0",
    "Altura mínima (en metros) para considerar elevación relevante.\nEste valor se usa en conjunto con la Longitud horizontal mínima\n Si la distancia y la elevación del tramo son ambas menores a estas variables, el tramo será juntado con el siguiente tramo.\nPara eliminar esta variable, hay que ponerla a 0",
]

global_help_text = """La herramienta funciona cogiendo todos los puntos del archivo GPX (Conjunto de Coordenadas + Altitud).
Calcula los tramos entre dichos puntos, sacando la distancia y la elevación. Diferenciando entre elevaciones positivas y negativas.
Luego, junta tramos (suma distancia y elevación) contiguos con elevaciones del mismo signo (Subidas con subidas, bajadas con bajadas, etc)
Tras esto, aplica los umbrales de la elevación, para reducir el número de tramos.
Después, se vuelven a juntar por elevaciones del mismo signo.
Y para sacar los tramos finales, se aplican el resto de umbrales (distancia, pendiente y umbral combinado) y se agrupan por elevaciones del mismo signo.
Por último, usando esos tramos y la plantilla del calculador de tramos, se genera el archivo de calculador de tramos con los datos ya introducidos.
"""

def resource_path(relative_path):
    # si estamos en modo onefile esto existe
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def colnum_a_letra(n):
    letra = ''
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra

def haversine(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)

    a = math.sin(dphi/2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda/2)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def calcular_pendientes(gpx_file_path):
    with open(gpx_file_path, 'r') as gpx_file:
        gpx = gpxpy.parse(gpx_file)

    puntos = []
    for track in gpx.tracks:
        for segment in track.segments:
            for point in segment.points:
                if point.elevation is not None and point.elevation > 5:
                    puntos.append((point.latitude, point.longitude, point.elevation))

    pendientes = []
    for i in range(1, len(puntos)):
        lat1, lon1, ele1 = puntos[i - 1]
        lat2, lon2, ele2 = puntos[i]
        distancia_horizontal = haversine(lat1, lon1, lat2, lon2)
        delta_elevacion = ele2 - ele1

        if distancia_horizontal > 0:
            pendientes.append({
                'distancia_m': distancia_horizontal,
                'elevacion_m': delta_elevacion
            })

    return pendientes

def agrupar_por_direccion(tramos):
    if not tramos:
        return []

    agrupados = []
    tramo_actual = tramos[0].copy()
    direccion = tramo_actual['elevacion_m'] >= 0

    for tramo in tramos[1:]:
        misma_direccion = (tramo['elevacion_m'] >= 0) == direccion
        if misma_direccion:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']
        else:
            agrupados.append(tramo_actual)
            tramo_actual = tramo.copy()
            direccion = tramo_actual['elevacion_m'] >= 0

    agrupados.append(tramo_actual)
    return agrupados

def agrupar_por_umbral(tramos, umbral_elevacion):
    if not tramos:
        return []

    finales = []
    tramo_actual = tramos[0].copy()

    for tramo in tramos[1:]:
        if abs(tramo['elevacion_m']) < umbral_elevacion:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']
        elif abs(tramo['elevacion_m'])/abs(tramo['distancia_m'])>0.6:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']

        else:
            finales.append(tramo_actual)
            tramo_actual = tramo.copy()

    finales.append(tramo_actual)
    return finales

def agrupar_por_umbral2(tramos, max_pendiente,min_tramo,distancia_horizontal, delta_elevacion):
    if not tramos:
        return []

    finales = []
    tramo_actual = tramos[0].copy()

    for tramo in tramos[1:]:
        if abs(tramo['elevacion_m'])/abs(tramo['distancia_m'])>max_pendiente:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']
        elif abs(tramo['distancia_m']) < min_tramo:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']
        elif abs(tramo['distancia_m']) < distancia_horizontal and abs(tramo['elevacion_m']) < delta_elevacion:
            tramo_actual['distancia_m'] += tramo['distancia_m']
            tramo_actual['elevacion_m'] += tramo['elevacion_m']
        else:
            finales.append(tramo_actual)
            tramo_actual = tramo.copy()

    finales.append(tramo_actual)
    return finales

def calcular_pendiente_y_enumerar(tramos):
    for i, tramo in enumerate(tramos, 1):
        if tramo['distancia_m'] > 0:
            pendiente = (tramo['elevacion_m'] / tramo['distancia_m']) * 100
        else:
            pendiente = 0
        tramo['tramo'] = i
        tramo['pendiente_%'] = round(pendiente, 2)
        tramo['distancia_m'] = round(tramo['distancia_m'], 2)
        tramo['elevacion_m'] = round(tramo['elevacion_m'], 2)
    return tramos

def get_tramos_finales(gpx_file, umbral_elevacion, pendiente_maxima_valida, longitud_minima_tramo,
         longitud_horizontal_minima, elevacion_minima_asociada,):
    tramos_raw = calcular_pendientes(gpx_file)
    tramos_direccion = agrupar_por_direccion(tramos_raw)
    tramos_umbral = agrupar_por_umbral(tramos_direccion, umbral_elevacion)
    tramos_mix = agrupar_por_direccion(tramos_umbral)

    tramos_mix = agrupar_por_direccion(agrupar_por_umbral2(
        tramos_mix,
        pendiente_maxima_valida,
        longitud_minima_tramo,
        longitud_horizontal_minima,
        elevacion_minima_asociada
    ))

    if abs(tramos_mix[0].get('elevacion_m')) < umbral_elevacion or abs(
            tramos_mix[0].get('distancia_m') < longitud_minima_tramo):
        tramos_mix[1]['distancia_m'] += tramos_mix[0]['distancia_m']
        tramos_mix[1]['elevacion_m'] += tramos_mix[0]['elevacion_m']
        tramos_mix.pop(0)

    tramos_finales = calcular_pendiente_y_enumerar(tramos_mix)
    return tramos_finales

def rellenar_plantilla(ws, tramos_finales, seccion, preparacion, descanso, cada):
    ws.range(f"C4").value = seccion
    ws.range(f"E4").value = preparacion
    ws.range(f"C6").value = descanso
    ws.range(f"E6").value = cada

    extra_rows = max(len(tramos_finales) - 10, 0)

    if extra_rows > 0:
        fila_origen = 21
        for i in range(extra_rows):
            # Inserta una fila en la posición deseada
            ws.range(f"{fila_origen + i}:{fila_origen + i}").insert(shift="down")

        ws.range(f'{fila_origen - 1}:{fila_origen - 1}').api.AutoFill(
            ws.range(f"{fila_origen - 1}:{fila_origen + extra_rows}").api, AutoFillType.xlFillDefault)

    fila_inicial = 12

    for i, tramo in enumerate(tramos_finales):
        fila = fila_inicial + i
        horizontal = tramo['distancia_m']
        desnivel = tramo['elevacion_m']
        if desnivel > 0:
            tipo = "Ascenso"
        elif desnivel < 0:
            tipo = "Descenso"
        else:
            tipo = "Llano"

        ws.range(f"C{fila}").value = round(horizontal / 1000, 2)  # por ejemplo, 3 decimales
        ws.range(f"D{fila}").value = tipo
        ws.range(f"E{fila}").value = abs(desnivel)


def main(gpx_file, gpx_output, umbral_elevacion, pendiente_maxima_valida, longitud_minima_tramo,
         longitud_horizontal_minima, elevacion_minima_asociada,
         seccion, preparacion, descanso, cada):
    # USO
    path_gpx = pathlib.Path(gpx_file)

    if path_gpx.is_file():

        tramos_finales = get_tramos_finales(gpx_file, umbral_elevacion, pendiente_maxima_valida, longitud_minima_tramo,
         longitud_horizontal_minima, elevacion_minima_asociada)

        app = xw.App(visible=False)

        plantilla = resource_path('plantilla.xlsx')
        wb = app.books.open(plantilla)

        ws = wb.sheets[0]  # Primera hoja

        rellenar_plantilla(ws, tramos_finales, seccion, preparacion, descanso, cada)

        wb.save(gpx_output)
        wb.close()
        app.quit()  # Cierra Excel por completo
    elif path_gpx.is_dir():
        tramos_de_ficheros  = []
        for file in path_gpx.iterdir():
            if not file.is_file() or file.suffix.upper() != ".GPX":
                continue
            tramos_finales = get_tramos_finales(file, umbral_elevacion, pendiente_maxima_valida,
                                                longitud_minima_tramo,
                                                longitud_horizontal_minima, elevacion_minima_asociada)
            tramos_de_ficheros.append((file.stem,tramos_finales))

        plantilla = resource_path('plantilla.xlsx')
        app = xw.App(visible=False)

        wb = app.books.open(plantilla)
        ws = wb.sheets[0]

        def remove_accents(s: str) -> str:
            import unicodedata
            """Elimina acentos de una cadena."""
            return ''.join(
                c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn'
            )

        processed = [
            (remove_accents(name.lower()).title(), path)
            for name, path in tramos_de_ficheros
        ]
        # Ordenar por el nombre (primer elemento de la tupla)
        tramos_de_ficheros = sorted(processed, key=lambda x: x[0])

        for i in range(len(tramos_de_ficheros)-1):
            # copia ws y la coloca al final
            ws.copy(after=wb.sheets[-1])

        for i, elem in enumerate(tramos_de_ficheros):
            ws = wb.sheets[i]
            ws.name = elem[0]
            rellenar_plantilla(ws, elem[1], seccion, preparacion, descanso, cada)

        wb.sheets[0].activate()

        wb.save(gpx_output)
        wb.close()

        app.quit()  # Cierra Excel por completo


def seleccionar_archivo(entry_widget, filetypes):
    archivo = filedialog.askopenfilename(filetypes=filetypes)
    if archivo:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, archivo)

def seleccionar_archivo_salida(entry_widget, filetypes):
    archivo = filedialog.asksaveasfilename(filetypes=filetypes)
    if archivo:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, archivo)

def seleccionar_carpeta(entry_widget):
    carpeta = filedialog.askdirectory()
    if carpeta:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, carpeta)
def autocompletar(entry_widget1, entry_widget2):
    if entry_widget1.get():
        path = pathlib.Path(entry_widget1.get())
        if not path.exists() or path.is_file():
            path = path.with_suffix(".xlsx")

        else:
            path = path / "CalculadorDeTramos.xlsx"

    else:
        path = "CalculadorDeTramos.xlsx"

    entry_widget2.delete(0, tk.END)
    entry_widget2.insert(0, path)


def crear_gui():
    root = tk.Tk()
    root.title("Auto Calculador de Tramos GPX")
    root.geometry("570x385")
    root.minsize(570, 385)
    root.resizable(True, False)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(4, weight=1)

    # Frame de selección de archivos
    frame_archivos = tk.Frame(root)
    frame_archivos.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
    frame_archivos.columnconfigure(1, weight=1)

    tk.Label(frame_archivos, text="Archivo GPX de entrada:").grid(row=0, column=0, sticky="w", padx=(0,5))
    entry_gpx = tk.Entry(frame_archivos)
    entry_gpx.grid(row=0, column=1, sticky="ew")
    tk.Button(frame_archivos, text="Seleccionar", command=lambda: seleccionar_archivo(entry_gpx, [("GPX files","*.gpx")])).grid(row=0, column=2, padx=(5,0))
    tk.Button(frame_archivos, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta(entry_gpx, )).grid(row=0, column=3, padx=(5,0))

    tk.Label(frame_archivos, text="Archivo XLSX de salida:").grid(row=1, column=0, sticky="w", padx=(0,5), pady=(5,0))
    entry_xlsx = tk.Entry(frame_archivos)
    entry_xlsx.grid(row=1, column=1, sticky="ew", pady=(5,0))
    tk.Button(frame_archivos, text="Seleccionar", command=lambda: seleccionar_archivo_salida(entry_xlsx, [("Excel files","*.xlsx")])).grid(row=1, column=2, padx=(5,0), pady=(5,0))
    tk.Button(frame_archivos, text="Autocompletar", command=lambda: autocompletar(entry_gpx, entry_xlsx)).grid(row=1, column=3, padx=(5,0), pady=(5,0), sticky='ew')

    # Frame de Seccion y Preparacion
    frame_opciones = tk.Frame(root)
    frame_opciones.grid(row=1, column=0, padx=10, sticky="ew")
    frame_opciones.columnconfigure(1, weight=1)
    frame_opciones.columnconfigure(3, weight=1)

    tk.Label(frame_opciones, text="Sección:").grid(row=0, column=0, sticky="w")
    combo_seccion = Combobox(frame_opciones, state="readonly", values=["Colonia", "Manada", "Scout", "Unidad Esculta", "Clan/Rovers"])
    combo_seccion.current(2)
    combo_seccion.grid(row=0, column=1, sticky="ew", padx=(5,15))

    tk.Label(frame_opciones, text="Preparación:").grid(row=0, column=2, sticky="w")
    combo_preparacion = Combobox(frame_opciones, state="readonly", values=["Muy baja", "Baja", "Media", "Alta", "Muy alta"])
    combo_preparacion.current(2)
    combo_preparacion.grid(row=0, column=3, sticky="ew", padx=(5,0))

    tk.Label(frame_opciones, text="Descanso (min):").grid(row=1, column=0, sticky="w", pady=(5, 0))
    entry_descanso = tk.Entry(frame_opciones)
    entry_descanso.insert(0, "10")
    entry_descanso.grid(row=1, column=1, sticky="ew", padx=(5, 15), pady=(5, 0))

    tk.Label(frame_opciones, text="Cada (min):").grid(row=1, column=2, sticky="w", pady=(5, 0))
    entry_cada = tk.Entry(frame_opciones)
    entry_cada.insert(0, "60")
    entry_cada.grid(row=1, column=3, sticky="ew", padx=(5, 0), pady=(5, 0))

    # Frame valores expertos
    expert_frame = tk.LabelFrame(root, text="VALORES PARA EXPERTOS - NO TOCAR SI NO SABES LO QUE HACES")
    expert_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
    expert_frame.columnconfigure(1, weight=1)

    tk.Label(expert_frame, text="Cuanto más bajos sean estos valores, mayor cantidad de tramos habrá, pero mayor la precisión final.", wraplength=550, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(5,10))

    # Creación de entradas expertas
    def add_expert_param(row, label, default_value, help_text):
        tk.Label(expert_frame, text=label).grid(row=row+1, column=0, sticky="w", padx=(0,5), pady=2)
        entry = tk.Entry(expert_frame)
        entry.insert(0, str(default_value))
        entry.grid(row=row+1, column=1, sticky="ew", pady=2)
        tk.Button(expert_frame, text="?", command=lambda: messagebox.showinfo("Ayuda", help_text)).grid(row=row+1, column=2, padx=(5,10))
        return entry

    entry_tramo = add_expert_param(0, "Longitud mínima del tramo:", 100, help_texts[0])
    entry_umbral = add_expert_param(1, "Umbral de elevación:", 10, help_texts[1])
    entry_pendiente = add_expert_param(2, "Pendiente máxima válida:", 0.6, help_texts[2])
    entry_horizontal = add_expert_param(3, "Longitud horizontal mínima:", 300, help_texts[3])
    entry_elevacion = add_expert_param(4, "Elevación mínima asociada:", 25, help_texts[4])

    def ejecutar(gpx_path, xlsx_path, seccion, preparacion):
        # Validación básica de rutas
        if not gpx_path:
            messagebox.showerror("Error", "Por favor, seleccione el archivo GPX de entrada.")
            return
        if not gpx_path.lower().endswith('.gpx') or not os.path.isfile(gpx_path):
            path = pathlib.Path(gpx_path)
            if not path.is_dir():
                messagebox.showerror("Error", "El archivo de entrada debe existir y tener extensión .gpx")
                return
        if not xlsx_path:
            path = pathlib.Path(gpx_path)
            if not path.exists() or path.is_file():
                path = path.with_suffix(".xlsx")

            else:
                path = path / "CalculadorDeTramos.xlsx"
            xlsx_path = str(path)
        # Asegurar extensión .xlsx
        if not xlsx_path.lower().endswith('.xlsx'):
            xlsx_path += '.xlsx'
        # Verificar carpeta de salida
        carpeta = os.path.dirname(xlsx_path) or os.getcwd()
        if not os.path.isdir(carpeta):
            messagebox.showerror("Error", f"La carpeta de salida '{carpeta}' no existe.")
            return
        # Intentar crear archivo vacío para verificar permisos
        try:
            open(xlsx_path, 'a').close()
            os.remove(xlsx_path)
        except Exception as e:
            messagebox.showerror("Error", f"No se puede escribir en '{xlsx_path}': {e}")
            return

        descanso_val = int(entry_descanso.get())
        cada_val = int(entry_cada.get())
        if cada_val <= 0:
            cada_val = 1

        # Ejecutar lógica principal
        try:
            main(
                gpx_file=gpx_path,
                gpx_output=xlsx_path,
                umbral_elevacion=int(entry_umbral.get()),
                pendiente_maxima_valida=float(entry_pendiente.get()),
                longitud_minima_tramo=int(entry_tramo.get()),
                longitud_horizontal_minima=int(entry_horizontal.get()),
                elevacion_minima_asociada=int(entry_elevacion.get()),
                seccion=seccion,
                preparacion=preparacion,
                descanso=descanso_val,
                cada=cada_val,
            )
            messagebox.showinfo("Proceso completado",
                                f"Se ha generado el calculador de tramos del archivo {os.path.basename(gpx_path)}.\n"
                                f"Lo puedes encontrar en el archivo {os.path.basename(xlsx_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante la ejecución:\n{e}")

    # --- Barra inferior: licencia | Ejecutar | ? --------------------------
    button_frame = tk.Frame(root)
    # padding arriba-abajo: 10 px arriba, 15 px abajo
    button_frame.grid(row=3, column=0, padx=10, pady=(10, 15), sticky="ew")

    # Que la fila ocupe todo el ancho y las columnas 0 y 2 “empujen” al centro
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=0)
    button_frame.columnconfigure(2, weight=1)

    # --- Columna 0: firma /licencia (alineada a la izquierda) -------------
    firma_frame = tk.Frame(button_frame)
    firma_frame.grid(row=0, column=0, sticky="w")

    tk.Label(firma_frame, text="By Marcos Martín Sandeogracias").grid(row=0, column=0, sticky="w")
    tk.Label(firma_frame, text="Licencia GPL v3").grid(row=1, column=0, sticky="w")

    # --- Columna 1: botón Ejecutar (centro exacto) ------------------------
    tk.Button(
        button_frame,
        text="Ejecutar",
        bg="green",
        fg="white",
        width=20,
        command=lambda: ejecutar(
            entry_gpx.get(),
            entry_xlsx.get(),
            combo_seccion.get(),
            combo_preparacion.get()
        )
    ).grid(row=0, column=1, padx=10)

    # --- Columna 2: botón de ayuda ? (alineado a la derecha) --------------
    tk.Button(
        button_frame,
        text="?",
        width=3,
        command=lambda: messagebox.showinfo("Ayuda general", global_help_text)
    ).grid(row=0, column=2, sticky="e")
    # ---------------------------------------------------------------------



    root.mainloop()



if __name__ == "__main__":
    crear_gui()