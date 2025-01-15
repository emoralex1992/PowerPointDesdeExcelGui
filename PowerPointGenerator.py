# COMANDO PARA REGENERAR EL INSTALADOR:
# pyinstaller --onefile --windowed --icon="assets/logo_engiacademy.ico" .\PowerPointGenerator.py

import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
from pathlib import Path
from main import main, configurar_variables, ExportarNombrePlaceholders
import os
import json

# Ruta para el archivo temporal de configuración
TEMP_CONFIG_PATH = os.path.join(os.getcwd(), "temp", "config.json")

def cargar_configuracion():
    """Cargar configuración desde el archivo temporal si existe."""
    if os.path.exists(TEMP_CONFIG_PATH):
        try:
            with open(TEMP_CONFIG_PATH, 'r') as file:
                return json.load(file)
        except Exception as e:
            print(f"Error al cargar configuración: {e}")
    return {}

def guardar_configuracion(config):
    """Guardar configuración en un archivo temporal."""
    os.makedirs(os.path.dirname(TEMP_CONFIG_PATH), exist_ok=True)
    try:
        with open(TEMP_CONFIG_PATH, 'w') as file:
            json.dump(config, file)
    except Exception as e:
        print(f"Error al guardar configuración: {e}")

def seleccionar_ruta(entry_widget, tipo):
    if tipo == "file":
        archivo = filedialog.askopenfilename(filetypes=[("Archivos", "*.pptx;*.xlsx")])
        if archivo:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, archivo)
    elif tipo == "folder":
        carpeta = filedialog.askdirectory()
        if carpeta:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, carpeta)

def verificar_sobrescritura(archivo):
    """Verificar si un archivo existe y preguntar si desea sobrescribirlo."""
    if Path(archivo).exists():
        return messagebox.askyesno("Advertencia", f"El archivo {archivo} ya existe. ¿Desea sobrescribirlo?")
    return True

def ejecutar_script(input_pptx, input_excel, output_path, output_file):
    if not input_pptx or not input_excel or not output_path or not output_file:
        messagebox.showerror("Error", "Todos los campos son obligatorios.")
        return

    if not Path(input_pptx).is_file():
        messagebox.showerror("Error", f"La plantilla PowerPoint no existe: {input_pptx}")
        return

    if not Path(input_excel).is_file():
        messagebox.showerror("Error", f"El archivo Excel no existe: {input_excel}")
        return

    if not Path(output_path).is_dir():
        messagebox.showerror("Error", f"La carpeta de salida no existe: {output_path}")
        return

    if not output_file.endswith(".pptx"):
        output_file += ".pptx"

    output_full_path = Path(output_path) / output_file

    if not verificar_sobrescritura(output_full_path):
        return

    guardar_configuracion({
        "input_pptx": input_pptx,
        "input_excel": input_excel,
        "output_path": output_path,
        "output_file": output_file
    })

    try:
        configurar_variables(input_pptx, input_excel, output_path, output_file)
        main()
        messagebox.showinfo("Éxito", f"Presentación generada correctamente en: {output_full_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

def ejecutar_marcadores(input_pptx, output_path):
    if not input_pptx or not output_path:
        messagebox.showerror("Error", "La plantilla y la carpeta de salida son obligatorias para generar los marcadores.")
        return

    if not Path(input_pptx).is_file():
        messagebox.showerror("Error", f"La plantilla PowerPoint no existe: {input_pptx}")
        return

    if not Path(output_path).is_dir():
        messagebox.showerror("Error", f"La carpeta de salida no existe: {output_path}")
        return

    output_full_path = Path(output_path) / "Marcadores.pptx"

    if not verificar_sobrescritura(output_full_path):
        return

    guardar_configuracion({
        "input_pptx": input_pptx,
        "output_path": output_path
    })

    try:
        configurar_variables(input_pptx, None, output_path, "Marcadores.pptx")
        ExportarNombrePlaceholders(input_pptx, output_path)
        messagebox.showinfo("Éxito", f"Marcadores generados correctamente en: {output_full_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

def exportar_lista_diseños(input_pptx, output_path):
    """Exporta una lista de nombres de diseños a un archivo Excel."""
    if not input_pptx or not output_path:
        messagebox.showerror("Error", "La plantilla y la carpeta de salida son obligatorias.")
        return

    if not Path(input_pptx).is_file():
        messagebox.showerror("Error", f"La plantilla PowerPoint no existe: {input_pptx}")
        return

    if not Path(output_path).is_dir():
        messagebox.showerror("Error", f"La carpeta de salida no existe: {output_path}")
        return

    # Ruta del archivo Excel de salida
    output_file = Path(output_path) / "Lista_Diseños.xlsx"

    # Verificar si el archivo ya existe
    if not verificar_sobrescritura(output_file):
        return

    try:
        # Llamar a la función del script principal para listar los diseños
        from main import listar_diseños
        lista_diseños = listar_diseños(input_pptx)

        # Crear un DataFrame con la lista de diseños
        import pandas as pd
        df = pd.DataFrame({"Nombres de Diseños": lista_diseños})

        # Guardar el DataFrame en un archivo Excel
        df.to_excel(output_file, index=False)

        messagebox.showinfo("Éxito", f"Lista de diseños exportada correctamente en: {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al exportar la lista de diseños: {str(e)}")


def main_gui():
    root = tk.Tk()
    root.title("Generador de Presentaciones PowerPoint")

    # Ruta del icono
    assets_path = os.path.join(os.getcwd(), "assets")
    icon_path = os.path.join(assets_path, "logo_engiacademy.png")

    try:
        icon_image = PhotoImage(file=icon_path)
        root.iconphoto(True, icon_image)  # Establecer el icono de la ventana
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el icono de la ventana: {str(e)}")

    # Cargar configuración previa
    config = cargar_configuracion()

    tk.Label(root, text="Plantilla PowerPoint:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    input_pptx_entry = tk.Entry(root, width=50)
    input_pptx_entry.grid(row=0, column=1, padx=10, pady=5)
    input_pptx_entry.insert(0, config.get("input_pptx", ""))
    tk.Button(root, text="Seleccionar", command=lambda: seleccionar_ruta(input_pptx_entry, "file")).grid(row=0, column=2, padx=10, pady=5)

    tk.Label(root, text="Archivo Excel de Configuración:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    input_excel_entry = tk.Entry(root, width=50)
    input_excel_entry.grid(row=1, column=1, padx=10, pady=5)
    input_excel_entry.insert(0, config.get("input_excel", ""))
    tk.Button(root, text="Seleccionar", command=lambda: seleccionar_ruta(input_excel_entry, "file")).grid(row=1, column=2, padx=10, pady=5)

    tk.Label(root, text="Carpeta de Salida:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    output_path_entry = tk.Entry(root, width=50)
    output_path_entry.grid(row=2, column=1, padx=10, pady=5)
    output_path_entry.insert(0, config.get("output_path", ""))
    tk.Button(root, text="Seleccionar", command=lambda: seleccionar_ruta(output_path_entry, "folder")).grid(row=2, column=2, padx=10, pady=5)

    tk.Label(root, text="Nombre del Archivo de Salida: ").grid(row=3, column=0, padx=10, pady=5, sticky="e")
    output_file_entry = tk.Entry(root, width=40)
    output_file_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
    output_file_entry.insert(0, config.get("output_file", ""))

    frame_buttons = tk.Frame(root)
    frame_buttons.grid(row=4, column=0, columnspan=3, pady=10)

    # Cargar iconos
    assets_path = os.path.join(os.getcwd(), "assets")
    try:
        play_icon = PhotoImage(file=os.path.join(assets_path, "play_icon.png"))
        markers_icon = PhotoImage(file=os.path.join(assets_path, "marcadores_icon.png"))
        slides_icon = PhotoImage(file=os.path.join(assets_path, "slides_icon.png"))
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron cargar los iconos: {str(e)}")
        play_icon = None
        markers_icon = None
        slides_icon = None

    # Botón Generar Lista Diseños
    tk.Button(frame_buttons, text=" Generar Lista Diseños", image=slides_icon, compound="left", padx=8, pady=8,
              command=lambda: exportar_lista_diseños(input_pptx_entry.get(), output_path_entry.get())).pack(side="left", padx=5)

    # Botón Generar Presentación
    tk.Button(frame_buttons, text=" Generar Presentación", image=play_icon, compound="left", padx=8, pady=8,
              command=lambda: ejecutar_script(input_pptx_entry.get(), input_excel_entry.get(), output_path_entry.get(), output_file_entry.get())).pack(side="right", padx=5)

    # Botón Generar Marcadores
    tk.Button(frame_buttons, text=" Generar Marcadores.pptx", image=markers_icon, compound="left", padx=8, pady=8,
              command=lambda: ejecutar_marcadores(input_pptx_entry.get(), output_path_entry.get())).pack(side="left", padx=5)

    root.mainloop()

if __name__ == "__main__":
    main_gui()
