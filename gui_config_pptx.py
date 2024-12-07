import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
from pathlib import Path
from main import main, ExportarNombrePlaceholders as generar_marcadores
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

    guardar_configuracion({
        "input_pptx": input_pptx,
        "input_excel": input_excel,
        "output_path": output_path,
        "output_file": output_file
    })

    try:
        main(input_pptx, input_excel, output_path, output_file)
        messagebox.showinfo("Éxito", f"Presentación generada correctamente en: {Path(output_path) / output_file}")
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

    guardar_configuracion({
        "input_pptx": input_pptx,
        "output_path": output_path
    })

    try:
        generar_marcadores(input_pptx, output_path)
        messagebox.showinfo("Éxito", f"Marcadores generados correctamente en: {Path(output_path) / 'Marcadores.pptx'}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

def main_gui():
    root = tk.Tk()
    root.title("Generador de Presentaciones PowerPoint")

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

    tk.Label(root, text="Nombre del Archivo de Salida:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
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
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron cargar los iconos: {str(e)}")
        play_icon = None
        markers_icon = None


    # Botón Generar Marcadores
    tk.Button(frame_buttons, text=" Generar Marcadores.pptx", image=markers_icon, compound="left", padx=8, pady=8,
              command=lambda: ejecutar_marcadores(input_pptx_entry.get(), output_path_entry.get())).pack(side="left", padx=5)

    # Botón Generar Presentación
    tk.Button(frame_buttons, text=" Generar Presentación", image=play_icon, compound="left", padx=8, pady=8,
              command=lambda: ejecutar_script(input_pptx_entry.get(), input_excel_entry.get(), output_path_entry.get(), output_file_entry.get())).pack(side="left", padx=5)


    root.mainloop()

if __name__ == "__main__":
    main_gui()
