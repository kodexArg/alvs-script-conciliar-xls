import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd
from pandas import DataFrame
from logic import process_logic  # Importa la lógica desde logic.py

def log_message(log_widget: scrolledtext.ScrolledText, message: str) -> None:
    """Agrega mensajes al cuadro de log."""
    log_widget.insert(tk.END, message + "\n")
    print(message + "\n")
    log_widget.see(tk.END)
    log_widget.update_idletasks()


def select_file(entry_widget: tk.Entry, log_widget: scrolledtext.ScrolledText, file_type: str) -> None:
    """Abre un diálogo para seleccionar un archivo y actualiza el campo."""
    file_path: str = filedialog.askopenfilename(title=f"Seleccionar archivo de {file_type}", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)
        log_message(log_widget, f"{file_type} seleccionado: {file_path}")


def import_cobranza_electronica(file_path: str, log_widget: scrolledtext.ScrolledText) -> DataFrame | None:
    """Importa y limpia el archivo COBRANZAS ELECTRONICAS desde la fila 7 en adelante, tanto de KM1151 como LAS BOVEDAS, excluyendo totales."""
    try:
        # Leer el archivo comenzando desde la fila 7
        df: DataFrame = pd.read_excel(file_path, skiprows=6)

        # Limpiar el DataFrame: eliminar columnas completamente vacías
        df = df.dropna(how='all', axis=1)

        # Reiniciar el índice para asegurar que está bien estructurado
        df = df.reset_index(drop=True)

        # Convertir columnas relevantes al tipo adecuado, por ejemplo, la columna de fecha
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')

        # Eliminar filas donde 'Transacción' sea NaN o no sea numérico (posibles totales)
        df = df[df['Transacción'].apply(lambda x: pd.notna(x) and str(x).isdigit())]

        log_message(log_widget, "Cobranzas Electrónicas importado y limpiado correctamente.")
        
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error al importar y limpiar Cobranzas Electrónicas: {e}")
        return None





def import_mercado_pago(file_path: str, log_widget: scrolledtext.ScrolledText) -> DataFrame | None:
    """Importa y limpia el archivo MERCADO PAGO."""
    try:
        # Leer el archivo Excel
        df: DataFrame = pd.read_excel(file_path)

        # Asegurar que la columna de 'Fecha de Pago' está en formato datetime
        df['Fecha de Pago'] = pd.to_datetime(df['Fecha de Pago'], format='%Y-%m-%dT%H:%M:%SZ', errors='coerce')

        # Convertir la columna 'Importe' a un tipo numérico
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce')

        log_message(log_widget, "Mercado Pago importado y limpiado correctamente.")

        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error al importar y limpiar Mercado Pago: {e}")
        return None


def import_planilla_1(file_path: str, log_widget: scrolledtext.ScrolledText) -> DataFrame | None:
    """Importa y limpia la hoja 'Transferencias' del archivo PLANILLA 1."""
    try:
        # Leer sólo la hoja 'Transferencias' empezando desde la segunda fila
        df: DataFrame = pd.read_excel(file_path, sheet_name='Transferencias', skiprows=1)

        # Mostrar los nombres de las columnas
        log_message(log_widget, f"Columnas detectadas en Planilla 1: {df.columns.tolist()}")

        # Limpiar el DataFrame: eliminar columnas completamente vacías
        df = df.dropna(how='all', axis=1)

        # Normalizar nombres de columnas
        df.columns = df.columns.str.strip()

        # Verificar si la columna 'Importe' existe después de la normalización
        if 'Importe' not in df.columns:
            log_message(log_widget, f"Error: La columna 'Importe' no se encuentra en el archivo Planilla 1.")
            return None

        # Asegurarse de que todos los valores en la columna 'Importe' sean cadenas
        df['Importe'] = df['Importe'].astype(str)

        # Limpiar y convertir la columna 'Importe' a numérico
        df['Importe'] = pd.to_numeric(df['Importe'].str.replace(r'[\$,]', '', regex=True), errors='coerce')

        # Convertir la columna 'Fecha' al tipo adecuado
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')

        log_message(log_widget, "Planilla 1 (hoja 'Transferencias') importado correctamente.")

        return df
    except Exception as e:
        # Mostrar el error exacto en el log
        log_message(log_widget, f"Error al importar y limpiar Planilla 1: {e}")
        messagebox.showerror("Error", f"Error al importar y limpiar Planilla 1: {e}")
        return None


def import_files(mp_entry: tk.Entry, planilla_entry: tk.Entry, cob_km_entry: tk.Entry, cob_bovedas_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> tuple[DataFrame | None, DataFrame | None, DataFrame | None, DataFrame | None]:
    """Importa los archivos XLSX seleccionados por el usuario y devuelve los DataFrames correspondientes."""
    log_message(log_widget, "Importando archivos...")

    file_mp = mp_entry.get()
    file_planilla = planilla_entry.get()
    file_km = cob_km_entry.get()
    file_bovedas = cob_bovedas_entry.get()

    if file_mp and file_planilla and file_km and file_bovedas:
        try:
            df_mp = import_mercado_pago(file_mp, log_widget)
            df_planilla = import_planilla_1(file_planilla, log_widget)
            df_cob_km = import_cobranza_electronica(file_km, log_widget)
            df_cob_bovedas = import_cobranza_electronica(file_bovedas, log_widget)
            return df_mp, df_planilla, df_cob_km, df_cob_bovedas
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar archivos: {e}")
            return None, None, None, None


def generate_output(df: DataFrame, save_path: str, log_widget: scrolledtext.ScrolledText) -> None:
    """Guarda el DataFrame resultante en un archivo XLSX en la ubicación seleccionada por el usuario."""
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            log_message(log_widget, f"Archivo guardado exitosamente en {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo: {e}")


def run_process(mp_entry: tk.Entry, planilla_entry: tk.Entry, cob_km_entry: tk.Entry, cob_bovedas_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> None:
    """
    Ejecuta el flujo principal de la aplicación:
    1. Importa los archivos seleccionados por el usuario como DataFrames.
    2. Llama a la lógica principal para procesar los DataFrames.
    3. Genera los archivos XLSX con los DataFrames resultantes.
    """
    df_mp, df_planilla, df_cob_km, df_cob_bovedas = import_files(mp_entry, planilla_entry, cob_km_entry, cob_bovedas_entry, log_widget)
    if df_mp is not None and df_planilla is not None and df_cob_km is not None and df_cob_bovedas is not None:
        log_message(log_widget, "Iniciando procesado de ficheros...")
        
        df_result, df_km1151_residuo, df_bovedas_residuo = process_logic(log_widget, df_mp, df_planilla, df_cob_km, df_cob_bovedas)
        
        # Generar los archivos resultantes
        generate_output(df_result, os.path.join(os.getcwd(), "conciliacion_result.xlsx"), log_widget)
        generate_output(df_km1151_residuo, os.path.join(os.getcwd(), "cobranzas_km1151_residuo.xlsx"), log_widget)
        generate_output(df_bovedas_residuo, os.path.join(os.getcwd(), "cobranzas_bovedas_residuo.xlsx"), log_widget)
        
        log_message(log_widget, "Proceso finalizado.")
        messagebox.showinfo("Éxito", "Proceso finalizado con éxito.")



def generate_no_conciliados_output(df_no_conciliados: DataFrame, output_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> None:
    """Guarda el DataFrame de registros no conciliados en un archivo XLSX."""
    save_path: str = output_entry.get()
    no_conciliados_path = os.path.join(os.path.dirname(save_path), 'no_conciliados.xlsx')
    
    if save_path:
        try:
            df_no_conciliados.to_excel(no_conciliados_path, index=False)
            log_message(log_widget, f"Archivo de no conciliados guardado exitosamente en {no_conciliados_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo de no conciliados: {e}")


def run_application() -> None:
    """Función principal para ejecutar la aplicación."""
    root: tk.Tk = tk.Tk()
    root.title("Conciliación de Cobranzas")
    root.geometry("850x500")

    log_widget: scrolledtext.ScrolledText = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=10, bg="#ffffff", fg="#333333", font=("Helvetica", 10))
    log_widget.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    frame: tk.Frame = tk.Frame(root)
    frame.pack(pady=10, anchor='w')  # Alinea a la izquierda con anchor='w'

    mp_label: tk.Label = tk.Label(frame, text="Mercado Pago:")
    mp_label.grid(row=0, column=0, sticky="w")
    mp_entry: tk.Entry = tk.Entry(frame, width=50)
    mp_entry.grid(row=0, column=1, padx=10)
    mp_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(mp_entry, log_widget, "Mercado Pago"))
    mp_button.grid(row=0, column=2)

    planilla_label: tk.Label = tk.Label(frame, text="Planilla 1:")
    planilla_label.grid(row=1, column=0, sticky="w")
    planilla_entry: tk.Entry = tk.Entry(frame, width=50)
    planilla_entry.grid(row=1, column=1, padx=10)
    planilla_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(planilla_entry, log_widget, "Planilla 1"))
    planilla_button.grid(row=1, column=2)

    cob_km_label: tk.Label = tk.Label(frame, text="Cobranzas Electrónicas KM1151:")
    cob_km_label.grid(row=2, column=0, sticky="w")
    cob_km_entry: tk.Entry = tk.Entry(frame, width=50)
    cob_km_entry.grid(row=2, column=1, padx=10)
    cob_km_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(cob_km_entry, log_widget, "Cobranzas Electrónicas KM1151"))
    cob_km_button.grid(row=2, column=2)

    cob_bovedas_label: tk.Label = tk.Label(frame, text="Cobranzas Electrónicas LAS BOVEDAS:")
    cob_bovedas_label.grid(row=3, column=0, sticky="w")
    cob_bovedas_entry: tk.Entry = tk.Entry(frame, width=50)
    cob_bovedas_entry.grid(row=3, column=1, padx=10)
    cob_bovedas_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(cob_bovedas_entry, log_widget, "Cobranzas Electrónicas LAS BOVEDAS"))
    cob_bovedas_button.grid(row=3, column=2)

    process_button: tk.Button = tk.Button(root, text="Ejecutar Proceso", command=lambda: run_process(mp_entry, planilla_entry, cob_km_entry, cob_bovedas_entry, log_widget), bg="#4CAF50", fg="white", font=("Helvetica", 12))
    process_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    run_application()

