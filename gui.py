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


def select_output_file(entry_widget: tk.Entry, log_widget: scrolledtext.ScrolledText) -> None:
    """Abre un diálogo para seleccionar la ubicación del archivo de salida."""
    file_path: str = filedialog.asksaveasfilename(title="Guardar archivo conciliado", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)
        log_message(log_widget, f"Archivo de salida seleccionado: {file_path}")


def import_cobranza_electronica(file_path: str, log_widget: scrolledtext.ScrolledText) -> DataFrame | None:
    """Importa y limpia el archivo COBRANZAS ELECTRONICAS desde la fila 7 en adelante."""
    try:
        # Leer el archivo comenzando desde la fila 7
        df: DataFrame = pd.read_excel(file_path, skiprows=6)

        # Limpiar el DataFrame: eliminar columnas completamente vacías
        df = df.dropna(how='all', axis=1)

        # Reiniciar el índice para asegurar que está bien estructurado
        df = df.reset_index(drop=True)

        # Convertir columnas relevantes al tipo adecuado, por ejemplo, la columna de fecha
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')

        log_message(log_widget, "Cobranzas Electrónicas importado y limpiado correctamente.")
        log_message(log_widget, f"Primeras 3 filas de Cobranzas Electrónicas:\n{df.head(3)}")
        log_message(log_widget, f"Descripción del dataframe:\n{df.describe}")
        
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
        log_message(log_widget, f"Primeras 3 filas de Mercado Pago:\n{df.head(3)}")
        log_message(log_widget, f"Descripción del dataframe:\n{df.describe}")

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
        log_message(log_widget, f"Primeras 3 filas de Planilla 1:\n{df.head(3)}")
        log_message(log_widget, f"Descripción del dataframe:\n{df.describe}")

        return df
    except Exception as e:
        # Mostrar el error exacto en el log
        log_message(log_widget, f"Error al importar y limpiar Planilla 1: {e}")
        messagebox.showerror("Error", f"Error al importar y limpiar Planilla 1: {e}")
        return None


def import_files(cob_entry: tk.Entry, mp_entry: tk.Entry, planilla_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> tuple[DataFrame | None, DataFrame | None, DataFrame | None]:
    """Importa los archivos XLSX seleccionados por el usuario y devuelve DataFrames correspondientes."""
    log_message(log_widget, "Importando archivos...")

    file1: str = cob_entry.get()
    file2: str = mp_entry.get()
    file3: str = planilla_entry.get()

    if file1 and file2:
        try:
            df1: DataFrame = import_cobranza_electronica(file1, log_widget)
            df2: DataFrame = import_mercado_pago(file2, log_widget)
            
            df3: DataFrame | None = None
            if file3:
                df3 = import_planilla_1(file3, log_widget)
            
            return df1, df2, df3
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar archivos: {e}")
            return None, None, None


def generate_output(result_df: DataFrame, output_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> None:
    """Guarda el DataFrame resultante en un archivo XLSX en la ubicación seleccionada por el usuario."""
    save_path: str = output_entry.get()
    
    if save_path:
        try:
            result_df.to_excel(save_path, index=False)
            log_message(log_widget, f"Archivo conciliado guardado exitosamente en {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el archivo: {e}")


def run_process(cob_entry: tk.Entry, mp_entry: tk.Entry, planilla_entry: tk.Entry, output_entry: tk.Entry, log_widget: scrolledtext.ScrolledText) -> None:
    """
    Ejecuta el flujo principal de la aplicación:
    1. Importa los archivos seleccionados por el usuario como DataFrames.
    2. Llama a la lógica principal para procesar los DataFrames.
    3. Genera un archivo XLSX con el DataFrame resultante.
    """
    df1, df2, df3 = import_files(cob_entry, mp_entry, planilla_entry, log_widget)
    if df1 is not None and df2 is not None:
        log_message(log_widget, "Iniciando procesado de ficheros...")
        
        result_df: DataFrame = process_logic(log_widget, df1, df2, df3)
        
        generate_output(result_df, output_entry, log_widget)
        log_message(log_widget, "Proceso finalizado.")


def run_application() -> None:
    """Función principal para ejecutar la aplicación."""
    root: tk.Tk = tk.Tk()
    root.title("Conciliación de Cobranzas")
    root.geometry("770x450")

    log_widget: scrolledtext.ScrolledText = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=10, bg="#ffffff", fg="#333333", font=("Helvetica", 10))
    log_widget.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    frame: tk.Frame = tk.Frame(root)
    frame.pack(pady=10)

    cob_label: tk.Label = tk.Label(frame, text="Cobranzas Electrónicas:")
    cob_label.grid(row=0, column=0, sticky="e")
    cob_entry: tk.Entry = tk.Entry(frame, width=50)
    cob_entry.grid(row=0, column=1, padx=10)
    cob_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(cob_entry, log_widget, "Cobranzas Electrónicas"))
    cob_button.grid(row=0, column=2)

    mp_label: tk.Label = tk.Label(frame, text="Mercado Pago:")
    mp_label.grid(row=1, column=0, sticky="e")
    mp_entry: tk.Entry = tk.Entry(frame, width=50)
    mp_entry.grid(row=1, column=1, padx=10)
    mp_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(mp_entry, log_widget, "Mercado Pago"))
    mp_button.grid(row=1, column=2)

    planilla_label: tk.Label = tk.Label(frame, text="Planilla 1:")
    planilla_label.grid(row=2, column=0, sticky="e")
    planilla_entry: tk.Entry = tk.Entry(frame, width=50)
    planilla_entry.grid(row=2, column=1, padx=10)
    planilla_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_file(planilla_entry, log_widget, "Planilla 1"))
    planilla_button.grid(row=2, column=2)

    output_label: tk.Label = tk.Label(frame, text="Archivo de salida:")
    output_label.grid(row=3, column=0, sticky="e")

    default_output_path: str = os.path.join(os.getcwd(), "conciliacion.xlsx")
    output_entry: tk.Entry = tk.Entry(frame, width=50)
    output_entry.grid(row=3, column=1, padx=10)
    output_entry.insert(0, default_output_path) 
    output_button: tk.Button = tk.Button(frame, text="Seleccionar", command=lambda: select_output_file(output_entry, log_widget))
    output_button.grid(row=3, column=2)

    process_button: tk.Button = tk.Button(root, text="Ejecutar Proceso", command=lambda: run_process(cob_entry, mp_entry, planilla_entry, output_entry, log_widget), bg="#4CAF50", fg="white", font=("Helvetica", 12))
    process_button.pack(pady=20)

    root.mainloop()



if __name__ == "__main__":
    run_application()
