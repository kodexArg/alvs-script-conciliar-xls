import pandas as pd
import tkinter as tk
from pandas import DataFrame
from datetime import timedelta


def log_message(log_widget, message: str) -> None:
    """Registra mensajes en el widget de log y en la consola."""
    log_widget.insert(tk.END, message + "\n")
    log_widget.update_idletasks()
    log_widget.see(tk.END)
    print(message)


def inicializar_df_resultado(df_mercado_pago: DataFrame) -> DataFrame:
    """Inicializa el DataFrame resultado con las columnas necesarias basadas en Mercado Pago y filtra operaciones relevantes."""
    operaciones_relevantes = ['Cobro', 'Ingreso de dinero', 'Dinero recibido']
    df_result = df_mercado_pago[df_mercado_pago['Tipo de Operación'].isin(operaciones_relevantes)].copy()

    df_result['Conciliación'] = ''
    df_result['Importe C.E.'] = None
    df_result['Fecha C.E.'] = None
    df_result['Dif Importe'] = None
    df_result['Dif Minutos'] = None
    df_result['Planilla'] = None
    
    return df_result


def marcar_coincidencias_cobranzas(df_result: DataFrame, df_cobranzas: DataFrame, conciliacion_label: str, tolerancia: int = 10) -> DataFrame:
    """Marca coincidencias basadas en Cobranzas Electrónicas y establece las conciliaciones correspondientes."""
    df_cobranzas['Conciliado'] = False  # Inicializa la columna 'Conciliado' como False

    for i, row in df_result.iterrows():
        match = df_cobranzas[df_cobranzas['Transacción'] == row['Operación Relacionada']]
        if not match.empty:
            df_result.at[i, 'Importe C.E.'] = match.iloc[0]['Cobrado']
            df_result.at[i, 'Fecha C.E.'] = match.iloc[0]['Fecha']
            df_result.at[i, 'Dif Importe'] = row['Importe'] - match.iloc[0]['Cobrado']
            df_result.at[i, 'Dif Minutos'] = (row['Fecha de Pago'] - match.iloc[0]['Fecha']).total_seconds() / 60
            
            if abs(df_result.at[i, 'Dif Importe']) > 1:
                df_result.at[i, 'Conciliación'] = f'{conciliacion_label} - revisar importe'
            elif abs(df_result.at[i, 'Dif Minutos']) > tolerancia:
                df_result.at[i, 'Conciliación'] = f'{conciliacion_label} - revisar fecha'
            else:
                df_result.at[i, 'Conciliación'] = conciliacion_label

            # Marca la fila correspondiente en df_cobranzas como conciliada
            df_cobranzas.at[match.index[0], 'Conciliado'] = True

    return df_result


def marcar_coincidencias_planilla(df_result: DataFrame, df_planilla: DataFrame) -> DataFrame:
    """Marca coincidencias basadas en la Planilla 1 y establece las conciliaciones correspondientes."""
    df_planilla['Conciliado'] = False  # Inicializa la columna 'Conciliado' como False

    for i, row in df_result.iterrows():
        if row['Conciliación'] == '':
            match = df_planilla[df_planilla['Nro Operación'] == row['Operación Relacionada']]
            if not match.empty:
                df_result.at[i, 'Importe C.E.'] = match.iloc[0]['Importe']
                df_result.at[i, 'Planilla'] = match.iloc[0]['Planilla']
                df_result.at[i, 'Dif Importe'] = row['Importe'] - match.iloc[0]['Importe']

                if abs(df_result.at[i, 'Dif Importe']) > 15:
                    df_result.at[i, 'Conciliación'] = 'Planilla 1 - revisar monto'
                else:
                    df_result.at[i, 'Conciliación'] = 'Planilla 1'

                # Marca la fila correspondiente en df_planilla como conciliada
                df_planilla.at[match.index[0], 'Conciliado'] = True

    return df_result


def finalizar_resultado(df_result: DataFrame) -> DataFrame:
    """Completa la conciliación marcando los no conciliados."""
    df_result.loc[df_result['Conciliación'] == '', 'Conciliación'] = 'No Conciliado'
    return df_result


def extract_non_conciliated(df_cobranzas: DataFrame) -> DataFrame:
    """Extrae registros de Cobranzas Electrónicas que no fueron conciliados."""
    non_conciliados = df_cobranzas[df_cobranzas['Conciliado'] == False].copy()
    return non_conciliados


def process_logic(log_widget: tk.Widget, df_mercado_pago: DataFrame, df_planilla_1: DataFrame, df_km1151: DataFrame, df_las_bovedas: DataFrame) -> tuple[DataFrame, DataFrame, DataFrame]:
    """Proceso principal que coordina la conciliación entre Mercado Pago, Cobranzas Electrónicas y Planilla 1."""
    log_message(log_widget, "Iniciando la conciliación.")
    
    df_result = inicializar_df_resultado(df_mercado_pago)

    # Paso 1: conciliar Cobranzas KM1151
    log_message(log_widget, "Conciliando con Cobranzas Electrónicas KM1151.")
    df_result = marcar_coincidencias_cobranzas(df_result, df_km1151, 'Cobranzas KM1151')
    
    # Paso 2: conciliar Cobranzas Las Bóvedas
    log_message(log_widget, "Conciliando con Cobranzas Electrónicas LAS BOVEDAS.")
    df_result = marcar_coincidencias_cobranzas(df_result, df_las_bovedas, 'Cobranzas Las Bovedas')
    
    # Paso 3: conciliar con Planilla 1
    log_message(log_widget, "Conciliando con Planilla 1.")
    df_result = marcar_coincidencias_planilla(df_result, df_planilla_1)
    
    log_message(log_widget, "Finalizando el resultado.")
    df_result = finalizar_resultado(df_result)
    
    log_message(log_widget, "Extrayendo registros no conciliados.")
    df_km1151_residuo = extract_non_conciliated(df_km1151)
    df_bovedas_residuo = extract_non_conciliated(df_las_bovedas)

    log_message(log_widget, "Conciliación finalizada.")
    
    return df_result, df_km1151_residuo, df_bovedas_residuo
