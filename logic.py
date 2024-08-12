import pandas as pd
import tkinter as tk
from pandas import DataFrame
from datetime import timedelta

def log_message(log_widget, message: str) -> None:
    """Registra mensajes en el widget de log y en la consola."""
    if isinstance(log_widget, tk.Widget):
        log_widget.insert(tk.END, message + "\n")
        log_widget.update_idletasks()  # Asegura que el log se actualiza en tiempo real
        log_widget.see(tk.END)
    print(message)

def inicializar_df_resultado(df_cobranzas: DataFrame) -> DataFrame:
    """Inicializa el DataFrame resultado con las columnas necesarias."""
    df_result = df_cobranzas.copy()
    df_result['Conciliación'] = ''
    df_result['Importe MP'] = None
    df_result['Fecha MP'] = None
    df_result['Diferencia Importe'] = None
    df_result['Diferencia Minutos'] = None
    df_result['Planilla'] = None
    return df_result

def marcar_coincidencias_directas(df_result: DataFrame, df_mercado_pago: DataFrame) -> DataFrame:
    """Marca coincidencias directas en base a transacciones y establece los valores correspondientes."""
    for i, row in df_result.iterrows():
        if row['Origen'] == 'Mercado Pago':
            match = df_mercado_pago[df_mercado_pago['Operación Relacionada'] == row['Transacción']]
            if not match.empty:
                df_result.at[i, 'Importe MP'] = match.iloc[0]['Importe']
                df_result.at[i, 'Fecha MP'] = match.iloc[0]['Fecha de Pago']
                df_result.at[i, 'Diferencia Importe'] = match.iloc[0]['Importe'] - row['Cobrado']
                df_result.at[i, 'Diferencia Minutos'] = (match.iloc[0]['Fecha de Pago'] - row['Fecha']).total_seconds() / 60

                if abs(df_result.at[i, 'Diferencia Importe']) > 1:
                    df_result.at[i, 'Conciliación'] = 'MP - Error en importe'
                elif abs(df_result.at[i, 'Diferencia Minutos']) > 10:
                    df_result.at[i, 'Conciliación'] = 'MP - revisar fecha'
                else:
                    df_result.at[i, 'Conciliación'] = 'MP'

                df_mercado_pago.drop(match.index, inplace=True)
    return df_result

def comparar_fechas_con_tolerancia(df_result: DataFrame, df_mercado_pago: DataFrame, tolerancia: int = 10) -> DataFrame:
    """Compara fechas con una tolerancia de 10 minutos para marcar conciliaciones."""
    for i, row in df_result.iterrows():
        if row['Conciliación'] == '' and row['Origen'] == 'Mercado Pago':
            coincidencias = df_mercado_pago[
                (df_mercado_pago['Operación Relacionada'] == row['Transacción']) &
                (abs(df_mercado_pago['Fecha de Pago'] - row['Fecha']) <= timedelta(minutes=tolerancia))
            ]
            if not coincidencias.empty:
                df_result.at[i, 'Importe MP'] = coincidencias.iloc[0]['Importe']
                df_result.at[i, 'Fecha MP'] = coincidencias.iloc[0]['Fecha de Pago']
                df_result.at[i, 'Diferencia Importe'] = coincidencias.iloc[0]['Importe'] - row['Cobrado']
                df_result.at[i, 'Diferencia Minutos'] = (coincidencias.iloc[0]['Fecha de Pago'] - row['Fecha']).total_seconds() / 60
                df_result.at[i, 'Conciliación'] = 'MP'
                df_mercado_pago.drop(coincidencias.index, inplace=True)
            else:
                posible_match = df_mercado_pago[df_mercado_pago['Operación Relacionada'] == row['Transacción']]
                if not posible_match.empty:
                    df_result.at[i, 'Importe MP'] = posible_match.iloc[0]['Importe']
                    df_result.at[i, 'Fecha MP'] = posible_match.iloc[0]['Fecha de Pago']
                    df_result.at[i, 'Diferencia Importe'] = posible_match.iloc[0]['Importe'] - row['Cobrado']
                    df_result.at[i, 'Diferencia Minutos'] = (posible_match.iloc[0]['Fecha de Pago'] - row['Fecha']).total_seconds() / 60

                    if row['Estado'] == 'Rechazado':
                        df_result.at[i, 'Conciliación'] = 'MP - rechazado'
                    elif abs(df_result.at[i, 'Diferencia Minutos']) > 10:
                        df_result.at[i, 'Conciliación'] = 'MP - revisar fecha'
                    elif abs(df_result.at[i, 'Diferencia Importe']) > 1:
                        df_result.at[i, 'Conciliación'] = 'MP - Error en importe'
    return df_result

def marcar_coincidencias_planilla(df_result: DataFrame, df_planilla: DataFrame) -> DataFrame:
    """Marca coincidencias basadas en la Planilla 1 y establece las conciliaciones correspondientes."""
    for i, row in df_result.iterrows():
        if row['Conciliación'] == '' or 'MP' in row['Conciliación']:
            match = df_planilla[df_planilla['Nro Operación'] == row['Transacción']]
            if not match.empty:
                df_result.at[i, 'Importe MP'] = match.iloc[0]['Importe']
                df_result.at[i, 'Planilla'] = match.iloc[0]['Planilla']
                df_result.at[i, 'Diferencia Importe'] = match.iloc[0]['Importe'] - row['Cobrado']
                
                if 'MP' in row['Conciliación']:
                    df_result.at[i, 'Conciliación'] = 'MP - P1 - DUPLICADO!'
                elif abs(df_result.at[i, 'Diferencia Importe']) > 15:
                    df_result.at[i, 'Conciliación'] = 'P1 - revisar monto'
                else:
                    df_result.at[i, 'Conciliación'] = 'P1'
    return df_result

def finalizar_resultado(df_result: DataFrame) -> DataFrame:
    """Completa la conciliación marcando los no conciliados."""
    df_result.loc[df_result['Conciliación'] == '', 'Conciliación'] = 'No Conciliado'
    return df_result

def process_logic(log_widget: tk.Widget, df_cobranzas: DataFrame, df_mercado_pago: DataFrame, df_planilla_1: DataFrame = None) -> tuple[DataFrame, DataFrame]:
    """Proceso principal que coordina la conciliación entre Cobranzas Electrónicas, Mercado Pago y Planilla 1.
    Retorna tanto el resultado conciliado como los no conciliados en Mercado Pago."""
    log_message(log_widget, "Iniciando la conciliación.")
    
    df_result = inicializar_df_resultado(df_cobranzas)
    df_mercado_pago_residuo = df_mercado_pago.copy()

    log_message(log_widget, "Realizando comparación de transacciones.")
    df_result = marcar_coincidencias_directas(df_result, df_mercado_pago_residuo)
    
    log_message(log_widget, "Comparando fechas con tolerancia de 10 minutos.")
    df_result = comparar_fechas_con_tolerancia(df_result, df_mercado_pago_residuo)
    
    if df_planilla_1 is not None:
        log_message(log_widget, "Buscando coincidencias en Planilla 1.")
        df_result = marcar_coincidencias_planilla(df_result, df_planilla_1)
    
    log_message(log_widget, "Finalizando el resultado.")
    df_result = finalizar_resultado(df_result)
    
    log_message(log_widget, "Conciliación finalizada.")
    
    return df_result, df_mercado_pago_residuo

