#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════════════
PROYECTO QUINA - AUTOMATIZACIÓN DE FACTURACIÓN (BILLING)
═══════════════════════════════════════════════════════════════════════════
Autor: Ingeniería de Datos - Proyecto Quina
Fecha: Febrero 2026
Versión: 1.0

PROPÓSITO:
    Script Python portable para automatizar el cálculo de interacciones
    cobrables del chatbot Quina, aplicando las 3 reglas de negocio críticas.

REGLAS IMPLEMENTADAS:
    1. Ventana de 24 horas ESTRICTA (>= 24.0 horas exactas)
    2. Clasificación "Evalúa tu crédito" (Mesa Comercial vs. Ambas Mesas)
    3. Descuento por mensajes de agentes humanos

REQUISITOS:
    - Python 3.7+
    - pandas
    - openpyxl

COMPILAR A EXE (Portable):
    pip install pyinstaller
    pyinstaller --onefile --name QuinaBilling quina_billing_automation.py
    
USO:
    python quina_billing_automation.py --mes "04. Abril"
    python quina_billing_automation.py --mes "05. Mayo" --output "C:\Resultados"
═══════════════════════════════════════════════════════════════════════════
"""

import pandas as pd
import numpy as np
import os
import sys
import glob
import argparse
from datetime import datetime, timedelta
from pathlib import Path


# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════════════════
RUTA_BASE = r"C:\Users\A365\Documents\QuinaQuery\2025"
FRASE_TRIGGER_CREDITO = "podrás evaluar si tienes un crédito"
THRESHOLD_AMBAS_MESAS = 7


# ═══════════════════════════════════════════════════════════════════════════
# FUNCIÓN 1: REGLA VENTANA 24 HORAS ESTRICTA
# ═══════════════════════════════════════════════════════════════════════════
def calcular_ventana_24h(ruta_rdc):
    """
    Aplica la lógica estricta de ventana de 24 horas sobre el RDC.
    
    Args:
        ruta_rdc (str): Ruta completa al archivo RDC Excel
        
    Returns:
        pd.DataFrame: DataFrame con columnas adicionales de cobrabilidad
    """
    print(f"[INFO] Cargando RDC desde: {ruta_rdc}")
    
    # Cargar archivo Excel
    df_rdc = pd.read_excel(ruta_rdc)
    
    # Renombrar columnas si es necesario (ajustar según estructura real)
    # df_rdc.rename(columns={'ID Usuario': 'ID'}, inplace=True)
    
    # Filtrar registros válidos
    df_rdc = df_rdc[df_rdc['ID'].notna() & df_rdc['F_Inicio_Chat'].notna()].copy()
    
    # Asegurar tipo datetime
    df_rdc['F_Inicio_Chat'] = pd.to_datetime(df_rdc['F_Inicio_Chat'])
    
    # Ordenar por usuario y fecha
    df_rdc = df_rdc.sort_values(['ID', 'F_Inicio_Chat']).reset_index(drop=True)
    
    # Calcular F_Inicio_Chat_Anterior (LAG en SQL)
    df_rdc['F_Inicio_Chat_Anterior'] = df_rdc.groupby('ID')['F_Inicio_Chat'].shift(1)
    
    # Calcular diferencia en horas con PRECISIÓN DECIMAL
    df_rdc['Diferencia_Horas_Exactas'] = (
        df_rdc['F_Inicio_Chat'] - df_rdc['F_Inicio_Chat_Anterior']
    ).dt.total_seconds() / 3600.0
    
    # Aplicar regla de cobrabilidad ESTRICTA
    def clasificar_cobrabilidad(row):
        if pd.isna(row['F_Inicio_Chat_Anterior']):
            return 'COBRABLE - Primera conversación'
        elif row['Diferencia_Horas_Exactas'] >= 24.0:
            return 'COBRABLE - Nueva ventana (>= 24h)'
        else:
            return 'NO COBRABLE - Misma ventana (<24h)'
    
    df_rdc['Estado_Cobrabilidad'] = df_rdc.apply(clasificar_cobrabilidad, axis=1)
    
    # Agregar flag binario
    df_rdc['Es_Cobrable'] = df_rdc['Estado_Cobrabilidad'].str.startswith('COBRABLE').astype(int)
    
    print(f"[OK] Procesados {len(df_rdc)} chats")
    print(f"     Cobrables: {df_rdc['Es_Cobrable'].sum()}")
    print(f"     No cobrables: {len(df_rdc) - df_rdc['Es_Cobrable'].sum()}")
    
    return df_rdc


# ═══════════════════════════════════════════════════════════════════════════
# FUNCIÓN 2: REGLA EVALÚA TU CRÉDITO
# ═══════════════════════════════════════════════════════════════════════════
def clasificar_evalua_credito(ruta_ddc_list, ruta_rdc):
    """
    Clasifica conversaciones de evaluación de crédito.
    
    Args:
        ruta_ddc_list (list): Lista de rutas a archivos DDC
        ruta_rdc (str): Ruta al archivo RDC
        
    Returns:
        pd.DataFrame: DataFrame con clasificación de mesa
    """
    print(f"[INFO] Procesando regla 'Evalúa tu crédito'...")
    
    # Cargar RDC para obtener tipificaciones
    df_rdc = pd.read_excel(ruta_rdc)
    df_rdc = df_rdc[['ID_Chat', 'Tipificación_Chat']].copy()
    
    # Cargar y combinar todos los DDC
    df_ddc_list = []
    for ruta in ruta_ddc_list:
        print(f"      Cargando: {Path(ruta).name}")
        df_temp = pd.read_excel(ruta)
        df_ddc_list.append(df_temp)
    
    df_ddc = pd.concat(df_ddc_list, ignore_index=True)
    
    # Unir con RDC para tener tipificación
    df_ddc = df_ddc.merge(df_rdc, on='ID_Chat', how='left')
    
    # Filtrar solo evaluaciones de crédito
    mask_credito = (
        df_ddc['Tipificación_Chat'].str.lower().str.contains('evalua', na=False) |
        df_ddc['Tipificación_Chat'].str.lower().str.contains('crédito', na=False) |
        df_ddc['Tipificación_Chat'].str.lower().str.contains('credito', na=False) |
        df_ddc['Tipificación_Chat'].str.lower().str.contains('opción 3', na=False) |
        df_ddc['Tipificación_Chat'].str.lower().str.contains('opcion 3', na=False)
    )
    
    df_evalua = df_ddc[mask_credito].copy()
    
    # Asegurar tipo datetime
    df_evalua['Fecha_Hora'] = pd.to_datetime(df_evalua['Fecha_Hora'])
    
    # Ordenar por chat y fecha
    df_evalua = df_evalua.sort_values(['ID_Chat', 'Fecha_Hora']).reset_index(drop=True)
    
    # Marcar mensaje trigger
    df_evalua['Es_Mensaje_Trigger'] = df_evalua['Mensaje'].str.lower().str.contains(
        FRASE_TRIGGER_CREDITO.lower(), na=False
    )
    
    # Agrupar por chat y calcular
    def clasificar_chat(grupo):
        total_mensajes = len(grupo)
        
        # Buscar índice del trigger
        trigger_rows = grupo[grupo['Es_Mensaje_Trigger']]
        
        if len(trigger_rows) == 0:
            return pd.Series({
                'Total_Mensajes': total_mensajes,
                'Indice_Trigger': None,
                'Mensajes_Antes_Trigger': total_mensajes,
                'Clasificacion_Mesa': 'SIN TRIGGER - Validar manualmente'
            })
        
        indice_trigger = trigger_rows.index[0]
        mensajes_antes = grupo.index.get_loc(grupo.index[0])
        mensajes_antes_trigger = len(grupo.loc[:indice_trigger]) - 1
        
        # Aplicar regla de clasificación
        if (total_mensajes - mensajes_antes_trigger) > THRESHOLD_AMBAS_MESAS:
            clasificacion = "Cobrar en ambas mesas"
        else:
            clasificacion = "Cobrar Mesa Comercial"
        
        return pd.Series({
            'Total_Mensajes': total_mensajes,
            'Indice_Trigger': mensajes_antes_trigger,
            'Mensajes_Antes_Trigger': mensajes_antes_trigger,
            'Clasificacion_Mesa': clasificacion
        })
    
    resultado_resumen = df_evalua.groupby('ID_Chat').apply(clasificar_chat).reset_index()
    
    # Unir resumen con detalle
    df_evalua = df_evalua.merge(resultado_resumen, on='ID_Chat', how='left')
    
    print(f"[OK] Procesados {len(resultado_resumen)} chats de evaluación de crédito")
    print(f"     Ambas mesas: {(resultado_resumen['Clasificacion_Mesa'] == 'Cobrar en ambas mesas').sum()}")
    print(f"     Mesa comercial: {(resultado_resumen['Clasificacion_Mesa'] == 'Cobrar Mesa Comercial').sum()}")
    
    return df_evalua


# ═══════════════════════════════════════════════════════════════════════════
# FUNCIÓN 3: DESCUENTO POR AGENTES
# ═══════════════════════════════════════════════════════════════════════════
def calcular_descuento_agentes(ruta_ddc_list):
    """
    Identifica y etiqueta mensajes de bot vs. agente.
    
    Args:
        ruta_ddc_list (list): Lista de rutas a archivos DDC
        
    Returns:
        pd.DataFrame: DataFrame con etiquetas de categoría
    """
    print(f"[INFO] Procesando regla 'Descuento por agentes'...")
    
    # Cargar y combinar todos los DDC
    df_ddc_list = []
    for ruta in ruta_ddc_list:
        print(f"      Cargando: {Path(ruta).name}")
        df_temp = pd.read_excel(ruta)
        df_ddc_list.append(df_temp)
    
    df_ddc = pd.concat(df_ddc_list, ignore_index=True)
    
    # Filtrar válidos
    df_ddc = df_ddc[df_ddc['ID_Chat'].notna() & df_ddc['Fecha_Hora'].notna()].copy()
    
    # Asegurar tipo datetime
    df_ddc['Fecha_Hora'] = pd.to_datetime(df_ddc['Fecha_Hora'])
    
    # Ordenar
    df_ddc = df_ddc.sort_values(['ID_Chat', 'Fecha_Hora']).reset_index(drop=True)
    
    # Marcar eventos NOTIFICATION
    df_ddc['Es_Notification'] = df_ddc['Tipo'].str.upper().str.strip() == 'NOTIFICATION'
    
    # Encontrar índice del primer NOTIFICATION por chat
    def encontrar_primer_notification(grupo):
        notifications = grupo[grupo['Es_Notification']]
        if len(notifications) == 0:
            return None
        return notifications.index[0]
    
    primer_notification = df_ddc.groupby('ID_Chat').apply(encontrar_primer_notification)
    primer_notification_dict = primer_notification.to_dict()
    
    # Aplicar etiquetado
    def etiquetar_mensaje(row):
        id_chat = row['ID_Chat']
        indice_actual = row.name
        indice_notification = primer_notification_dict.get(id_chat)
        
        if indice_notification is None:
            return "Conteo Normal - Sin pase a agente"
        elif indice_actual < indice_notification:
            return "Conteo Normal - Bot"
        elif indice_actual == indice_notification:
            return "Evento Notification - No contar"
        else:
            return "Descontar Mensajes de Agente"
    
    df_ddc['Categoria_Mensaje'] = df_ddc.apply(etiquetar_mensaje, axis=1)
    
    # Flags binarios
    df_ddc['Contar_Como_Bot'] = df_ddc['Categoria_Mensaje'].str.startswith('Conteo Normal').astype(int)
    df_ddc['Descontar_Como_Agente'] = (df_ddc['Categoria_Mensaje'] == 'Descontar Mensajes de Agente').astype(int)
    
    # Resumen por chat
    resumen = df_ddc.groupby('ID_Chat').agg({
        'Contar_Como_Bot': 'sum',
        'Descontar_Como_Agente': 'sum'
    }).rename(columns={
        'Contar_Como_Bot': 'Mensajes_Bot',
        'Descontar_Como_Agente': 'Mensajes_Agente'
    }).reset_index()
    
    print(f"[OK] Procesados {len(df_ddc)} mensajes en {len(resumen)} chats")
    print(f"     Mensajes Bot: {df_ddc['Contar_Como_Bot'].sum()}")
    print(f"     Mensajes Agente: {df_ddc['Descontar_Como_Agente'].sum()}")
    
    return df_ddc, resumen


# ═══════════════════════════════════════════════════════════════════════════
# MAIN: ORQUESTADOR PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(
        description='Automatización de facturación Proyecto Quina'
    )
    parser.add_argument(
        '--mes',
        type=str,
        required=True,
        help='Nombre de la carpeta del mes (ej: "04. Abril")'
    )
    parser.add_argument(
        '--output',
        type=str,
        default=None,
        help='Carpeta de salida para resultados (default: misma carpeta del mes)'
    )
    
    args = parser.parse_args()
    
    # Construir rutas
    ruta_mes = os.path.join(RUTA_BASE, args.mes)
    
    if not os.path.exists(ruta_mes):
        print(f"[ERROR] No existe la ruta: {ruta_mes}")
        sys.exit(1)
    
    # Buscar archivos
    archivos_rdc = glob.glob(os.path.join(ruta_mes, "RDC_*.xlsx"))
    archivos_ddc = glob.glob(os.path.join(ruta_mes, "DDC_*.xlsx"))
    
    if not archivos_rdc:
        print(f"[ERROR] No se encontraron archivos RDC en {ruta_mes}")
        sys.exit(1)
    
    if not archivos_ddc:
        print(f"[ERROR] No se encontraron archivos DDC en {ruta_mes}")
        sys.exit(1)
    
    print("═" * 70)
    print(f"PROYECTO QUINA - AUTOMATIZACIÓN DE FACTURACIÓN")
    print(f"Mes: {args.mes}")
    print(f"Archivos RDC: {len(archivos_rdc)}")
    print(f"Archivos DDC: {len(archivos_ddc)}")
    print("═" * 70)
    print()
    
    # Definir carpeta de salida
    if args.output:
        ruta_salida = args.output
    else:
        ruta_salida = ruta_mes
    
    os.makedirs(ruta_salida, exist_ok=True)
    
    # ─────────────────────────────────────────────────────────────────────
    # EJECUTAR REGLA 1: VENTANA 24H
    # ─────────────────────────────────────────────────────────────────────
    print("\n[REGLA 1] Calculando ventana de 24 horas...")
    df_rdc_cobrable = calcular_ventana_24h(archivos_rdc[0])
    
    archivo_salida_1 = os.path.join(ruta_salida, f"Resultado_Regla1_Ventana24h.xlsx")
    df_rdc_cobrable.to_excel(archivo_salida_1, index=False)
    print(f"✓ Guardado: {archivo_salida_1}\n")
    
    # ─────────────────────────────────────────────────────────────────────
    # EJECUTAR REGLA 2: EVALÚA TU CRÉDITO
    # ─────────────────────────────────────────────────────────────────────
    print("[REGLA 2] Clasificando evaluaciones de crédito...")
    df_evalua_credito = clasificar_evalua_credito(archivos_ddc, archivos_rdc[0])
    
    archivo_salida_2 = os.path.join(ruta_salida, f"Resultado_Regla2_EvaluaCredito.xlsx")
    df_evalua_credito.to_excel(archivo_salida_2, index=False)
    print(f"✓ Guardado: {archivo_salida_2}\n")
    
    # ─────────────────────────────────────────────────────────────────────
    # EJECUTAR REGLA 3: DESCUENTO AGENTES
    # ─────────────────────────────────────────────────────────────────────
    print("[REGLA 3] Identificando mensajes de agentes...")
    df_agentes_detalle, df_agentes_resumen = calcular_descuento_agentes(archivos_ddc)
    
    archivo_salida_3a = os.path.join(ruta_salida, f"Resultado_Regla3_Agentes_Detalle.xlsx")
    archivo_salida_3b = os.path.join(ruta_salida, f"Resultado_Regla3_Agentes_Resumen.xlsx")
    
    df_agentes_detalle.to_excel(archivo_salida_3a, index=False)
    df_agentes_resumen.to_excel(archivo_salida_3b, index=False)
    
    print(f"✓ Guardado: {archivo_salida_3a}")
    print(f"✓ Guardado: {archivo_salida_3b}\n")
    
    # ─────────────────────────────────────────────────────────────────────
    # RESUMEN FINAL
    # ─────────────────────────────────────────────────────────────────────
    print("═" * 70)
    print("PROCESAMIENTO COMPLETADO EXITOSAMENTE")
    print("═" * 70)
    print(f"Resultados guardados en: {ruta_salida}")
    print()
    print("Archivos generados:")
    print(f"  1. {Path(archivo_salida_1).name}")
    print(f"  2. {Path(archivo_salida_2).name}")
    print(f"  3. {Path(archivo_salida_3a).name}")
    print(f"  4. {Path(archivo_salida_3b).name}")
    print("═" * 70)


if __name__ == "__main__":
    main()
