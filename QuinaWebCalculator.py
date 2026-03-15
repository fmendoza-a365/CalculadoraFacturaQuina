import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from QuinaLogic import QuinaCalculator

st.set_page_config(page_title="Facturación Quina", page_icon="💼", layout="wide")

st.title("📋 Calculadora de Facturación - Quina")
st.markdown("""
**📁 Instrucciones:** Sube los archivos RDC y DDC mensuales para generar la factura.  
Se aplicarán automáticamente las reglas de ventana 24h, descuentos por agente, crédito y la **regla de 7 mensajes (Doble Mesa)**.
""")

# Carga de archivos
st.sidebar.header("📂 Archivos de Entrada")

file_rdc = st.sidebar.file_uploader("Subir Archivo RDC (Resumen)", type=["xlsx", "csv"])
files_ddc = st.sidebar.file_uploader("Subir Archivos DDC (Detalle)", type=["xlsx", "csv"], accept_multiple_files=True)

# Procesamiento de facturación

# La lógica de generación de Excel ahora se maneja en QuinaLogic.py
# para evitar duplicidad de código.

# Botón de procesamiento
if st.sidebar.button("⚙️ PROCESAR FACTURA", type="primary"):
    if not file_rdc or not files_ddc:
        st.error("⚠️ Error: Debes subir ambos archivos (RDC y DDC) para continuar.")
    else:
        status_container = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # Instanciar calculadora centralizada
            calc = QuinaCalculator()
            
            # Paso 1: Procesamiento RDC
            status_container.info("⏳ Paso 1/3: Preparando datos...")
            progress_bar.progress(20)
            
            if file_rdc.name.lower().endswith('.csv'):
                df_rdc = pd.read_csv(file_rdc, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_rdc = pd.read_excel(file_rdc)
            
            # Paso 2: Procesamiento DDC
            status_container.info("⏳ Paso 2/3: Procesando conversaciones y mensajes...")
            progress_bar.progress(50)
            
            df_ddc_list = []
            for f in files_ddc:
                if f.name.lower().endswith('.csv'):
                    df_ddc_list.append(pd.read_csv(f, sep=None, engine='python', encoding='utf-8-sig'))
                else:
                    df_ddc_list.append(pd.read_excel(f))
            
            # Ejecutar lógica centralizada
            summary = calc.process_data(df_rdc, df_ddc_list)
            
            # Resultados finales
            progress_bar.progress(100)
            status_container.success("✅ Cálculo completado exitosamente")
            
            # Tarjetas de KPI
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(label="HSM Bruto", value=f"{summary['HSM Bruto']:.0f}", delta=f"- {summary['HSM Credito']} (Crédito)")
            with col2:
                st.metric(label="Q HSM (Final Facturable)", value=f"{summary['Total HSM Final']:.0f}", delta="- 1,000 (Meta)")
            with col3:
                st.metric(label="Q Mensajes (Facturables)", value=f"{summary['Total Mensajes Final']:.0f}")
            
            # Descarga de reporte
            st.markdown("---")
            st.subheader("📥 Descargar Reporte")
            
            # Generar Excel usando la lógica centralizada
            excel_data = calc.generate_excel_report()
            
            st.download_button(
                label="📄 Descargar FACTURA_FINAL.xlsx",
                data=excel_data,
                file_name="FACTURA_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
        except Exception as e:
            status_container.error(f"❌ Error en el procesamiento: {str(e)}")

# Información Footer
st.sidebar.markdown("---")
st.sidebar.info("v1.1 - Calculadora Web (Lógica Centralizada)")
