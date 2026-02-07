import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Configuración de Página
st.set_page_config(page_title="Facturación Quina", page_icon="💰", layout="wide")

# Título y Descripción
st.title("🤖 Calculadora de Facturación Automática")
st.markdown("""
Sube tus archivos mensuales (**RDC** y **DDC**) para generar la factura oficial.
Esta app procesa las reglas de negocio, descuentos de agentes y ventanas de 24h.
""")

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR: CARGA DE ARCHIVOS
# ═══════════════════════════════════════════════════════════════════════════
st.sidebar.header("📂 1. Carga de Archivos")

file_rdc = st.sidebar.file_uploader("Subir Archivo RDC (Resumen)", type=["xlsx"])
files_ddc = st.sidebar.file_uploader("Subir Archivos DDC (Detalle)", type=["xlsx"], accept_multiple_files=True)

# ═══════════════════════════════════════════════════════════════════════════
# LÓGICA DE PROCESAMIENTO
# ═══════════════════════════════════════════════════════════════════════════

def get_excel_bytes(q_hsm, q_mensajes, hsm_bruto, hsm_credito, mensajes_bruto, mensajes_agente, mensajes_credito, df_detalle):
    """Genera el Excel de factura en memoria y devuelve bytes"""
    output = io.BytesIO()
    
    # Tarifas
    FEE_MENSUAL = 760.00
    TARIFA_HSM = 0.077
    META_FREE_TIER = 1000
    
    # TARIFAS ESCALONADAS PARA MENSAJES (según tabla real)
    def calcular_costo_mensajes(cantidad):
        """Calcula el costo total aplicando tarifas escalonadas"""
        if cantidad <= 0:
            return 0.0
        elif cantidad <= 9999:
            return cantidad * 0.0456
        elif cantidad <= 99999:
            return cantidad * 0.0380
        elif cantidad <= 249999:
            return cantidad * 0.0304
        else: # 250000+
            return cantidad * 0.0228
    
    total_hsm_money = q_hsm * TARIFA_HSM
    total_msg_money = calcular_costo_mensajes(q_mensajes)
    subtotal = FEE_MENSUAL + total_hsm_money + total_msg_money
    igv = subtotal * 0.18
    total_facturar = subtotal + igv

    wb = openpyxl.Workbook()
    
    # ------------------ HOJA 1: FACTURA DESGLOSADA ------------------
    ws = wb.active
    ws.title = "Factura"

    # Estilos
    header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    sub_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Determinar tarifa aplicable (Referencial para el texto)
    if q_mensajes <= 9999:
        tarifa_msg_aplicable = 0.0456
    elif q_mensajes <= 99999:
        tarifa_msg_aplicable = 0.0380
    elif q_mensajes <= 249999:
        tarifa_msg_aplicable = 0.0304
    else:
        tarifa_msg_aplicable = 0.0228

    # Estructura Datos EXPANDIDA
    data = [
        ["CONCEPTOS", "ABR / CANTIDAD", "MONTO S/", "OBSERVACIONES"], # Header
        ["Fee Mensual", 1, FEE_MENSUAL, "Fee Mensual Broker Whatsapp API Oficial"],
        ["", "", "", ""],  # Espacio
        ["CÁLCULO HSM (Detallado)", "", "", ""],  # Sección
        ["HSM Bruto (Total Conversaciones 24h)", hsm_bruto, "", "Antes de descuentos"],
        ["(-) HSM Opción 3 (Evalúa tu Crédito)", -hsm_credito, "", "Sesiones que derivaron a crédito"],
        ["(-) HSM Meta Free Tier", -META_FREE_TIER, "", "1,000 conversaciones gratuitas Meta"],
        ["Q HSM Neto Facturable", q_hsm, "Calculado", "HSM a cobrar después de descuentos"],
        ["Tarifa por HSM", TARIFA_HSM, "Tarifa", "Según adenda N° 2"],
        ["TOTAL HSM", "", total_hsm_money, ""],
        ["", "", "", ""],  # Espacio
        ["CÁLCULO MENSAJES (Detallado)", "", "", ""],  # Sección
        ["Mensajes Bruto (Total)", mensajes_bruto, "", "Todos los mensajes del periodo"],
        ["(-) Mensajes Post-Agente", -mensajes_agente, "", "Mensajes después de pase a humano"],
        ["(-) Mensajes Post-Crédito", -mensajes_credito, "", "Mensajes después de trigger crédito"],
        ["Q Mensajes Neto Facturable", q_mensajes, "Calculado", "Mensajes a cobrar después de descuentos"],
        ["Tarifa por mensajes", tarifa_msg_aplicable, "Tarifa", f"Tarifa escalonada aplicada al volumen"],
        ["TOTAL MENSAJES", "", total_msg_money, ""],
        ["", "", "", ""],  # Espacio
        ["SUB TOTAL", "", subtotal, ""],
        ["IGV (18%)", "", igv, ""],
        ["TOTAL A FACTURAR", "", total_facturar, ""]
    ]

    for i, row in enumerate(data, start=1):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = border
            if isinstance(val, (int, float)) and i > 1 and j == 3: cell.number_format = '"S/ " #,##0.00'
            if i == 1:
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal='center')
            # Filas amarillas
            if row[0] in ["Fee Mensual", "TOTAL HSM", "TOTAL MENSAJES"]:
                    if j in [1,2,3]: cell.fill = sub_fill
            # Negritas finales
            if row[0] in ["SUB TOTAL", "TOTAL A FACTURAR"]:
                    cell.font = bold_font
            # Negritas secciones
            if row[0] in ["CÁLCULO HSM (Detallado)", "CÁLCULO MENSAJES (Detallado)"]:
                    cell.font = bold_font

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 50

    # ------------------ HOJA 2: AUDITORÍA DETALLADA ------------------
    if not df_detalle.empty:
        ws2 = wb.create_sheet("Detalle Auditoría")
        
        # Headers Auditoría (EXPANDIDO CON DESGLOSE COMPLETO)
        headers = [
            "ID Chat", 
            "Fecha (Día)", 
            "F.Inicio (RDC)",
            "Tipificación Chat",
            "Es HSM Bruto? (1=Sí)", 
            "Tuvo Crédito? (1=Sí)",
            "Mensajes Bruto",
            "(-) Mensajes Post-Agente",
            "(-) Mensajes Post-Crédito",
            "Mensajes Facturables (Neto)",
            "Fecha Corte Agente", 
            "Fecha Corte Crédito"
        ]
        
        for col, h in enumerate(headers, 1):
            cell = ws2.cell(row=1, column=col, value=h)
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal='center')
        
        # Datos
        rows = df_detalle.values.tolist()
        for i, row_data in enumerate(rows, start=2):
            for j, val in enumerate(row_data, 1):
                cell = ws2.cell(row=i, column=j, value=val)
                # Forzar Texto para ID Chat (Columna 1)
                if j == 1:
                    cell.number_format = '@'  # Formato Texto explícito en Excel
        
        # Anchos de columna optimizados
        ws2.column_dimensions['A'].width = 25  # ID Chat
        ws2.column_dimensions['B'].width = 12  # Fecha
        ws2.column_dimensions['C'].width = 20  # F.Inicio
        ws2.column_dimensions['D'].width = 30  # Tipificación Chat
        ws2.column_dimensions['E'].width = 18  # Es HSM Bruto
        ws2.column_dimensions['F'].width = 18  # Tuvo Crédito
        ws2.column_dimensions['G'].width = 15  # Mensajes Bruto
        ws2.column_dimensions['H'].width = 22  # Mensajes Post-Agente
        ws2.column_dimensions['I'].width = 22  # Mensajes Post-Crédito
        ws2.column_dimensions['J'].width = 20  # Mensajes Facturables
        ws2.column_dimensions['K'].width = 20  # Fecha Corte Agente
        ws2.column_dimensions['L'].width = 20  # Fecha Corte Crédito

    wb.save(output)
    return output.getvalue()

# Botón Principal
if st.sidebar.button("🚀 PROCESAR FACTURA", type="primary"):
    if not file_rdc or not files_ddc:
        st.error("⚠️ Por favor sube AMBOS archivos (RDC y DDC) para continuar.")
    else:
        status_container = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # 1. RDC
            status_container.info("⏳ Paso 1/3: Procesando Archivo RDC (Regla 24h)...")
            progress_bar.progress(20)
            
            df_rdc = pd.read_excel(file_rdc, usecols=["ID", "F.Inicio Chat", "ID Chat", "Tipificación Chat"])
            df_rdc.dropna(subset=["ID", "F.Inicio Chat"], inplace=True)
            df_rdc["F.Inicio Chat"] = pd.to_datetime(df_rdc["F.Inicio Chat"])
            df_rdc.sort_values(by=["ID", "F.Inicio Chat"], inplace=True)
            
            # Cálculo Vectorizado 24h
            df_rdc["Prev_ID"] = df_rdc["ID"].shift(1)
            df_rdc["Prev_Time"] = df_rdc["F.Inicio Chat"].shift(1)
            time_diff = (df_rdc["F.Inicio Chat"] - df_rdc["Prev_Time"]).dt.total_seconds() / 3600.0
            
            is_new_id = df_rdc["ID"] != df_rdc["Prev_ID"]
            is_new_window = time_diff >= 24.0
            
            df_rdc["Es_Cobrable"] = (is_new_id | is_new_window).astype(int)
            total_q_hsm = df_rdc["Es_Cobrable"].sum()
            
            # DETECCIÓN DE CRÉDITO (Para HSM): Inmediatamente después de cargar RDC
            # Asegurar ID Chat string
            df_rdc["ID Chat"] = df_rdc["ID Chat"].astype(str)
            
            # Detectar chats con tipificación de crédito (cualquier tipificación que contenga "evalú")
            mask_tipif_credito = df_rdc["Tipificación Chat"].astype(str).str.contains("evalú", case=False, na=False)
            chats_con_credito_tipif = set(df_rdc[mask_tipif_credito]["ID Chat"].unique())
            
            # Marcar filas RDC que son de crédito (por Tipificación)
            df_rdc["Es_Credito"] = df_rdc["ID Chat"].isin(chats_con_credito_tipif)

            # 2. DDC
            status_container.info("⏳ Paso 2/3: Procesando DDC (Mensajes, Agentes, Crédito)...")
            progress_bar.progress(50)
            
            dfs = []
            for f in files_ddc:
                dfs.append(pd.read_excel(f, usecols=["ID Chat", "Mensaje", "Fecha Hora", "Tipo"]))
            
            if dfs:
                df_ddc = pd.concat(dfs, ignore_index=True)
                df_ddc["Fecha Hora"] = pd.to_datetime(df_ddc["Fecha Hora"])
                df_ddc["ID Chat"] = df_ddc["ID Chat"].astype(str)
                df_ddc["Tipo"] = df_ddc["Tipo"].astype(str).str.upper().str.strip()
                df_ddc["Mensaje"] = df_ddc["Mensaje"].astype(str).str.lower()
                
                # Hitos
                status_container.info("⏳ Paso 2/3: Analizando interacciones de Agente y Crédito...")
                
                agente_times = df_ddc[df_ddc["Tipo"] == "NOTIFICATION"].groupby("ID Chat")["Fecha Hora"].min()
                
                # Detectar trigger de crédito (Opción 3 del menú)
                # Buscar variaciones del mensaje de crédito
                credito_mask = (
                    df_ddc["Mensaje"].str.contains("evalúa si tienes un crédito", na=False) |
                    df_ddc["Mensaje"].str.contains("evalua si tienes un credito", na=False) |
                    df_ddc["Mensaje"].str.contains("3. evalúa", na=False) |
                    df_ddc["Mensaje"].str.contains("3. evalua", na=False)
                )
                credito_times = df_ddc[credito_mask].groupby("ID Chat")["Fecha Hora"].min()
                
                df_ddc["Time_Agente"] = df_ddc["ID Chat"].map(agente_times)
                df_ddc["Time_Credito"] = df_ddc["ID Chat"].map(credito_times)
                
                cond_antes_agente = df_ddc["Time_Agente"].isna() | (df_ddc["Fecha Hora"] < df_ddc["Time_Agente"])
                cond_antes_credito = df_ddc["Time_Credito"].isna() | (df_ddc["Fecha Hora"] < df_ddc["Time_Credito"])
                
                df_ddc["Es_Facturable"] = (cond_antes_agente & cond_antes_credito).astype(int)
                total_q_mensajes = df_ddc["Es_Facturable"].sum()
                
                # --- CÁLCULO DESGLOSE MENSAJES (para factura detallada) ---
                mensajes_bruto = len(df_ddc)  # Total de mensajes
                
                # Mensajes post-agente (los que NO son facturables por agente)
                cond_post_agente = ~cond_antes_agente  # Después del agente
                mensajes_agente = cond_post_agente.sum()
                
                # Mensajes post-crédito (los que NO son facturables por crédito)
                # Nota: Algunos pueden estar ya descontados por agente, así que solo contamos los ADICIONALES
                # Es decir, mensajes que pasaron el filtro de agente pero fallaron el de crédito
                cond_post_credito = cond_antes_agente & (~cond_antes_credito)
                mensajes_credito = cond_post_credito.sum()
                
                # --- CÁLCULO HSM: Descontar Crédito y 1K Meta ---
                # La detección de crédito por Tipificación ya se hizo al cargar RDC
                # df_rdc["Es_Credito"] ya está marcado
                
                # Total HSM Inicial (Bruto)
                hsm_bruto = df_rdc["Es_Cobrable"].sum()
                
                # HSM que son de crédito (para restar)
                # Solo contamos como "HSM de Crédito" si era cobrable Y tuvo crédito
                hsm_credito = df_rdc[(df_rdc["Es_Cobrable"] == 1) & (df_rdc["Es_Credito"])].shape[0]
                
                # Cálculo Final HSM
                # Total - Crédito - 1000 (Meta Free Tier)
                total_q_hsm = max(0, hsm_bruto - hsm_credito - 1000)

                # --- PREPARAR MESA DE AUDITORÍA DETALLADA (Actualizado con Metadatos) ---
                # RDC Join Ready
                rdc_view = df_rdc[["ID Chat", "F.Inicio Chat", "Tipificación Chat", "Es_Cobrable", "Es_Credito"]].copy()
                rdc_view["ID Chat"] = rdc_view["ID Chat"].astype(str)
                
                # DDC Join Ready - EXPANDIDO CON DESGLOSE COMPLETO
                # 1. Totales Facturables
                ddc_counts = df_ddc.groupby("ID Chat")["Es_Facturable"].sum().reset_index()
                ddc_counts.rename(columns={"Es_Facturable": "Mensajes_Facturables"}, inplace=True)
                
                # 2. Mensajes Bruto por Chat
                ddc_bruto = df_ddc.groupby("ID Chat").size().reset_index(name="Mensajes_Bruto")
                
                # 3. Mensajes Post-Agente por Chat
                df_ddc["Es_Post_Agente"] = (~cond_antes_agente).astype(int)
                ddc_post_agente = df_ddc.groupby("ID Chat")["Es_Post_Agente"].sum().reset_index()
                ddc_post_agente.rename(columns={"Es_Post_Agente": "Mensajes_Post_Agente"}, inplace=True)
                
                # 4. Mensajes Post-Crédito por Chat (adicionales, no ya descontados por agente)
                df_ddc["Es_Post_Credito"] = (cond_antes_agente & (~cond_antes_credito)).astype(int)
                ddc_post_credito = df_ddc.groupby("ID Chat")["Es_Post_Credito"].sum().reset_index()
                ddc_post_credito.rename(columns={"Es_Post_Credito": "Mensajes_Post_Credito"}, inplace=True)
                
                # 5. Metadatos (Tiempos de corte)
                ddc_meta = df_ddc[["ID Chat", "Time_Agente", "Time_Credito"]].groupby("ID Chat").first().reset_index()
                
                # Unir todas las métricas DDC
                ddc_view = pd.merge(ddc_counts, ddc_bruto, on="ID Chat", how="left")
                ddc_view = pd.merge(ddc_view, ddc_post_agente, on="ID Chat", how="left")
                ddc_view = pd.merge(ddc_view, ddc_post_credito, on="ID Chat", how="left")
                ddc_view = pd.merge(ddc_view, ddc_meta, on="ID Chat", how="left")
                ddc_view["ID Chat"] = ddc_view["ID Chat"].astype(str)
                
                # Merge Master (Left Join al RDC porque es la base de conversaciones)
                df_detalle = pd.merge(rdc_view, ddc_view, on="ID Chat", how="left")
                
                # Limpieza final para el Excel
                df_detalle["Mensajes_Facturables"] = df_detalle["Mensajes_Facturables"].fillna(0).astype(int)
                df_detalle["Mensajes_Bruto"] = df_detalle["Mensajes_Bruto"].fillna(0).astype(int)
                df_detalle["Mensajes_Post_Agente"] = df_detalle["Mensajes_Post_Agente"].fillna(0).astype(int)
                df_detalle["Mensajes_Post_Credito"] = df_detalle["Mensajes_Post_Credito"].fillna(0).astype(int)
                
                # Convertir Es_Credito de booleano a numérico (1/0)
                df_detalle["Es_Credito"] = df_detalle["Es_Credito"].astype(int)
                
                # Extraer Fecha (Día) para análisis temporal
                df_detalle["Fecha_Dia"] = df_detalle["F.Inicio Chat"].dt.date
                
                # Seleccionar y Ordenar Columnas para el Excel (SIN HSM_Facturable_Individual)
                df_detalle = df_detalle[[
                    "ID Chat",
                    "Fecha_Dia",
                    "F.Inicio Chat",
                    "Tipificación Chat",
                    "Es_Cobrable",
                    "Es_Credito",
                    "Mensajes_Bruto",
                    "Mensajes_Post_Agente",
                    "Mensajes_Post_Credito",
                    "Mensajes_Facturables",
                    "Time_Agente",
                    "Time_Credito"
                ]]

            else:
                total_q_mensajes = 0
                mensajes_bruto = 0
                mensajes_agente = 0
                mensajes_credito = 0
                hsm_bruto = total_q_hsm 
                hsm_credito = 0
                total_q_hsm = max(0, total_q_hsm - 1000)
                
                # Detalle solo RDC (sin DDC)
                df_detalle = df_rdc[["ID Chat", "F.Inicio Chat", "Tipificación Chat", "Es_Cobrable"]].copy()
                df_detalle["Fecha_Dia"] = df_detalle["F.Inicio Chat"].dt.date
                df_detalle["Es_Credito"] = df_rdc["Es_Credito"].astype(int)
                df_detalle["Mensajes_Bruto"] = 0
                df_detalle["Mensajes_Post_Agente"] = 0
                df_detalle["Mensajes_Post_Credito"] = 0
                df_detalle["Mensajes_Facturables"] = 0
                df_detalle["Time_Agente"] = pd.NaT
                df_detalle["Time_Credito"] = pd.NaT

            # 3. RESULTADOS
            progress_bar.progress(100)
            status_container.success("✅ ¡Cálculo Completado!")
            
            # Tarjetas de KPI
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(label="HSM Bruto", value=f"{hsm_bruto:,.0f}", delta=f"- {hsm_credito} (Crédito)")
            with col2:
                st.metric(label="Q HSM (Final Facturable)", value=f"{total_q_hsm:,.0f}", delta="- 1,000 (Meta)")
            with col3:
                st.metric(label="Q Mensajes (Facturables)", value=f"{total_q_mensajes:,.0f}")
            
            # Descarga
            st.markdown("---")
            st.subheader("📥 Descargar Reporte")
            
            excel_data = get_excel_bytes(total_q_hsm, total_q_mensajes, hsm_bruto, hsm_credito, mensajes_bruto, mensajes_agente, mensajes_credito, df_detalle)
            
            st.download_button(
                label="📄 Descargar FACTURA_FINAL.xlsx",
                data=excel_data,
                file_name="FACTURA_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
        except Exception as e:
            status_container.error(f"❌ Error: {str(e)}")

# Información Footer
st.sidebar.markdown("---")
st.sidebar.info("v1.0 - Calculadora Web Local")
