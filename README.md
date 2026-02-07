# Calculadora de Facturación Quina

Aplicación web desarrollada con Streamlit para automatizar el cálculo de facturación mensual de servicios de WhatsApp Business API.

## 🚀 Características

- **Cálculo Automático de HSM (Conversaciones 24h)**
  - Detección de conversaciones únicas en ventanas de 24 horas
  - Descuento automático de conversaciones de crédito (Tipificación "evalú")
  - Descuento de 1,000 conversaciones gratuitas de Meta
  
- **Cálculo de Mensajes Facturables**
  - Corte automático de mensajes post-agente
  - Corte automático de mensajes post-crédito
  - Tarifas escalonadas por volumen
  
- **Factura Detallada**
  - Desglose completo de HSM (Bruto, Descuentos, Neto)
  - Desglose completo de Mensajes (Bruto, Descuentos, Neto)
  - Cálculo automático de IGV y total
  
- **Hoja de Auditoría**
  - Detalle por chat con todas las métricas
  - Columna de fecha para análisis temporal
  - Tipificación de cada conversación
  - Timestamps de corte (agente y crédito)

## 📋 Requisitos

- Python 3.8+
- Streamlit
- Pandas
- NumPy
- OpenPyXL

## 🔧 Instalación

```bash
# Clonar el repositorio
git clone https://github.com/fmendoza-a365/CalculadoraFacturaQuina.git
cd CalculadoraFacturaQuina

# Instalar dependencias
pip install -r requirements.txt
```

## 💻 Uso

```bash
# Ejecutar la aplicación
streamlit run QuinaWebCalculator.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

## 📁 Archivos de Entrada

La aplicación requiere dos archivos Excel mensuales:

1. **RDC (Reporte de Conversaciones)**
   - Columnas requeridas: `ID`, `F.Inicio Chat`, `ID Chat`, `Tipificación Chat`
   
2. **DDC (Detalle de Conversaciones)** *(Opcional)*
   - Columnas requeridas: `ID Chat`, `Mensaje`, `Fecha Hora`, `Tipo`

## 📊 Archivo de Salida

La aplicación genera un archivo Excel `FACTURA_FINAL.xlsx` con dos hojas:

### Hoja 1: Factura
- Fee Mensual
- Cálculo HSM Detallado (Bruto, Descuentos, Neto)
- Cálculo Mensajes Detallado (Bruto, Descuentos, Neto)
- Subtotal, IGV y Total

### Hoja 2: Detalle Auditoría
- Análisis por chat individual
- Métricas de HSM y Mensajes
- Timestamps de eventos clave
- Tipificación de conversaciones

## 🚂 Despliegue en Railway

Para desplegar la aplicación en Railway:

1. **Conecta tu repositorio GitHub** en Railway
2. **Railway detectará automáticamente** la configuración de Python
3. **El Procfile** iniciará la aplicación automáticamente
4. **Accede a tu app** en la URL proporcionada por Railway

La aplicación está configurada para:
- Usar el puerto dinámico de Railway (`$PORT`)
- Ejecutarse en modo headless (sin interfaz local)
- Aceptar conexiones externas

**Archivos de configuración incluidos:**
- `Procfile` - Comando de inicio
- `runtime.txt` - Versión de Python
- `setup.sh` - Configuración de Streamlit

## 🔍 Lógica de Negocio

### HSM (Conversaciones)
- Se cobra 1 HSM por cada conversación única en ventana de 24h
- Se descuentan conversaciones con tipificación que contenga "evalú"
- Se descuentan 1,000 conversaciones gratuitas de Meta

### Mensajes
- Se corta el conteo cuando el cliente es transferido a agente humano
- Se corta el conteo cuando el cliente activa la opción de crédito
- Tarifas escalonadas según volumen mensual

## 📝 Tarifas

### HSM
- S/ 0.077 por conversación

### Mensajes (Escalonadas)
- 1 - 9,999: S/ 0.0456
- 10,000 - 99,999: S/ 0.0380
- 100,000 - 249,999: S/ 0.0304
- 250,000+: S/ 0.0228

## 👨‍💻 Autor

Desarrollado para Quina - Automatización de Facturación WhatsApp Business API

## 📄 Licencia

Este proyecto es de uso interno.
