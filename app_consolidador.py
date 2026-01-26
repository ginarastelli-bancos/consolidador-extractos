import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Extractos Bancarios",
    page_icon="🏦",
    layout="wide"
)

# Título
st.title("🏦 Consolidador de Extractos Bancarios - Despegar")
st.markdown("---")

# MAPEO DE COLUMNAS: nombre_estandar -> [posibles_variantes]
COLUMNAS_MAPPING = {
    'ENTIDAD_LEGAL': ['ENTIDAD_LEGAL', 'ENTIDAD LEGAL'],
    'NOMBRE_BANCO': ['NOMBRE_BANCO', 'NOMBRE BANCO', 'NOMBRE_BAN', 'NOMBRE BAN'],
    'CTA_BANCO': ['CTA_BANCO', 'CTA BANCO', 'CTA_BANC', 'CTA BANC'],
    'CTA_NUMERO': ['CTA_NUMERO', 'CTA NUMERO', 'CTA_NUMEI', 'CTA NUMEI'],
    'EXTRACTO_NUM': ['EXTRACTO_NUM', 'EXTRACTO NUM', 'EXTRACT'],
    'EXTRACTO_FECHA': ['EXTRACTO_FECHA', 'EXTRACTO FECHA', 'EXTRACT'],
    'EXT_LINEA_NUM': ['EXT_LINEA_NUM', 'EXT LINEA NUM', 'EXT_LINE', 'EXT LINE'],
    'EXT_TIPO_TRX': ['EXT_TIPO_TRX', 'EXT TIPO TRX', 'EXT_TIPO', 'EXT TIPO'],
    'TRX_CODE': ['TRX_CODE', 'TRX CODE'],
    'EXT_LIN_MONTO': ['EXT_LIN_MONTO', 'EXT LIN MONTO', 'EXT_LIN_MONT', 'EXT LIN MONT'],
    'EXT_LIN_ID': ['EXT_LIN_ID', 'EXT LIN ID'],
    'STATUS': ['STATUS'],
    'TRX_TEXT': ['TRX_TEXT', 'TRX TEXT'],
    'NRO_DOCUMENTO': ['NRO_DOCUMENTO', 'NRO DOCUMENTO', 'NRO_DO', 'NRO DO'],
    'SOCIO_COMERCIAL': ['SOCIO_COMERCIAL', 'SOCIO COMERCIAL'],
    'COMENTARIO_ESPERADO': ['COMENTARIO_ESPERADO', 'COMENTARIO ESPERADO']
}

# ORDEN FINAL DE COLUMNAS (con MES al inicio)
COLUMNAS_FINALES = [
    'MES',
    'ENTIDAD_LEGAL',
    'NOMBRE_BANCO',
    'CTA_BANCO',
    'CTA_NUMERO',
    'EXTRACTO_NUM',
    'EXTRACTO_FECHA',
    'EXT_LINEA_NUM',
    'EXT_TIPO_TRX',
    'TRX_CODE',
    'EXT_LIN_MONTO',
    'EXT_LIN_ID',
    'STATUS',
    'TRX_TEXT',
    'NRO_DOCUMENTO',
    'SOCIO_COMERCIAL',
    'COMENTARIO_ESPERADO'
]

def calcular_mes(fecha):
    """
    Calcula el mes en formato MM/YYYY a partir de una fecha.
    """
    try:
        if pd.isna(fecha):
            return None
        
        # Convertir a datetime si no lo es
        if not isinstance(fecha, pd.Timestamp):
            fecha = pd.to_datetime(fecha)
        
        # Formato: MM/YYYY (por ejemplo: 09/2025)
        return fecha.strftime('%m/%Y')
    except:
        return None

def normalizar_nombre(nombre):
    """
    Normaliza un nombre de columna eliminando espacios, guiones bajos y puntos.
    """
    return str(nombre).strip().upper().replace('_', '').replace(' ', '').replace('.', '')

def buscar_columna(df_columns, variantes):
    """
    Busca una columna entre sus posibles variantes.
    """
    for variante in variantes:
        variante_normalizada = normalizar_nombre(variante)
        for col in df_columns:
            col_normalizada = normalizar_nombre(col)
            if col_normalizada == variante_normalizada:
                return col
    return None

def procesar_pestaña(df, nombre_pestaña, nombre_archivo):
    """
    Procesa una pestaña de Excel y extrae las columnas requeridas.
    """
    logs = []
    
    try:
        # Buscar las columnas requeridas (incluyendo variantes)
        columnas_encontradas = {}
        columnas_faltantes = []
        
        for col_estandar, variantes in COLUMNAS_MAPPING.items():
            col_encontrada = buscar_columna(df.columns, variantes)
            
            if col_encontrada:
                columnas_encontradas[col_estandar] = col_encontrada
            else:
                columnas_faltantes.append(col_estandar)
        
        # Si faltan columnas, esta pestaña no es válida
        if columnas_faltantes:
            return None, logs
        
        # Extraer las columnas y renombrarlas al estándar
        df_extraido = pd.DataFrame()
        for col_estandar in COLUMNAS_FINALES[1:]:  # Excluir MES que se calculará
            df_extraido[col_estandar] = df[columnas_encontradas[col_estandar]]
        
        # Calcular la columna MES a partir de EXTRACTO_FECHA
        df_extraido['EXTRACTO_FECHA'] = pd.to_datetime(df_extraido['EXTRACTO_FECHA'], errors='coerce')
        df_extraido['MES'] = df_extraido['EXTRACTO_FECHA'].apply(calcular_mes)
        
        # Reordenar columnas con MES al inicio
        df_extraido = df_extraido[COLUMNAS_FINALES]
        
        logs.append(f"  ✅ Pestaña '{nombre_pestaña}': {len(df_extraido)} filas procesadas")
        return df_extraido, logs
        
    except Exception as e:
        logs.append(f"  ❌ Pestaña '{nombre_pestaña}': Error - {str(e)}")
        return None, logs

def procesar_archivo(archivo, logs):
    """
    Procesa un archivo Excel buscando en TODAS sus pestañas.
    """
    try:
        # Reiniciar el puntero del archivo
        archivo.seek(0)
        
        # Leer el archivo Excel
        try:
            xls = pd.ExcelFile(archivo, engine='openpyxl')
        except Exception as e:
            logs.append(f"{archivo.name}: ❌ No se pudo abrir el archivo: {str(e)}")
            return None
        
        # Listar todas las pestañas
        pestañas = xls.sheet_names
        logs.append(f"{archivo.name}: 📋 Pestañas encontradas: {', '.join(pestañas)}")
        
        # Procesar cada pestaña
        datos_validos = []
        
        for nombre_pestaña in pestañas:
            try:
                df = pd.read_excel(xls, sheet_name=nombre_pestaña)
                
                # Verificar que no esté vacío
                if df.empty:
                    logs.append(f"  ⚠️ Pestaña '{nombre_pestaña}': vacía, omitiendo")
                    continue
                
                # Intentar procesar esta pestaña
                df_procesado, logs_pestaña = procesar_pestaña(df, nombre_pestaña, archivo.name)
                logs.extend(logs_pestaña)
                
                if df_procesado is not None:
                    datos_validos.append(df_procesado)
                    
            except Exception as e:
                logs.append(f"  ⚠️ Pestaña '{nombre_pestaña}': Error al leer - {str(e)}")
                continue
        
        # Consolidar datos de todas las pestañas válidas
        if datos_validos:
            df_final = pd.concat(datos_validos, ignore_index=True)
            logs.append(f"{archivo.name}: ✅ Total procesado: {len(df_final)} filas de {len(datos_validos)} pestaña(s)")
            return df_final
        else:
            logs.append(f"{archivo.name}: ❌ No se encontraron pestañas con las columnas requeridas")
            logs.append(f"  Columnas buscadas:")
            for col in COLUMNAS_FINALES[1:]:
                logs.append(f"    • {col}")
            return None
        
    except Exception as e:
        logs.append(f"{archivo.name}: ❌ Error inesperado: {str(e)}")
        import traceback
        logs.append(f"  Detalle: {traceback.format_exc()}")
        return None

def consolidar_archivos(archivos):
    """
    Consolida múltiples archivos Excel en uno solo.
    """
    logs = []
    datos_consolidados = []
    
    for archivo in archivos:
        df = procesar_archivo(archivo, logs)
        if df is not None:
            datos_consolidados.append(df)
    
    if datos_consolidados:
        df_final = pd.concat(datos_consolidados, ignore_index=True)
        
        # Ordenar por fecha (más reciente primero)
        try:
            df_final = df_final.sort_values('EXTRACTO_FECHA', ascending=False)
        except:
            logs.append("⚠️ No se pudo ordenar por fecha")
        
        logs.append(f"\n✅ Consolidación completada: {len(df_final)} filas totales de {len(datos_consolidados)} archivo(s)")
        return df_final, logs
    else:
        logs.append("\n❌ No se pudieron extraer datos de ningún archivo")
        return None, logs

# Mostrar información de dependencias instaladas
with st.expander("🔧 Información del sistema (debug)"):
    import sys
    st.code(f"Python: {sys.version}")
    st.code(f"Pandas: {pd.__version__}")
    
    # Verificar qué engines están disponibles
    engines_disponibles = []
    try:
        import openpyxl
        engines_disponibles.append(f"openpyxl: {openpyxl.__version__}")
    except:
        engines_disponibles.append("openpyxl: NO DISPONIBLE")
    
    try:
        import xlrd
        engines_disponibles.append(f"xlrd: {xlrd.__version__}")
    except:
        engines_disponibles.append("xlrd: NO DISPONIBLE")
    
    try:
        import xlsxwriter
        engines_disponibles.append(f"xlsxwriter: {xlsxwriter.__version__}")
    except:
        engines_disponibles.append("xlsxwriter: NO DISPONIBLE")
    
    st.code("\n".join(engines_disponibles))
    
    st.markdown("**Columnas esperadas (con variantes aceptadas):**")
    for col_estandar, variantes in COLUMNAS_MAPPING.items():
        st.text(f"• {col_estandar}: {', '.join(variantes)}")
    st.info("⚠️ El sistema busca en TODAS las pestañas de cada archivo Excel")

# Interfaz de usuario
st.markdown("### 📁 Subir archivos Excel")
archivos_subidos = st.file_uploader(
    "Selecciona uno o más archivos Excel (.xlsx) con extractos bancarios",
    type=['xlsx'],
    accept_multiple_files=True
)

if archivos_subidos:
    st.info(f"📊 {len(archivos_subidos)} archivo(s) cargado(s)")
    
    if st.button("🔄 Consolidar archivos", type="primary"):
        with st.spinner("Procesando..."):
            df_consolidado, logs = consolidar_archivos(archivos_subidos)
            
            # Mostrar log de procesamiento
            st.markdown("### 📋 Log de procesamiento")
            for log in logs:
                if "✅" in log:
                    st.success(log)
                elif "⚠️" in log:
                    st.warning(log)
                elif "❌" in log:
                    st.error(log)
                else:
                    st.info(log)
            
            # Si hay datos consolidados, mostrar y permitir descarga
            if df_consolidado is not None:
                st.markdown("---")
                st.markdown("### 📊 Vista previa del consolidado")
                st.dataframe(df_consolidado.head(100), use_container_width=True)
                
                st.markdown(f"**Total de registros:** {len(df_consolidado)}")
                
                # Mostrar resumen por mes
                if 'MES' in df_consolidado.columns:
                    st.markdown("### 📅 Resumen por mes")
                    resumen_mes = df_consolidado['MES'].value_counts().sort_index(ascending=False)
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.dataframe(resumen_mes.reset_index().rename(columns={'MES': 'Mes', 'count': 'Cantidad'}), 
                                   use_container_width=True, hide_index=True)
                
                # Botón de descarga
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name='Consolidado')
                    
                    # Formatear el Excel
                    workbook = writer.book
                    worksheet = writer.sheets['Consolidado']
                    
                    # Formato de encabezado
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#0066CC',
                        'font_color': 'white',
                        'border': 1
                    })
                    
                    # Aplicar formato a encabezados
                    for col_num, value in enumerate(df_consolidado.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Ajustar ancho de columnas
                    worksheet.set_column('A:Q', 15)
                    worksheet.set_column('N:N', 50)  # TRX_TEXT más ancho
                
                output.seek(0)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"consolidado_extractos_{timestamp}.xlsx"
                
                st.download_button(
                    label="📥 Descargar archivo consolidado",
                    data=output,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.error("❌ Error al consolidar los archivos. Revisa el log de procesamiento.")

else:
    st.info("👆 Sube uno o más archivos Excel para comenzar")
    
    # Información de uso
    with st.expander("ℹ️ Instrucciones de uso"):
        st.markdown("""
        **Cómo usar esta aplicación:**
        
        1. Haz clic en "Browse files" para seleccionar tus archivos Excel
        2. Puedes seleccionar múltiples archivos a la vez
        3. Haz clic en "Consolidar archivos" para procesarlos
        4. Revisa el log de procesamiento para verificar que todo está correcto
        5. Descarga el archivo consolidado
        
        **Características:**
        
        - ✅ **Busca en TODAS las pestañas** de cada archivo Excel automáticamente
        - ✅ Acepta variantes de nombres de columnas (con espacios, guiones bajos, abreviadas)
        - ✅ Consolida datos de múltiples archivos y pestañas
        - ✅ Ordena por fecha (más reciente primero)
        - ✅ Calcula la columna MES automáticamente (formato MM/YYYY)
        
        **Columnas del archivo consolidado:**
        
        1. **MES** (calculada automáticamente)
        2. ENTIDAD_LEGAL
        3. NOMBRE_BANCO
        4. CTA_BANCO
        5. CTA_NUMERO
        6. EXTRACTO_NUM
        7. EXTRACTO_FECHA
        8. EXT_LINEA_NUM
        9. EXT_TIPO_TRX
        10. TRX_CODE
        11. EXT_LIN_MONTO
        12. EXT_LIN_ID
        13. STATUS
        14. TRX_TEXT
        15. NRO_DOCUMENTO
        16. SOCIO_COMERCIAL
        17. COMENTARIO_ESPERADO
        """)

# Footer
st.markdown("---")
st.markdown("*Consolidador de Extractos Bancarios - Despegar © 2026*")
