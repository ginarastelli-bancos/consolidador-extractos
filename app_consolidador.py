import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Extractos Bancarios",
    page_icon="🏦",
    layout="wide"
)

# Título
st.title("🏦 Consolidador de Extractos Bancarios")
st.markdown("---")

# Nombres de columnas esperadas (ajusta según tu estructura)
COLUMNAS_REQUERIDAS = [
    'Fecha',
    'Descripción',
    'Importe',
    'Saldo',
    'Banco',
    'Cuenta'
]

def procesar_archivo(archivo, logs):
    """
    Procesa un archivo Excel y extrae las columnas requeridas.
    """
    try:
        # Leer el archivo Excel con openpyxl
        df = pd.read_excel(archivo, sheet_name=0, engine='openpyxl')
        
        # Buscar las columnas requeridas (sin importar el orden)
        columnas_encontradas = {}
        for col_requerida in COLUMNAS_REQUERIDAS:
            # Buscar la columna por nombre (case-insensitive)
            col_encontrada = None
            for col in df.columns:
                if str(col).strip().lower() == col_requerida.lower():
                    col_encontrada = col
                    break
            
            if col_encontrada:
                columnas_encontradas[col_requerida] = col_encontrada
            else:
                logs.append(f"{archivo.name}: ⚠️ Columna '{col_requerida}' no encontrada")
        
        # Si no se encontraron todas las columnas requeridas
        if len(columnas_encontradas) < len(COLUMNAS_REQUERIDAS):
            logs.append(f"{archivo.name}: ❌ Faltan columnas requeridas")
            return None
        
        # Extraer las columnas en el orden requerido
        df_extraido = pd.DataFrame()
        for col_requerida in COLUMNAS_REQUERIDAS:
            df_extraido[col_requerida] = df[columnas_encontradas[col_requerida]]
        
        logs.append(f"{archivo.name}: ✅ Procesado correctamente ({len(df_extraido)} filas)")
        return df_extraido
        
    except Exception as e:
        logs.append(f"{archivo.name}: ❌ Error al leer archivo: {str(e)}")
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
        logs.append(f"\n✅ Consolidación completada: {len(df_final)} filas totales")
        return df_final, logs
    else:
        logs.append("\n❌ No se pudieron extraer datos")
        return None, logs

# Interfaz de usuario
st.markdown("### 📁 Subir archivos Excel")
archivos_subidos = st.file_uploader(
    "Selecciona uno o más archivos Excel (.xlsx)",
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
                
                # Botón de descarga
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name='Consolidado')
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
        
        **Columnas requeridas en cada archivo:**
        - Fecha
        - Descripción
        - Importe
        - Saldo
        - Banco
        - Cuenta
        
        **Nota:** Las columnas pueden estar en cualquier orden en los archivos originales.
        """)

# Footer
st.markdown("---")
st.markdown("*Consolidador de Extractos Bancarios - Despegar*")
