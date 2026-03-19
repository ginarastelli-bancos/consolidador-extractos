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

# PESTAÑAS A EXCLUIR
PESTAÑAS_EXCLUIDAS = ['GLOSARIO', 'CONTROL SALDOS', 'CONTROL DE SALDOS', 'CONTROLSALDOS', 'LLAVE']

# COLUMNAS OBLIGATORIAS
COLUMNAS_OBLIGATORIAS = {
    'ENTIDAD_LEGAL': ['ENTIDAD_LEGAL', 'ENTIDAD LEGAL'],
    'NOMBRE_BANCO': ['NOMBRE_BANCO', 'NOMBRE BANCO', 'NOMBRE_BAN', 'NOMBRE BAN'],
    'CTA_BANCO': ['CTA_BANCO', 'CTA BANCO', 'CTA_BANC', 'CTA BANC'],
    'CTA_NUMERO': ['CTA_NUMERO', 'CTA NUMERO', 'CTA_NUMEI', 'CTA NUMEI'],
    'EXTRACTO_FECHA': ['EXTRACTO_FECHA', 'EXTRACTO FECHA', 'EXTRACT'],
    'EXT_TIPO_TRX': ['EXT_TIPO_TRX', 'EXT TIPO TRX', 'EXT_TIPO', 'EXT TIPO'],
    'EXT_LIN_MONTO': ['EXT_LIN_MONTO', 'EXT LIN MONTO', 'EXT_LIN_MONT', 'EXT LIN MONT'],
    'STATUS': ['STATUS'],
    'TRX_TEXT': ['TRX_TEXT', 'TRX TEXT'],
    'COMENTARIO_ESPERADO': ['COMENTARIO_ESPERADO', 'COMENTARIO ESPERADO']
}

# COLUMNAS OPCIONALES
COLUMNAS_OPCIONALES = {
    'EXTRACTO_NUM': ['EXTRACTO_NUM', 'EXTRACTO NUM', 'EXTRACT'],
    'EXT_LINEA_NUM': ['EXT_LINEA_NUM', 'EXT LINEA NUM', 'EXT_LINE', 'EXT LINE'],
    'TRX_CODE': ['TRX_CODE', 'TRX CODE'],
    'EXT_LIN_ID': ['EXT_LIN_ID', 'EXT LIN ID'],
    'NRO_DOCUMENTO': ['NRO_DOCUMENTO', 'NRO DOCUMENTO', 'NRO_DO', 'NRO DO'],
    'SOCIO_COMERCIAL': ['SOCIO_COMERCIAL', 'SOCIO COMERCIAL'],
    'INFORMACION_ADICIONAL': ['INFORMACION_ADICIONAL', 'INFORMACION ADICIONAL', 'INFO_ADICIONAL', 'INFO ADICIONAL', 'INFORMACION_ADIC', 'INFORMACION ADIC']
}

# COLUMNAS QUE DEBEN SER TEXTO (preservar formato original)
COLUMNAS_FORMATO_TEXTO = [
    'ENTIDAD_LEGAL',
    'NOMBRE_BANCO',
    'CTA_BANCO',
    'CTA_NUMERO',
    'EXTRACTO_NUM',
    'EXT_LINEA_NUM',
    'EXT_TIPO_TRX',
    'TRX_CODE',
    'EXT_LIN_ID',
    'STATUS',
    'TRX_TEXT',
    'NRO_DOCUMENTO',
    'SOCIO_COMERCIAL',
    'COMENTARIO_ESPERADO',
    'INFORMACION_ADICIONAL'
]

# ORDEN FINAL DE COLUMNAS
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
    'COMENTARIO_ESPERADO',
    'INFORMACION_ADICIONAL'
]

def calcular_mes(fecha):
    """Calcula el mes en formato MM/YYYY"""
    try:
        if pd.isna(fecha):
            return None
        if not isinstance(fecha, pd.Timestamp):
            fecha = pd.to_datetime(fecha)
        return fecha.strftime('%m/%Y')
    except:
        return None

def normalizar_nombre(nombre):
    """Normaliza un nombre de columna"""
    return str(nombre).strip().upper().replace('_', '').replace(' ', '').replace('.', '')

def buscar_columna(df_columns, variantes):
    """Busca una columna entre sus variantes"""
    for variante in variantes:
        variante_normalizada = normalizar_nombre(variante)
        for col in df_columns:
            col_normalizada = normalizar_nombre(col)
            if col_normalizada == variante_normalizada:
                return col
    return None

def es_pestaña_excluida(nombre_pestaña):
    """Verifica si una pestaña debe ser excluida"""
    nombre_normalizado = nombre_pestaña.strip().upper()
    for excluida in PESTAÑAS_EXCLUIDAS:
        if excluida in nombre_normalizado:
            return True
    return False

def crear_converters(columnas_originales, columnas_encontradas):
    """Crea converters para leer columnas como texto"""
    converters = {}
    for col_estandar, col_original in columnas_encontradas.items():
        if col_estandar in COLUMNAS_FORMATO_TEXTO:
            converters[col_original] = lambda x: str(x) if pd.notna(x) and x != '' else ''
    return converters

def procesar_pestaña(df, nombre_pestaña):
    """Procesa una pestaña de Excel"""
    logs = []
    
    try:
        if df.empty:
            return None, logs
        
        # Buscar columnas obligatorias
        columnas_encontradas = {}
        columnas_obligatorias_faltantes = []
        
        for col_estandar, variantes in COLUMNAS_OBLIGATORIAS.items():
            col_encontrada = buscar_columna(df.columns, variantes)
            if col_encontrada:
                columnas_encontradas[col_estandar] = col_encontrada
            else:
                columnas_obligatorias_faltantes.append(col_estandar)
        
        # Si faltan columnas obligatorias, omitir
        if columnas_obligatorias_faltantes:
            return None, logs
        
        # Buscar columnas opcionales
        columnas_opcionales_faltantes = []
        for col_estandar, variantes in COLUMNAS_OPCIONALES.items():
            col_encontrada = buscar_columna(df.columns, variantes)
            if col_encontrada:
                columnas_encontradas[col_estandar] = col_encontrada
            else:
                columnas_opcionales_faltantes.append(col_estandar)
        
        # Extraer columnas preservando el formato
        df_extraido = pd.DataFrame()
        for col_estandar in COLUMNAS_FINALES[1:]:
            if col_estandar in columnas_encontradas:
                col_original = columnas_encontradas[col_estandar]
                if col_estandar in COLUMNAS_FORMATO_TEXTO:
                    df_extraido[col_estandar] = df[col_original].apply(
                        lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != '' and str(x).lower() != 'nan' else ''
                    )
                # ── CAMBIO 1: EXT_LIN_MONTO siempre como número ──
                elif col_estandar == 'EXT_LIN_MONTO':
                    df_extraido[col_estandar] = pd.to_numeric(
                        df[col_original]
                            .astype(str)
                            .str.strip()
                            .str.replace(r'[^\d.,-]', '', regex=True)
                            .str.replace(',', '.', regex=False),
                        errors='coerce'
                    )
                # ─────────────────────────────────────────────────
                else:
                    df_extraido[col_estandar] = df[col_original]
            else:
                df_extraido[col_estandar] = ''
        
        # Calcular MES
        df_extraido['EXTRACTO_FECHA'] = pd.to_datetime(df_extraido['EXTRACTO_FECHA'], errors='coerce')
        df_extraido['MES'] = df_extraido['EXTRACTO_FECHA'].apply(calcular_mes)
        df_extraido = df_extraido[COLUMNAS_FINALES]
        
        logs.append(f"  ✅ {nombre_pestaña}: {len(df_extraido):,} registros")
        return df_extraido, logs
        
    except Exception as e:
        logs.append(f"  ❌ {nombre_pestaña}: Error - {str(e)}")
        return None, logs

def procesar_archivo(archivo, logs):
    """Procesa un archivo Excel"""
    try:
        archivo.seek(0)
        
        try:
            xls = pd.ExcelFile(archivo, engine='openpyxl')
        except Exception as e:
            logs.append(f"❌ {archivo.name}: No se pudo abrir")
            return None
        
        pestañas = xls.sheet_names
        logs.append(f"\n📁 {archivo.name}")
        
        # Filtrar pestañas excluidas
        pestañas_excluidas = [p for p in pestañas if es_pestaña_excluida(p)]
        pestañas_a_procesar = [p for p in pestañas if not es_pestaña_excluida(p)]
        
        if pestañas_excluidas:
            logs.append(f"  🚫 Omitidas: {', '.join(pestañas_excluidas)}")
        
        # Procesar pestañas
        datos_validos = []
        pestañas_procesadas = []
        
        for nombre_pestaña in pestañas_a_procesar:
            try:
                df = pd.read_excel(xls, sheet_name=nombre_pestaña, dtype=str)
                df_procesado, logs_pestaña = procesar_pestaña(df, nombre_pestaña)
                logs.extend(logs_pestaña)
                
                if df_procesado is not None and len(df_procesado) > 0:
                    datos_validos.append(df_procesado)
                    pestañas_procesadas.append(nombre_pestaña)
                    
            except Exception as e:
                continue
        
        # Consolidar
        if datos_validos:
            df_final = pd.concat(datos_validos, ignore_index=True)
            
            if len(pestañas_procesadas) > 1:
                logs.append(f"  📊 Consolidadas: {', '.join(pestañas_procesadas)}")
            
            logs.append(f"  ✅ TOTAL: {len(df_final):,} registros de {len(pestañas_procesadas)} pestaña(s)")
            return df_final
        else:
            logs.append(f"  ⚠️ No se encontraron datos válidos")
            return None
        
    except Exception as e:
        logs.append(f"❌ {archivo.name}: Error inesperado")
        return None

def consolidar_archivos(archivos):
    """Consolida múltiples archivos Excel"""
    logs = []
    datos_consolidados = []
    
    logs.append("🔄 PROCESANDO ARCHIVOS...")
    
    for archivo in archivos:
        df = procesar_archivo(archivo, logs)
        if df is not None:
            datos_consolidados.append(df)
    
    if datos_consolidados:
        df_final = pd.concat(datos_consolidados, ignore_index=True)
        
        # Ordenar por fecha
        try:
            df_final['EXTRACTO_FECHA'] = pd.to_datetime(df_final['EXTRACTO_FECHA'], errors='coerce')
            df_final = df_final.sort_values('EXTRACTO_FECHA', ascending=False)
        except:
            pass
        
        logs.append(f"\n{'='*60}")
        logs.append(f"✅ CONSOLIDACIÓN EXITOSA")
        logs.append(f"Total de registros: {len(df_final):,}")
        logs.append(f"Archivos procesados: {len(datos_consolidados)}")
        logs.append(f"{'='*60}")
        
        return df_final, logs
    else:
        logs.append(f"\n❌ No se pudieron extraer datos")
        return None, logs

# Interfaz de usuario
st.markdown("### 📁 Subir archivos Excel")

archivos_subidos = st.file_uploader(
    "Arrastra o selecciona los archivos de extractos bancarios (.xlsx)",
    type=['xlsx'],
    accept_multiple_files=True
)

if archivos_subidos:
    st.success(f"✅ {len(archivos_subidos)} archivo(s) cargado(s)")
    
    if st.button("🔄 Consolidar archivos", type="primary", use_container_width=True):
        with st.spinner("⏳ Procesando..."):
            df_consolidado, logs = consolidar_archivos(archivos_subidos)
            
            # Mostrar log
            st.markdown("### 📋 Resultado del procesamiento")
            for log in logs:
                if "✅" in log and "EXITOSA" in log:
                    st.success(log)
                elif "✅" in log:
                    st.info(log)
                elif "⚠️" in log:
                    st.warning(log)
                elif "❌" in log:
                    st.error(log)
                elif log.strip().startswith("="):
                    st.markdown(f"**{log}**")
                else:
                    st.text(log)
            
            # Mostrar datos consolidados
            if df_consolidado is not None:
                st.markdown("---")
                st.markdown("### 📊 Vista previa del consolidado")
                st.dataframe(df_consolidado.head(50), use_container_width=True)
                
                # Métricas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total registros", f"{len(df_consolidado):,}")
                with col2:
                    if 'ENTIDAD_LEGAL' in df_consolidado.columns:
                        st.metric("Entidades", df_consolidado['ENTIDAD_LEGAL'].nunique())
                with col3:
                    if 'NOMBRE_BANCO' in df_consolidado.columns:
                        st.metric("Bancos", df_consolidado['NOMBRE_BANCO'].nunique())
                with col4:
                    if 'MES' in df_consolidado.columns:
                        st.metric("Meses", df_consolidado['MES'].nunique())
                
                # Resumen por mes
                if 'MES' in df_consolidado.columns:
                    st.markdown("### 📅 Distribución por mes")
                    resumen_mes = df_consolidado['MES'].value_counts().sort_index(ascending=False)
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.dataframe(
                            resumen_mes.reset_index().rename(columns={'MES': 'Mes', 'count': 'Registros'}), 
                            use_container_width=True, 
                            hide_index=True
                        )
                
                # Descarga
                st.markdown("---")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name='Consolidado')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Consolidado']
                    
                    # Formato para encabezados
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#0066CC',
                        'font_color': 'white',
                        'border': 1,
                        'align': 'center',
                        'valign': 'vcenter'
                    })
                    
                    # Formato de texto
                    text_format = workbook.add_format({
                        'num_format': '@',
                        'align': 'left'
                    })
                    
                    # Formato de fecha
                    date_format = workbook.add_format({
                        'num_format': 'dd/mm/yyyy',
                        'align': 'center'
                    })
                    
                    # Formato de número/monto
                    number_format = workbook.add_format({
                        'num_format': '#,##0.00',
                        'align': 'right'
                    })
                    
                    # Escribir encabezados
                    for col_num, value in enumerate(df_consolidado.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Aplicar formatos a columnas específicas
                    num_rows = len(df_consolidado)
                    
                    for col_num, col_name in enumerate(df_consolidado.columns):
                        if col_name in COLUMNAS_FORMATO_TEXTO:
                            worksheet.set_column(col_num, col_num, None, text_format)
                            for row_num in range(1, num_rows + 1):
                                value = df_consolidado.iloc[row_num - 1, col_num]
                                worksheet.write_string(row_num, col_num, str(value) if pd.notna(value) else '', text_format)
                        elif col_name == 'EXTRACTO_FECHA':
                            worksheet.set_column(col_num, col_num, 12, date_format)
                        # ── CAMBIO 2: escribir EXT_LIN_MONTO como número real ──
                        elif col_name == 'EXT_LIN_MONTO':
                            worksheet.set_column(col_num, col_num, 15, number_format)
                            for row_num in range(1, num_rows + 1):
                                value = df_consolidado.iloc[row_num - 1, col_num]
                                if pd.notna(value):
                                    try:
                                        worksheet.write_number(row_num, col_num, float(value), number_format)
                                    except (ValueError, TypeError):
                                        worksheet.write_blank(row_num, col_num, number_format)
                                else:
                                    worksheet.write_blank(row_num, col_num, number_format)
                        # ───────────────────────────────────────────────────────
                    
                    # Ajustar anchos de columna
                    worksheet.set_column('A:A', 12)
                    worksheet.set_column('B:B', 20)
                    worksheet.set_column('C:C', 20)
                    worksheet.set_column('D:D', 15)
                    worksheet.set_column('E:E', 20)
                    worksheet.set_column('F:F', 15)
                    worksheet.set_column('G:G', 15)
                    worksheet.set_column('H:H', 15)
                    worksheet.set_column('I:I', 12)
                    worksheet.set_column('J:J', 12)
                    worksheet.set_column('K:K', 15)
                    worksheet.set_column('L:L', 15)
                    worksheet.set_column('M:M', 12)
                    worksheet.set_column('N:N', 50)
                    worksheet.set_column('O:O', 20)
                    worksheet.set_column('P:P', 25)
                    worksheet.set_column('Q:Q', 25)
                    worksheet.set_column('R:R', 40)
                    
                    worksheet.freeze_panes(1, 0)
                
                output.seek(0)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"consolidado_extractos_{timestamp}.xlsx"
                
                st.download_button(
                    label="📥 Descargar archivo consolidado",
                    data=output,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

else:
    st.info("👆 Sube uno o más archivos Excel para comenzar")
    
    with st.expander("ℹ️ Información"):
        st.markdown("""
        **Características:**
        
        - ✅ Consolida múltiples archivos y pestañas
        - ✅ Extrae datos de todas las pestañas válidas
        - 🚫 Excluye automáticamente: Glosario, Control Saldos, Llave
        - ✅ Genera columna MES automáticamente
        - ✅ Ordena por fecha (más reciente primero)
        - ✅ Formato Excel profesional
        - ✅ Incluye columna INFORMACION_ADICIONAL
        - ⭐ **PRESERVA FORMATO DE ORIGEN** (CTA_BANCO, CTA_NUMERO, etc. como texto)
        
        **Columnas obligatorias:** ENTIDAD_LEGAL, NOMBRE_BANCO, CTA_BANCO, CTA_NUMERO, 
        EXTRACTO_FECHA, EXT_TIPO_TRX, EXT_LIN_MONTO, STATUS, TRX_TEXT, COMENTARIO_ESPERADO
        
        **Columnas opcionales:** EXTRACTO_NUM, EXT_LINEA_NUM, TRX_CODE, EXT_LIN_ID, 
        NRO_DOCUMENTO, SOCIO_COMERCIAL, INFORMACION_ADICIONAL
        
        **⚠️ Importante:** Las columnas CTA_BANCO, CTA_NUMERO y otras identificadoras 
        se exportan como TEXTO para preservar ceros iniciales y formato original.
        """)

st.markdown("---")
st.markdown("*Consolidador de Extractos Bancarios - Despegar © 2026*")
