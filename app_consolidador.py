
"""
CONSOLIDADOR DE EXTRACTOS BANCARIOS - WEB APP
Aplicación web para consolidar extractos bancarios
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
import zipfile

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Extractos - Despegar",
    page_icon="📊",
    layout="wide"
)

class ConsolidadorExtractosWeb:
    
    def __init__(self):
        self.mapeo_columnas = {
            0: 'ENTIDAD_LEGAL', 1: 'NOMBRE_BANCO', 2: 'CTA_BANCO',
            3: 'CTA_NUMERO', 4: 'EXTRACTO_NUM', 5: 'EXTRACTO_FECHA',
            6: 'EXT_LINEA_NUM', 7: 'EXT_TIPO_TRX', 8: 'TRX_CODE',
            9: 'EXT_LIN_MONTO', 10: 'EXT_LIN_ID', 11: 'STATUS',
            12: 'TRX_TEXT', 13: 'NRO_DOCUMENTO', 14: 'SOCIO_COMERCIAL',
            15: 'COMENTARIO_ESPERADO'
        }
        
        self.columnas_finales = [
            'MES', 'ENTIDAD_LEGAL', 'NOMBRE_BANCO', 'CTA_BANCO', 'CTA_NUMERO',
            'EXTRACTO_NUM', 'EXTRACTO_FECHA', 'EXT_LINEA_NUM', 'EXT_TIPO_TRX',
            'TRX_CODE', 'EXT_LIN_MONTO', 'EXT_LIN_ID', 'STATUS', 'TRX_TEXT',
            'NRO_DOCUMENTO', 'SOCIO_COMERCIAL', 'COMENTARIO_ESPERADO'
        ]
        
        self.df_consolidado = pd.DataFrame()
        self.log = []
    
    def calcular_mes(self, fecha):
        if pd.isna(fecha) or fecha == '' or fecha is None:
            return ''
        try:
            if isinstance(fecha, str) and '/' in fecha:
                partes = fecha.split('/')
                if len(partes) == 3:
                    dia, mes, año = partes
                    return f"{mes}/{año}"
            if isinstance(fecha, (pd.Timestamp, datetime)):
                return fecha.strftime('%m/%Y')
            dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
            return dt.strftime('%m/%Y') if pd.notna(dt) else ''
        except:
            return ''
    
    def normalizar_fecha(self, fecha):
        if pd.isna(fecha) or fecha == '' or fecha is None:
            return None
        if isinstance(fecha, str) and '/' in fecha:
            partes = fecha.split('/')
            if len(partes) == 3 and len(partes[2]) == 4:
                return fecha
        try:
            if isinstance(fecha, (pd.Timestamp, datetime)):
                return fecha.strftime('%d/%m/%Y')
            dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
            return dt.strftime('%d/%m/%Y') if pd.notna(dt) else None
        except:
            return None
    
    def aplicar_signo_monto(self, monto, tipo_trx):
        try:
            monto_float = float(str(monto).replace(',', '').replace('$', '').strip())
        except:
            return 0.0
        tipo_str = str(tipo_trx).upper().strip()
        if tipo_str in ['DEBIT', 'MISC_DEBIT']:
            return -abs(monto_float)
        elif tipo_str in ['CREDIT', 'MISC_CREDIT']:
            return abs(monto_float)
        return monto_float
    
    def formato_texto_cuenta(self, valor):
        """Convierte número de cuenta a texto preservando ceros"""
        if pd.isna(valor) or valor == '':
            return ''
        # Convertir a string y eliminar decimales si existen
        valor_str = str(valor)
        if '.' in valor_str:
            valor_str = valor_str.split('.')[0]
        return valor_str
    
    def extraer_columnas_por_nombre(self, df):
        df_extraido = pd.DataFrame()
        for nombre_col in list(self.mapeo_columnas.values()):
            if nombre_col in df.columns:
                df_extraido[nombre_col] = df[nombre_col]
            else:
                df_extraido[nombre_col] = ''
        return df_extraido
    
    def procesar_archivo(self, archivo_bytes, nombre_archivo):
        """Procesa un archivo Excel desde bytes"""
        self.log.append(f"📄 Procesando: {nombre_archivo}")
        
        dfs_archivo = []
        
        try:
            xls = pd.ExcelFile(archivo_bytes)
            hojas = xls.sheet_names
            
            for nombre_hoja in hojas:
                if nombre_hoja.startswith('_'):
                    self.log.append(f"   ⊘ {nombre_hoja}: Saltada (auxiliar)")
                    continue
                
                try:
                    df_hoja = pd.read_excel(archivo_bytes, sheet_name=nombre_hoja)
                    
                    if df_hoja.empty:
                        self.log.append(f"   ⊘ {nombre_hoja}: Vacía")
                        continue
                    
                    df_procesado = self.extraer_columnas_por_nombre(df_hoja)
                    df_procesado = df_procesado.dropna(how='all')
                    
                    if len(df_procesado) > 0:
                        # Normalizar fechas
                        df_procesado['EXTRACTO_FECHA'] = df_procesado['EXTRACTO_FECHA'].apply(self.normalizar_fecha)
                        
                        # Calcular MES
                        df_procesado['MES'] = df_procesado['EXTRACTO_FECHA'].apply(self.calcular_mes)
                        
                        # Aplicar signos a montos
                        df_procesado['EXT_LIN_MONTO'] = df_procesado.apply(
                            lambda row: self.aplicar_signo_monto(row['EXT_LIN_MONTO'], row['EXT_TIPO_TRX']),
                            axis=1
                        )
                        
                        # Convertir cuentas a formato texto
                        df_procesado['CTA_BANCO'] = df_procesado['CTA_BANCO'].apply(self.formato_texto_cuenta)
                        df_procesado['CTA_NUMERO'] = df_procesado['CTA_NUMERO'].apply(self.formato_texto_cuenta)
                        
                        # Limpiar otros campos de texto
                        campos_texto = [
                            'ENTIDAD_LEGAL', 'NOMBRE_BANCO', 'EXTRACTO_NUM',
                            'EXT_LINEA_NUM', 'EXT_TIPO_TRX', 'TRX_CODE',
                            'EXT_LIN_ID', 'STATUS', 'TRX_TEXT', 'NRO_DOCUMENTO',
                            'SOCIO_COMERCIAL', 'COMENTARIO_ESPERADO', 'MES'
                        ]
                        
                        for col in campos_texto:
                            if col in df_procesado.columns:
                                df_procesado[col] = df_procesado[col].fillna('').astype(str)
                        
                        # Filtrar filas válidas
                        df_procesado = df_procesado[
                            (df_procesado['EXTRACTO_FECHA'].notna()) | 
                            (df_procesado['EXT_LIN_MONTO'] != 0)
                        ]
                        
                        if len(df_procesado) > 0:
                            df_procesado = df_procesado[self.columnas_finales]
                            dfs_archivo.append(df_procesado)
                            self.log.append(f"   ✓ {nombre_hoja}: {len(df_procesado)} registros")
                        else:
                            self.log.append(f"   ⊘ {nombre_hoja}: Sin datos válidos")
                    else:
                        self.log.append(f"   ⊘ {nombre_hoja}: Sin datos")
                        
                except Exception as e:
                    self.log.append(f"   ✗ {nombre_hoja}: Error - {str(e)}")
            
            return dfs_archivo
            
        except Exception as e:
            self.log.append(f"   ✗ Error al procesar archivo: {str(e)}")
            return []
    
    def consolidar(self, archivos_dict):
        """
        Consolida múltiples archivos
        archivos_dict: dict con {nombre_archivo: bytes}
        """
        self.log = []
        self.log.append("🚀 Iniciando consolidación...")
        self.log.append("")
        
        todos_dfs = []
        
        for nombre_archivo, archivo_bytes in archivos_dict.items():
            dfs = self.procesar_archivo(archivo_bytes, nombre_archivo)
            todos_dfs.extend(dfs)
        
        if not todos_dfs:
            self.log.append("")
            self.log.append("❌ No se pudieron procesar los archivos")
            return None
        
        # Consolidar
        self.df_consolidado = pd.concat(todos_dfs, ignore_index=True)
        
        # Eliminar duplicados
        filas_antes = len(self.df_consolidado)
        self.df_consolidado = self.df_consolidado.drop_duplicates(
            subset=['EXTRACTO_FECHA', 'CTA_NUMERO', 'EXT_LIN_MONTO', 'TRX_TEXT'],
            keep='first'
        )
        duplicados = filas_antes - len(self.df_consolidado)
        
        # Ordenar
        self.df_consolidado['FECHA_SORT'] = pd.to_datetime(
            self.df_consolidado['EXTRACTO_FECHA'],
            format='%d/%m/%Y',
            errors='coerce'
        )
        self.df_consolidado = self.df_consolidado.sort_values(
            by=['FECHA_SORT', 'NOMBRE_BANCO']
        ).drop('FECHA_SORT', axis=1)
        
        self.log.append("")
        self.log.append("=" * 50)
        self.log.append(f"✅ CONSOLIDACIÓN EXITOSA")
        self.log.append(f"📊 Total de registros: {len(self.df_consolidado):,}")
        self.log.append(f"🔄 Duplicados eliminados: {duplicados}")
        self.log.append(f"📁 Archivos procesados: {len(archivos_dict)}")
        self.log.append("=" * 50)
        
        return self.df_consolidado
    
    def generar_excel(self):
        """Genera archivo Excel en memoria con formato texto para cuentas"""
        if self.df_consolidado.empty:
            return None
        
        output = io.BytesIO()
        
        # Usar xlsxwriter para tener control sobre el formato
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            self.df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
            
            # Obtener el workbook y worksheet
            workbook = writer.book
            worksheet = writer.sheets['Consolidado']
            
            # Formato de texto para las cuentas
            text_format = workbook.add_format({'num_format': '@'})
            
            # Aplicar formato texto a columnas CTA_BANCO (D) y CTA_NUMERO (E)
            # Columnas: A=0, B=1, C=2, D=3, E=4
            worksheet.set_column('D:D', 15, text_format)  # CTA_BANCO
            worksheet.set_column('E:E', 15, text_format)  # CTA_NUMERO
            
            # Ajustar anchos de otras columnas
            worksheet.set_column('A:A', 10)  # MES
            worksheet.set_column('B:B', 20)  # ENTIDAD_LEGAL
            worksheet.set_column('C:C', 20)  # NOMBRE_BANCO
            worksheet.set_column('F:F', 12)  # EXTRACTO_NUM
            worksheet.set_column('G:G', 12)  # EXTRACTO_FECHA
            worksheet.set_column('K:K', 15)  # EXT_LIN_MONTO
            worksheet.set_column('N:N', 30)  # TRX_TEXT
        
        output.seek(0)
        return output


# INTERFAZ DE STREAMLIT
def main():
    
    # Header
    st.title("📊 Consolidador de Extractos Bancarios")
    st.markdown("### Despegar - Sistema de Consolidación")
    st.markdown("---")
    
    # Instrucciones
    with st.expander("📖 Instrucciones de uso", expanded=False):
        st.markdown("""
        **¿Cómo usar el consolidador?**
        
        1. **Selecciona archivos**: Haz clic en "Browse files" y selecciona todos tus archivos Excel de extractos
        2. **Sube múltiples archivos**: Puedes seleccionar varios archivos a la vez
        3. **Procesa**: Haz clic en "🚀 Consolidar Extractos"
        4. **Descarga**: Una vez procesado, descarga el archivo consolidado
        
        **Estructura requerida:**
        - Los archivos deben ser Excel (.xlsx o .xls)
        - Cada hoja puede contener extractos de diferentes períodos
        - Las hojas que empiecen con "_" serán ignoradas
        - Las columnas deben estar en el orden: ENTIDAD_LEGAL, NOMBRE_BANCO, CTA_BANCO, CTA_NUMERO, etc.
        
        **Características:**
        - ✅ Calcula automáticamente la columna MES (MM/YYYY)
        - ✅ Aplica signos correctos a montos (DEBIT negativo, CREDIT positivo)
        - ✅ Preserva ceros en números de cuenta (formato texto)
        - ✅ Elimina duplicados automáticamente
        - ✅ Ordena por fecha
        """)
    
    st.markdown("---")
    
    # Uploader
    st.subheader("1️⃣ Selecciona tus archivos Excel")
    uploaded_files = st.file_uploader(
        "Sube uno o más archivos Excel",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Puedes seleccionar múltiples archivos a la vez"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} archivo(s) cargado(s)")
        
        # Mostrar lista de archivos
        with st.expander("📁 Archivos cargados", expanded=True):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name} ({file.size / 1024:.2f} KB)")
        
        st.markdown("---")
        
        # Botón de consolidar
        st.subheader("2️⃣ Procesar archivos")
        
        if st.button("🚀 Consolidar Extractos", type="primary", use_container_width=True):
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Inicializando consolidador...")
            progress_bar.progress(10)
            
            # Crear diccionario de archivos
            archivos_dict = {}
            for file in uploaded_files:
                archivos_dict[file.name] = io.BytesIO(file.read())
            
            status_text.text("Procesando archivos...")
            progress_bar.progress(30)
            
            # Consolidar
            consolidador = ConsolidadorExtractosWeb()
            df = consolidador.consolidar(archivos_dict)
            
            progress_bar.progress(70)
            
            if df is not None:
                status_text.text("Generando archivo Excel...")
                progress_bar.progress(90)
                
                # Generar Excel
                excel_bytes = consolidador.generar_excel()
                
                progress_bar.progress(100)
                status_text.text("✅ ¡Completado!")
                
                st.markdown("---")
                
                # Mostrar log
                st.subheader("📋 Log de procesamiento")
                log_container = st.container()
                with log_container:
                    for linea in consolidador.log:
                        st.text(linea)
                
                st.markdown("---")
                
                # Estadísticas
                st.subheader("📊 Estadísticas del consolidado")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Registros", f"{len(df):,}")
                
                with col2:
                    st.metric("Meses", df['MES'].nunique())
                
                with col3:
                    st.metric("Bancos", df['NOMBRE_BANCO'].nunique())
                
                with col4:
                    st.metric("Cuentas", df['CTA_NUMERO'].nunique())
                
                # Métricas financieras
                st.markdown("#### 💰 Análisis Financiero")
                
                col1, col2, col3 = st.columns(3)
                
                creditos = df[df['EXT_LIN_MONTO'] > 0]
                debitos = df[df['EXT_LIN_MONTO'] < 0]
                balance = df['EXT_LIN_MONTO'].sum()
                
                with col1:
                    st.metric(
                        "Créditos", 
                        f"${creditos['EXT_LIN_MONTO'].sum():,.2f}",
                        delta=f"{len(creditos):,} movimientos"
                    )
                
                with col2:
                    st.metric(
                        "Débitos", 
                        f"${debitos['EXT_LIN_MONTO'].sum():,.2f}",
                        delta=f"{len(debitos):,} movimientos"
                    )
                
                with col3:
                    st.metric(
                        "Balance Neto", 
                        f"${balance:,.2f}",
                        delta="Total"
                    )
                
                st.markdown("---")
                
                # Preview de datos
                st.subheader("👁️ Vista previa del consolidado")
                st.dataframe(
                    df.head(20),
                    use_container_width=True,
                    height=400
                )
                
                st.markdown("---")
                
                # Botón de descarga
                st.subheader("3️⃣ Descargar archivo consolidado")
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"CONSOLIDADO_EXTRACTOS_{timestamp}.xlsx"
                
                st.download_button(
                    label="⬇️ Descargar Consolidado (Excel)",
                    data=excel_bytes,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                st.success("✅ El archivo está listo para descargar y subir a Google Sheets")
                
            else:
                st.error("❌ Error al consolidar los archivos. Revisa el log de procesamiento.")
                
                # Mostrar log de errores
                st.subheader("📋 Log de errores")
                for linea in consolidador.log:
                    st.text(linea)
    
    else:
        st.info("👆 Por favor, sube uno o más archivos Excel para comenzar")
    
    # Footer
    st.markdown("---")
    st.markdown("**Consolidador de Extractos Bancarios** | Despegar © 2024")


if __name__ == "__main__":
    main()
