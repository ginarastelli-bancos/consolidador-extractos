
"""
CONSOLIDADOR DE EXTRACTOS BANCARIOS - DESPEGAR
Busca columnas POR NOMBRE sin importar el orden en el Excel original
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Extractos - Despegar",
    page_icon="📊",
    layout="wide"
)

class ConsolidadorExtractosWeb:
    
    def __init__(self):
        # Columnas requeridas EN ORDEN FINAL
        self.columnas_orden_final = [
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
        
        # Columnas requeridas (sin MES porque se calcula)
        self.columnas_requeridas = [
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
        
        self.df_consolidado = pd.DataFrame()
        self.log_procesamiento = []
    
    def calcular_mes(self, fecha):
        """Calcula MES en formato MM/YYYY desde EXTRACTO_FECHA"""
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
        """Normaliza fecha a DD/MM/AAAA"""
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
        """Aplica signo al monto según tipo de transacción"""
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
        valor_str = str(valor)
        if '.' in valor_str:
            valor_str = valor_str.split('.')[0]
        return valor_str
    
    def buscar_columna(self, df, nombre_columna):
        """
        Busca una columna por nombre exacto o similar
        Retorna el nombre de la columna encontrada o None
        """
        # Primero buscar nombre exacto
        if nombre_columna in df.columns:
            return nombre_columna
        
        # Buscar ignorando mayúsculas/minúsculas
        for col in df.columns:
            if str(col).upper() == nombre_columna.upper():
                return col
        
        # Buscar nombre similar (sin espacios, guiones, etc.)
        nombre_limpio = nombre_columna.upper().replace('_', '').replace(' ', '').replace('-', '')
        for col in df.columns:
            col_limpio = str(col).upper().replace('_', '').replace(' ', '').replace('-', '')
            if col_limpio == nombre_limpio:
                return col
        
        return None
    
    def extraer_columnas_por_nombre(self, df, nombre_archivo, nombre_hoja):
        """
        Extrae columnas buscándolas por nombre
        """
        df_extraido = pd.DataFrame()
        columnas_no_encontradas = []
        columnas_encontradas = []
        
        for nombre_col in self.columnas_requeridas:
            col_encontrada = self.buscar_columna(df, nombre_col)
            
            if col_encontrada:
                df_extraido[nombre_col] = df[col_encontrada]
                columnas_encontradas.append(f"{nombre_col} → {col_encontrada}")
            else:
                df_extraido[nombre_col] = ''
                columnas_no_encontradas.append(nombre_col)
        
        # Log de columnas
        if columnas_encontradas:
            self.log_procesamiento.append({
                'archivo': nombre_archivo,
                'hoja': nombre_hoja,
                'mensaje': f"Columnas encontradas: {len(columnas_encontradas)}"
            })
        
        if columnas_no_encontradas:
            self.log_procesamiento.append({
                'archivo': nombre_archivo,
                'hoja': nombre_hoja,
                'mensaje': f"⚠️ Columnas NO encontradas: {', '.join(columnas_no_encontradas)}"
            })
        
        return df_extraido
    
    def procesar_hoja(self, archivo, nombre_hoja, df_hoja):
        """Procesa una hoja del Excel"""
        try:
            if df_hoja.empty:
                return None
            
            # Extraer columnas por nombre
            df_procesado = self.extraer_columnas_por_nombre(df_hoja, archivo, nombre_hoja)
            
            # Eliminar filas vacías
            df_procesado = df_procesado.dropna(how='all')
            
            if len(df_procesado) == 0:
                return None
            
            # Normalizar fechas
            df_procesado['EXTRACTO_FECHA'] = df_procesado['EXTRACTO_FECHA'].apply(self.normalizar_fecha)
            
            # Calcular MES
            df_procesado['MES'] = df_procesado['EXTRACTO_FECHA'].apply(self.calcular_mes)
            
            # Aplicar signos a montos
            df_procesado['EXT_LIN_MONTO'] = df_procesado.apply(
                lambda row: self.aplicar_signo_monto(row['EXT_LIN_MONTO'], row['EXT_TIPO_TRX']),
                axis=1
            )
            
            # Convertir cuentas a texto
            df_procesado['CTA_BANCO'] = df_procesado['CTA_BANCO'].apply(self.formato_texto_cuenta)
            df_procesado['CTA_NUMERO'] = df_procesado['CTA_NUMERO'].apply(self.formato_texto_cuenta)
            
            # Limpiar campos de texto
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
            
            if len(df_procesado) == 0:
                return None
            
            # Reordenar columnas en el orden final
            df_procesado = df_procesado[self.columnas_orden_final]
            
            self.log_procesamiento.append({
                'archivo': archivo,
                'hoja': nombre_hoja,
                'filas': len(df_procesado),
                'estado': '✓ OK'
            })
            
            return df_procesado
            
        except Exception as e:
            self.log_procesamiento.append({
                'archivo': archivo,
                'hoja': nombre_hoja,
                'filas': 0,
                'estado': f'✗ Error: {str(e)}'
            })
            return None
    
    def consolidar(self, archivos_dict):
        """Consolida múltiples archivos"""
        self.log_procesamiento = []
        self.log_procesamiento.append({'tipo': 'inicio', 'mensaje': '🚀 Iniciando consolidación...'})
        
        todos_dfs = []
        
        for nombre_archivo, archivo_bytes in archivos_dict.items():
            try:
                self.log_procesamiento.append({'tipo': 'archivo', 'archivo': nombre_archivo})
                
                xls = pd.ExcelFile(archivo_bytes)
                hojas = xls.sheet_names
                
                for nombre_hoja in hojas:
                    if nombre_hoja.startswith('_'):
                        continue
                    
                    df_hoja = pd.read_excel(archivo_bytes, sheet_name=nombre_hoja)
                    df_procesado = self.procesar_hoja(nombre_archivo, nombre_hoja, df_hoja)
                    
                    if df_procesado is not None and len(df_procesado) > 0:
                        todos_dfs.append(df_procesado)
                        
            except Exception as e:
                self.log_procesamiento.append({
                    'archivo': nombre_archivo,
                    'estado': f'✗ Error al leer archivo: {str(e)}'
                })
        
        if not todos_dfs:
            self.log_procesamiento.append({'tipo': 'error', 'mensaje': '❌ No se pudieron extraer datos'})
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
        
        self.log_procesamiento.append({
            'tipo': 'resumen',
            'total': len(self.df_consolidado),
            'duplicados': duplicados,
            'archivos': len(archivos_dict)
        })
        
        return self.df_consolidado
    
    def generar_excel(self):
        """Genera Excel con formato texto en cuentas"""
        if self.df_consolidado.empty:
            return None
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            self.df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Consolidado']
            
            # Formato texto para cuentas
            text_format = workbook.add_format({'num_format': '@'})
            
            # CTA_BANCO (columna D, índice 3) y CTA_NUMERO (columna E, índice 4)
            worksheet.set_column('D:D', 15, text_format)
            worksheet.set_column('E:E', 15, text_format)
            
            # Ajustar anchos
            worksheet.set_column('A:A', 10)  # MES
            worksheet.set_column('B:B', 20)  # ENTIDAD_LEGAL
            worksheet.set_column('C:C', 20)  # NOMBRE_BANCO
            worksheet.set_column('G:G', 12)  # EXTRACTO_FECHA
            worksheet.set_column('K:K', 15)  # EXT_LIN_MONTO
            worksheet.set_column('N:N', 30)  # TRX_TEXT
        
        output.seek(0)
        return output


# INTERFAZ STREAMLIT
def main():
    st.title("📊 Consolidador de Extractos Bancarios")
    st.markdown("### Despegar - Sistema de Consolidación")
    st.markdown("---")
    
    # Instrucciones
    with st.expander("📖 Instrucciones de uso", expanded=False):
        st.markdown("""
        **¿Cómo usar el consolidador?**
        
        1. Sube tus archivos Excel (pueden tener columnas en cualquier orden)
        2. El sistema buscará las columnas por nombre
        3. Generará el consolidado con las columnas ordenadas correctamente
        
        **Columnas requeridas (pueden estar en cualquier orden):**
        - ENTIDAD_LEGAL
        - NOMBRE_BANCO
        - CTA_BANCO
        - CTA_NUMERO
        - EXTRACTO_NUM
        - EXTRACTO_FECHA
        - EXT_LINEA_NUM
        - EXT_TIPO_TRX
        - TRX_CODE
        - EXT_LIN_MONTO
        - EXT_LIN_ID
        - STATUS
        - TRX_TEXT
        - NRO_DOCUMENTO
        - SOCIO_COMERCIAL
        - COMENTARIO_ESPERADO
        
        **El consolidado incluirá:**
        - Columna MES calculada automáticamente
        - Todas las columnas en el orden correcto
        - Formato texto en CTA_BANCO y CTA_NUMERO (preserva ceros)
        """)
    
    st.markdown("---")
    
    # Uploader
    st.subheader("1️⃣ Selecciona tus archivos Excel")
    uploaded_files = st.file_uploader(
        "Sube uno o más archivos Excel",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} archivo(s) cargado(s)")
        
        with st.expander("📁 Archivos cargados", expanded=True):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name} ({file.size / 1024:.2f} KB)")
        
        st.markdown("---")
        st.subheader("2️⃣ Procesar archivos")
        
        if st.button("🚀 Consolidar Extractos", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Inicializando...")
            progress_bar.progress(10)
            
            archivos_dict = {}
            for file in uploaded_files:
                archivos_dict[file.name] = io.BytesIO(file.read())
            
            status_text.text("Procesando...")
            progress_bar.progress(30)
            
            consolidador = ConsolidadorExtractosWeb()
            df = consolidador.consolidar(archivos_dict)
            
            progress_bar.progress(70)
            
            if df is not None:
                status_text.text("Generando Excel...")
                progress_bar.progress(90)
                
                excel_bytes = consolidador.generar_excel()
                progress_bar.progress(100)
                status_text.text("✅ ¡Completado!")
                
                st.markdown("---")
                
                # Log
                st.subheader("📋 Log de procesamiento")
                for log_entry in consolidador.log_procesamiento:
                    if 'tipo' in log_entry:
                        if log_entry['tipo'] == 'inicio':
                            st.info(log_entry['mensaje'])
                        elif log_entry['tipo'] == 'archivo':
                            st.write(f"📄 **{log_entry['archivo']}**")
                        elif log_entry['tipo'] == 'resumen':
                            st.success(f"✅ Total: {log_entry['total']:,} registros | Duplicados eliminados: {log_entry['duplicados']}")
                    elif 'archivo' in log_entry and 'hoja' in log_entry:
                        if log_entry.get('estado') == '✓ OK':
                            st.write(f"  └─ {log_entry['hoja']}: {log_entry['filas']} registros ✓")
                        else:
                            st.write(f"  └─ {log_entry['hoja']}: {log_entry.get('estado', 'Sin datos')}")
                        if 'mensaje' in log_entry:
                            st.write(f"     {log_entry['mensaje']}")
                
                st.markdown("---")
                
                # Estadísticas
                st.subheader("📊 Estadísticas")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Registros", f"{len(df):,}")
                with col2:
                    st.metric("Meses", df['MES'].nunique())
                with col3:
                    st.metric("Bancos", df['NOMBRE_BANCO'].nunique())
                with col4:
                    st.metric("Cuentas", df['CTA_NUMERO'].nunique())
                
                # Vista previa
                st.subheader("👁️ Vista previa")
                st.dataframe(df.head(20), use_container_width=True, height=400)
                
                st.markdown("---")
                
                # Descarga
                st.subheader("3️⃣ Descargar consolidado")
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
                
                st.success("✅ Listo para descargar")
                
            else:
                st.error("❌ Error al consolidar los archivos. Revisa el log de procesamiento.")
                
                st.subheader("📋 Log de errores")
                for log_entry in consolidador.log_procesamiento:
                    if 'tipo' in log_entry and log_entry['tipo'] == 'error':
                        st.error(log_entry['mensaje'])
                    elif 'estado' in log_entry and '✗' in log_entry['estado']:
                        st.warning(f"{log_entry.get('archivo', '')}: {log_entry['estado']}")
    
    else:
        st.info("👆 Sube uno o más archivos Excel para comenzar")
    
    st.markdown("---")
    st.markdown("**Consolidador de Extractos Bancarios** | Despegar © 2024")


if __name__ == "__main__":
    main()
