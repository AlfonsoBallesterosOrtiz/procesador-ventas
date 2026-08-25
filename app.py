import streamlit as st
import pandas as pd
import io

# Título de la aplicación
st.title("Procesador y Validador de Ventas")

# 1. Componente para subir el archivo Excel
uploaded_file = st.file_uploader("Sube tu archivo de ventas (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    if 'MONTO' in df.columns:
        
        # --- Lógica de procesamiento ---
        def procesar_monto(valor):
            if pd.isna(valor):
                return "SIN DATO"
            try:
                return float(valor)
            except (ValueError, TypeError):
                return "NECESITA REVISION"

        def calcular_iva(valor):
            if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                return round(valor * 0.16, 2)
            return valor

        def calcular_total(valor, iva):
            if isinstance(valor, (int, float)) and isinstance(iva, (int, float)):
                return round(valor + iva, 2)
            return valor

        df['MONTO'] = df['MONTO'].apply(procesar_monto)
        df['IVA'] = df['MONTO'].apply(calcular_iva)
        df['TOTAL'] = df.apply(lambda row: calcular_total(row['MONTO'], row['IVA']), axis=1)

        # --- Métricas y Resumen ---
        total_filas = len(df)
        filas_revision_df = df[df['MONTO'].isin(["SIN DATO", "NECESITA REVISION"])]
        filas_revision_cant = len(filas_revision_df)
        
        es_valido = df['MONTO'].apply(lambda x: isinstance(x, (int, float)) and not isinstance(x, bool))
        suma_monto_valido = df.loc[es_valido, 'MONTO'].sum()
        suma_iva_valido = df.loc[es_valido, 'IVA'].sum()

        st.subheader("Resumen del Procesamiento")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Filas", total_filas)
        col2.metric("Requieren Revisión", filas_revision_cant)
        col3.metric("Ventas Válidas / IVA", f"${suma_monto_valido:,.2f} / ${suma_iva_valido:,.2f}")

        # --- Reglas de Estilo CSS (Fondo) ---
        def resaltar_filas_invalidas(row):
            if row['MONTO'] in ["SIN DATO", "NECESITA REVISION"]:
                return ['background-color: #fff9c4'] * len(row)  # Amarillo suave
            return [''] * len(row)

        # --- Regla de Formato numérico a texto (Moneda) ---
        def formato_moneda(val):
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return f"${val:,.2f}"
            return str(val)

        # Diccionario de formato aplicado ÚNICAMENTE a las columnas financieras
        dict_formato_moneda = {
            'MONTO': formato_moneda,
            'IVA': formato_moneda,
            'TOTAL': formato_moneda
        }

        # --- Requisito 1: Tabla de filas con conflicto (si existen) ---
        if filas_revision_cant > 0:
            st.markdown(
                "<h3 style='color: #b78103; margin-top: 20px;'>⚠️ Filas que requieren revisión manual</h3>", 
                unsafe_allow_html=True
            )
            
            # Formatear la tabla de revisión sin tocar FECHA ni CLIENTE
            df_revision_styled = filas_revision_df.style.format(dict_formato_moneda)
            st.dataframe(df_revision_styled, use_container_width=True)

        # --- Requisito 2, 3 y 4: Tabla completa estilizada ---
        st.subheader("Detalle de Ventas Procesadas")
        
        # Combinar resaltado CSS (.apply) con formateo visual por columna (.format)
        df_completo_styled = (
            df.style
            .apply(resaltar_filas_invalidas, axis=1)
            .format(dict_formato_moneda)
        )
        st.dataframe(df_completo_styled, use_container_width=True)

        # --- Preparar archivo Excel para descarga ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 Descargar Ventas Revisadas (.xlsx)",
            data=buffer,
            file_name="Ventas_revisadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("El archivo cargado no contiene la columna obligatoria 'MONTO'.")