import streamlit as st
import pandas as pd
import io
from io import BytesIO

st.set_page_config(page_title="Procesos ANS9", page_icon="🧪")

def to_excel_multiple_sheets(dfdto, dfpcl):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        dfdto.to_excel(writer, index=False, sheet_name="DTO")
        dfpcl.to_excel(writer, index=False, sheet_name="PCL")
    output.seek(0)
    return output.getvalue()

st.sidebar.title("🔧 Menú de Procesos")
opcion = st.sidebar.selectbox(
    "Selecciona el proceso que quieres ejecutar:",
    ["📅 Filtro por Fechas de Corte y Termino", "📊 Base Courier"]
)

# PROCESO 1 - FILTRO POR FECHAS Y TERMINO --
if opcion == "📅 Filtro por Fechas de Corte y Termino":
    st.title("📅 Filtro ANS9 - Fechas de Corte y Término")

    archivo = st.file_uploader("📤 Sube el archivo Excel (.xlsx)", type=["xlsx"])

    if archivo is not None:
        xls = pd.ExcelFile(archivo)

        if "DTO" in xls.sheet_names and "PCL" in xls.sheet_names:
            df_dto = pd.read_excel(xls, sheet_name="DTO")
            df_pcl = pd.read_excel(xls, sheet_name="PCL")

            # Convertir FECHA_VISADO a datetime, por si acaso
            df_dto["FECHA_VISADO"] = pd.to_datetime(df_dto["FECHA_VISADO"], errors='coerce')
            df_pcl["FECHA_VISADO"] = pd.to_datetime(df_pcl["FECHA_VISADO"], errors='coerce')

            # Calcular rango global de fechas para los inputs
            fecha_min = min(df_dto["FECHA_VISADO"].min(), df_pcl["FECHA_VISADO"].min())
            fecha_max = max(df_dto["FECHA_VISADO"].max(), df_pcl["FECHA_VISADO"].max())

            col1, col2 = st.columns(2)
            with col1:
                fecha_inicio = st.date_input("📅 Fecha Inicio", value=fecha_min)
            with col2:
                fecha_fin = st.date_input("📅 Fecha Fin", value=fecha_max)

            ejecutar_filtro = st.button("🔍 Ejecutar Filtro", key="filtro_ans9", use_container_width=True)

            if ejecutar_filtro:
                # Filtrar solo por fechas
                dto_filtrado_fechas = df_dto[
                    (df_dto["FECHA_VISADO"] >= pd.to_datetime(fecha_inicio)) &
                    (df_dto["FECHA_VISADO"] <= pd.to_datetime(fecha_fin))
                ].copy()

                pcl_filtrado_fechas = df_pcl[
                    (df_pcl["FECHA_VISADO"] >= pd.to_datetime(fecha_inicio)) &
                    (df_pcl["FECHA_VISADO"] <= pd.to_datetime(fecha_fin))
                ].copy()

                # Normalizar columna TERMINOS para evitar problemas
                dto_filtrado_fechas["TERMINOS"] = dto_filtrado_fechas["TERMINOS"].astype(str).str.strip().str.upper()
                pcl_filtrado_fechas["TERMINOS"] = pcl_filtrado_fechas["TERMINOS"].astype(str).str.strip().str.upper()

                # Filtrar por TERMINOS = "FUERA DE TERMINOS"
                dto_filtrado_fechas_fuera = dto_filtrado_fechas[
                    dto_filtrado_fechas["TERMINOS"] == "FUERA DE TERMINOS"
                ]

                pcl_filtrado_fechas_fuera = pcl_filtrado_fechas[
                    pcl_filtrado_fechas["TERMINOS"] == "FUERA DE TERMINOS"
                ]

                st.success(f"✅ Filtrado aplicado para fechas: {fecha_inicio} - {fecha_fin}")

                col1, col2 = st.columns(2)

                with col1:
                    data_solo_fechas = to_excel_multiple_sheets(dto_filtrado_fechas, pcl_filtrado_fechas)
                    st.download_button(
                        label="Descargar Archivo Solo fechas",
                        data=data_solo_fechas,
                        file_name="general.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with col2:
                    data_fechas_fuera = to_excel_multiple_sheets(dto_filtrado_fechas_fuera, pcl_filtrado_fechas_fuera)
                    st.download_button(
                        label="Descargar Filtro Fechas y fuera de término",
                        data=data_fechas_fuera,
                        file_name="filtrado_fechas_fuera_termino.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            else:
                st.info("📥 Selecciona las fechas y presiona 'Ejecutar Filtro' para procesar los datos.")

        else:
            st.error("❌ No se encontraron hojas DTO y PCL en el archivo.")
    else:
        st.info("📥 Por favor, sube un archivo Excel para comenzar.")



# PROCESO 2 - BASE COURIER --
elif opcion == "📊 Base Courier":
    st.title("📊 BASE COURIER")

    archivo1 = st.file_uploader(
        "📤 Sube el archivo Excel Base General (.xlsx)", 
        type=["xlsx"], 
        key="base_courier_1"
    )
    archivo2 = st.file_uploader(
        "📤 Sube el archivo Excel Base Courier (.xlsx)", 
        type=["xlsx"], 
        key="base_courier_2"
    )

    if archivo1 is not None and archivo2 is not None:
        xls1 = pd.ExcelFile(archivo1)
        if "DTO" in xls1.sheet_names and "PCL" in xls1.sheet_names:
            df_dto = pd.read_excel(xls1, sheet_name="DTO")
            df_pcl = pd.read_excel(xls1, sheet_name="PCL")

            # ✅ Agregamos la columna CALIFICACION
            df_dto["CALIFICACION"] = "DTO"
            df_pcl["CALIFICACION"] = "PCL"
        else:
            st.error("❌ El archivo 1 debe contener hojas DTO y PCL.")
            st.stop()

        xls2 = pd.ExcelFile(archivo2)
        if "COURIER" in xls2.sheet_names and "MENSAJERO" in xls2.sheet_names:
            df_courier = pd.read_excel(xls2, sheet_name="COURIER")
            df_mensajero = pd.read_excel(xls2, sheet_name="MENSAJERO")
        else:
            st.error("❌ El archivo 2 debe contener hojas COURIER y MENSAJERO.")
            st.stop()

        df_base_general = pd.concat([df_dto, df_pcl], ignore_index=True)
        df_base_courier = pd.concat([df_courier, df_mensajero], ignore_index=True)

        ids_1 = set(df_base_general["ID_FURAT_FUREP"].dropna().unique())
        ids_2 = set(df_base_courier["ID DEL SINIESTRO"].dropna().unique())
        ids_comunes = ids_1.intersection(ids_2)

        mask = (
            df_base_general["ID_FURAT_FUREP"].isin(ids_comunes) &
            (df_base_general["ESTADO_INFORME"].str.upper() == "PENDIENTE")
        )

        df_base_general_mod = df_base_general.copy()
        df_base_general_mod.loc[mask, "ESTADO_INFORME"] = "PENDIENTE ENTREGA DE GUIA"

        # 🚀 Agregar columnas vacías nuevas
        for col in ["OPORTUNIDAD FINAL", "OBSERVACIÓN", "DEFINICIÓN"]:
            if col not in df_base_general_mod.columns:
                df_base_general_mod[col] = ""

        st.write(f"🟢 IDs comunes: {len(ids_comunes)}")
        st.write(f"🟠 Registros con ESTADO_INFORME 'PENDIENTE' actualizados: {mask.sum()}")

        # Mostrar TODOS los datos en pantalla con los cambios aplicados
        st.dataframe(df_base_general_mod)

        # Normalizamos TERMINOS para filtrar los que queremos en Excel
        df_base_general_mod["TERMINOS_NORM"] = df_base_general_mod["TERMINOS"].astype(str).str.strip().str.upper()

        # Filtrar SOLO los registros que tengan TERMINOS = "FUERA DE TERMINO"
        df_para_descarga = df_base_general_mod[df_base_general_mod["TERMINOS_NORM"] == "FUERA DE TERMINOS"]

        if df_para_descarga.empty:
            st.info("No hay registros con TERMINOS = 'FUERA DE TERMINO' para descargar.")
        else:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                notificadores = df_para_descarga["NOTIFICADOR"].dropna().unique()

                columnas_renombrar = {
                    "ID_FURAT_FUREP": "ID DEL SINIESTRO",
                    "FECHA_VISADO": "FECHA VISADO",
                    "NOMBRE_COMITE": "NOMBRE COMITE",
                    "ID_TRABAJADOR": "ID TRABAJADOR",
                    "FECHA_NOTIFICACION": "FECHA NOTIFICACION",
                    "RADICADO_SALIDA": "RAD DE SALIDA",
                    "FECHA_RADICACION": "FECHA RADICACION",
                    "NOTIFICADOR": "NOTIFICADOR",
                    "EMPRESA": "EMPRESA",
                    "DIAS TRANSCURRIDOS HABILES": "DIAS TRANSCURRIDOS HABILES",
                    "ESTADO_INFORME": "ESTADO INFORME",
                    "CALIFICACION": "CALIFICACION"
                }

                columnas_a_incluir = [
                    "ID_FURAT_FUREP",
                    "FECHA_VISADO",
                    "NOMBRE_COMITE",
                    "ID_TRABAJADOR",
                    "FECHA_NOTIFICACION",
                    "RADICADO_SALIDA",
                    "FECHA_RADICACION",
                    "NOTIFICADOR",
                    "EMPRESA",
                    "DIAS TRANSCURRIDOS HABILES",
                    "ESTADO_INFORME",
                    "CALIFICACION",
                    "OPORTUNIDAD FINAL",
                    "OBSERVACIÓN",
                    "DEFINICIÓN"
                ]

                for notif in notificadores:
                    df_notif = df_para_descarga[df_para_descarga["NOTIFICADOR"] == notif]

                    cols_existentes = [c for c in columnas_a_incluir if c in df_notif.columns]
                    df_export = df_notif[cols_existentes].copy()

                    cols_renombrar_existentes = {k: v for k, v in columnas_renombrar.items() if k in df_export.columns}
                    df_export.rename(columns=cols_renombrar_existentes, inplace=True)

                    df_export.insert(0, "Nº", range(1, len(df_export) + 1))

                    sheet_name = str(notif)[:31]
                    df_export.to_excel(writer, sheet_name=sheet_name, index=False)

            output.seek(0)

            st.download_button(
                label="📥 Descargar Excel con registros FUERA DE TERMINO",
                data=output,
                file_name="fuera_de_termino.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
