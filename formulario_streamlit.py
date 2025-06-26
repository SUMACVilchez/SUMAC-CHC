import streamlit as st
import pandas as pd
import os
import re
import yagmail
from datetime import datetime
from zipfile import ZipFile

# CONFIGURACIÓN GENERAL
st.set_page_config(page_title="Formulario Huella de Carbono - SUMAC", layout="wide")
os.makedirs("datos", exist_ok=True)
os.makedirs("evidencias", exist_ok=True)

estructura = {
    'A1_Vehículos_propios_móviles': {'Tipo de vehículo': 'Auto,Camioneta,Furgón,SUV,Motocicleta,Otro', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído', 'Consumo anual': None, 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Generador_Electri_móvile': {'Tipo de vehículo': None, 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído,Eléctrico', 'Consumo anual': None, 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Maquinari_propios_móvile': {'Tipo de vehículo': 'Grúa,Escabadora,Montacarga', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído,Eléctrico', 'Consumo anual': None, 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Equipos_estacionarios': {'Tipo de equipo': 'Elevador,Prensa,Calderas,Trituradora,Motobombas,Generador,Cocina', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído', 'Consumo anual': None, 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Aire_acondicionado': {'Equipo': None, 'Tipo de gas': 'R-22,R44,R410,HCFC-22,R-410A,R-134a,R-407C,R-404A,R-32,R-600a', '¿Presentó fugas y/o recargas?': 'Fuga,Recarga', 'Capacidad': None, 'Unidad': 'kg', 'Cantidad de recargas': None},
    'A1_Extintores': {'Equipo': None, '¿Presentó fugas y/o recargas?': 'Fuga,Recarga', 'Capacidad': None, 'Unidad': 'kg', 'Cantidad de recargas': None},
    'A2_Electricidad': {'Consumo anual (KWh)': None, 'Año': '2022,2023,2024'},
    'A3_Transporte__casa__trabajo': {'Alcance 3: Emisiones indirectas de GEI de productos usados por la organización': None},
    'A3_Papelería': {'Descripción': None, 'Largo (cm)': None, 'Ancho (cm)': None, 'Gramaje (gr/m2)': None, 'Cantidad': None, 'Unidad': None},
    'A3_Transporte_contratado': {'Tipo de vehículo': 'Bus,Van,Taxi,Otro', 'Tipo de combustible': 'Diesel,Gasolina,GNV,GLP,Otro', 'Consumo de combustible anual (litros)': None, 'Gastos por el servicio': None, 'Destino incial': None, 'Destino final': None, 'Kilómetros recorridos': None, 'Año': '2022,2023,2024'},
    'A3_Consumo_de_agua': {'Uso': None, 'Consumo anual (m3)': None, 'Año': '2022,2023,2024'},
    'A3_Materiales_Bien_y_Servicio': {'Tipo de bien y servicio': 'Bien adquirido,Bien capital,Servicio adquirido,Servicio capital', 'Descripción': None, 'Tipo de moneda': 'Peso argentino,Boliviano,Real,Peso chileno,Peso colombiano,Dólar,Guaraní,Sol,Peso uruguayo,Bolívar', 'Monto': None},
    'A3_Residuos': {'Tipo de residuo sólidos': 'Peligrosos,No peligrosos', 'Clasificación': 'Madera,Papel,Cartón,Comida,Textiles,Jardines,Plástico,Residuos urbanos,Lodos tratados,Lodos no tratados,Lodos industriales', 'Cantidad total': None, 'Unidad': 'Kg,Tn', 'Año': '2022,2023,2024'},
    'A3_Transporte_Residuos': {'Origen': None, 'Destino': None, 'Número de viajes anuales': None},
    'A3_Consumo_de_electricidad_loca': {'Consumo anual (KWh)': None, 'Año': '2022,2023,2024'}
}


def es_email_valido(email):
    return re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email)

def enviar_correo(destinatario, asunto, cuerpo, archivos):
    remitente = "avilchez@sumacinc.com"
    password = "xbna iizl vhta aync"  # contraseña generada
    yag = yagmail.SMTP(user=remitente, password=password)
    yag.send(to=destinatario, subject=asunto, contents=cuerpo, attachments=archivos)

# TÍTULO E INSTRUCCIONES
st.title("Formulario de Huella de Carbono - SUMAC")
with st.expander("📘 Instrucciones de uso"):
    st.markdown("""
    Este programa registra información por alcance y fuente de emisión según la siguiente estructura:

    - **A1**: Combustión móvil (vehículos u otros).
    - **A2**: Electricidad adquirida.
    - **A3**: Otros consumos o emisiones indirectas.

    **Pasos:**
    1. Llena los datos de tu organización.
    2. Elige una categoría del formulario.
    3. Llena los datos y sube las evidencias correspondientes.
    4. Repite para otras categorías si es necesario.
    5. Presiona **Enviar todo** para generar y enviar los archivos.
    """)

# FORMULARIO DATOS EMPRESA
if "datos_empresa" not in st.session_state:
    with st.form("form_empresa"):
        st.subheader("Datos de la Empresa")
        col1, col2 = st.columns(2)
        nombre = col1.text_input("Nombre de la empresa")
        ruc = col2.text_input("RUC o ID fiscal")
        pais = col1.selectbox("País", sorted(["Argentina", "Bolivia", "Chile", "Colombia", "Ecuador", "España", "México", "Paraguay", "Perú", "Uruguay", "Estados Unidos"]))
        responsable = col2.text_input("Responsable")
        email = st.text_input("Email del responsable")
        enviado = st.form_submit_button("Iniciar")

        if enviado:
            if not (nombre and ruc.isdigit() and responsable and es_email_valido(email)):
                st.error("Completa todos los campos correctamente.")
            else:
                st.session_state.datos_empresa = {
                    "Nombre": nombre, "RUC": ruc, "País": pais, "Responsable": responsable, "Email": email
                }
                st.session_state.entradas = {}
                st.success("Datos guardados. Continúa llenando el formulario.")

# FORMULARIO POR CATEGORÍA
if "datos_empresa" in st.session_state:
    st.markdown("---")
    categoria = st.selectbox("Selecciona una categoría", list(estructura.keys()))
    with st.form(f"form_{categoria}"):
        datos = {}
        for campo, opciones in estructura[categoria].items():
            if opciones:
                datos[campo] = st.selectbox(campo, opciones.split(","), key=f"{categoria}_{campo}")
            else:
                datos[campo] = st.text_input(campo, key=f"{categoria}_{campo}")
        evidencias = st.file_uploader("Sube evidencias (PDF, JPG, XLSX, etc.)", accept_multiple_files=True, key=f"{categoria}_files")
        agregar = st.form_submit_button("Agregar entrada")

        if agregar:
            st.session_state.entradas.setdefault(categoria, []).append({"datos": datos, "evidencias": evidencias})
            st.success("Entrada agregada correctamente.")

# BOTÓN FINAL DE ENVÍO
if "datos_empresa" in st.session_state and st.button("📤 Enviar todo"):
    nombre_archivo = f"SUMAC_{st.session_state.datos_empresa['Nombre'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    excel_path = f"datos/{nombre_archivo}.xlsx"
    zip_path = f"datos/{nombre_archivo}.zip"
    writer = pd.ExcelWriter(excel_path, engine="openpyxl")
    adjuntos = [excel_path]

    for hoja, entradas in st.session_state.entradas.items():
        df = pd.DataFrame([e["datos"] for e in entradas])
        df.to_excel(writer, sheet_name=hoja[:31], index=False)
        carpeta = f"evidencias/{hoja}"
        os.makedirs(carpeta, exist_ok=True)
        for i, entrada in enumerate(entradas):
            for archivo in entrada["evidencias"]:
                ruta = os.path.join(carpeta, f"{i+1}_{archivo.name}")
                with open(ruta, "wb") as f:
                    f.write(archivo.read())
    writer.close()

    with ZipFile(zip_path, "w") as z:
        z.write(excel_path, arcname=os.path.basename(excel_path))
        for root, _, files in os.walk("evidencias"):
            for file in files:
                path = os.path.join(root, file)
                z.write(path, arcname=os.path.relpath(path, "evidencias"))
    adjuntos.append(zip_path)

    enviar_correo(
        st.session_state.datos_empresa["Email"],
        "Formulario CHC recibido",
        f"Formulario enviado por: {st.session_state.datos_empresa['Responsable']}",
        adjuntos
    )
    st.success("Formulario y evidencias enviados correctamente.")
