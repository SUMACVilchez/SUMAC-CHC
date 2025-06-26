import streamlit as st
import pandas as pd
import os, re, yagmail
from datetime import datetime
from zipfile import ZipFile

# ------------------ CONFIGURACIÓN ------------------
st.set_page_config(page_title="Formulario Huella de Carbono - SUMAC", layout="wide")
os.makedirs("datos", exist_ok=True)
os.makedirs("evidencias", exist_ok=True)

# ------------------ ESTRUCTURA ------------------
estructura = {
    'A1_Vehículos_propios_móviles': {'Tipo de vehículo': 'Auto,Camioneta,Furgón,SUV,Motocicleta,Otro', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído', 'Consumo anual': 'numeric', 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Generador_Electri_móvile': {'Tipo de vehículo': None, 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído,Eléctrico', 'Consumo anual': 'numeric', 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Maquinari_propios_móvile': {'Tipo de vehículo': 'Grúa,Escabadora,Montacarga', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído,Eléctrico', 'Consumo anual': 'numeric', 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Equipos_estacionarios': {'Tipo de equipo': 'Elevador,Prensa,Calderas,Trituradora,Motobombas,Generador,Cocina', 'Tipo de combustible': 'Diesel,Gasohol,GNV,GLP,Hibrído', 'Consumo anual': 'numeric', 'Unidad': 'Litros,Galones', 'Año': '2022,2023,2024'},
    'A1_Aire_acondicionado': {'Equipo': None, 'Tipo de gas': 'R-22,R44,R410,HCFC-22,R-410A,R-134a,R-407C,R-404A,R-32,R-600a', '¿Presentó fugas y/o recargas?': 'Fuga,Recarga', 'Capacidad': 'numeric', 'Unidad': 'kg', 'Cantidad de recargas': 'numeric'},
    'A1_Extintores': {'Equipo': None, '¿Presentó fugas y/o recargas?': 'Fuga,Recarga', 'Capacidad': 'numeric', 'Unidad': 'kg', 'Cantidad de recargas': 'numeric'},
    'A2_Electricidad': {'Consumo anual (KWh)': 'numeric', 'Año': '2022,2023,2024'},
    'A3_Transporte__casa__trabajo': {'Alcance 3: Emisiones indirectas de GEI de productos usados por la organización': None},
    'A3_Papelería': {'Descripción': None, 'Largo (cm)': 'numeric', 'Ancho (cm)': 'numeric', 'Gramaje (gr/m2)': 'numeric', 'Cantidad': 'numeric', 'Unidad': None},
    'A3_Transporte_contratado': {'Tipo de vehículo': 'Bus,Van,Taxi,Otro', 'Tipo de combustible': 'Diesel,Gasolina,GNV,GLP,Otro', 'Consumo de combustible anual (litros)': 'numeric', 'Gastos por el servicio': 'numeric', 'Destino incial': None, 'Destino final': None, 'Kilómetros recorridos': 'numeric', 'Año': '2022,2023,2024'},
    'A3_Consumo_de_agua': {'Uso': None, 'Consumo anual (m3)': 'numeric', 'Año': '2022,2023,2024'},
    'A3_Materiales_Bien_y_Servicio': {'Tipo de bien y servicio': 'Bien adquirido,Bien capital,Servicio adquirido,Servicio capital', 'Descripción': None, 'Tipo de moneda': 'Peso argentino,Boliviano,Real,Peso chileno,Peso colombiano,Dólar,Guaraní,Sol,Peso uruguayo,Bolívar', 'Monto': 'numeric'},
    'A3_Residuos': {'Tipo de residuo sólidos': 'Peligrosos,No peligrosos', 'Clasificación': 'Madera,Papel,Cartón,Comida,Textiles,Jardines,Plástico,Residuos urbanos,Lodos tratados,Lodos no tratados,Lodos industriales', 'Cantidad total': 'numeric', 'Unidad': 'Kg,Tn', 'Año': '2022,2023,2024'},
    'A3_Transporte_Residuos': {'Origen': None, 'Destino': None, 'Número de viajes anuales': 'numeric'},
    'A3_Consumo_de_electricidad_loca': {'Consumo anual (KWh)': 'numeric', 'Año': '2022,2023,2024'}
}

# ------------------ FUNCIONES ------------------
def es_email_valido(email):
    return re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email)

def enviar_correo(destinatario, asunto, cuerpo, archivos):
    remitente = "avilchez@sumacinc.com"
    password = "xbna iizl vhta aync"
    yag = yagmail.SMTP(user=remitente, password=password)
    yag.send(to=destinatario, subject=asunto, contents=cuerpo, attachments=archivos)

# ------------------ INSTRUCCIONES ------------------
with st.expander("📘 Instrucciones de uso"):
    st.markdown("""
**Bienvenido a la Calculadora de Huella de Carbono de SUMAC.**

Este programa permite registrar información por alcance y fuente de emisión:

- **A1**: Combustión móvil (vehículos, maquinarias, generadores, etc.)
- **A2**: Electricidad adquirida
- **A3**: Consumo de agua, papelería, transporte contratado, residuos, etc.

**Pasos para completar el formulario:**
1. Ingresa los datos de tu empresa.
2. Selecciona la categoría de emisión que deseas llenar.
3. Completa los campos requeridos y adjunta evidencias si las tienes.
4. Agrega cada entrada haciendo clic en el botón.
5. Cuando hayas terminado, haz clic en “📤 Finalizar y Enviar” para enviar la información y evidencias a SUMAC.
""")

# ------------------ ESTADO ------------------
if "datos_empresa" not in st.session_state:
    st.session_state.datos_empresa = {}
if "entradas" not in st.session_state:
    st.session_state.entradas = {}
if "categoria_actual" not in st.session_state:
    st.session_state.categoria_actual = list(estructura.keys())[0]

# ------------------ FORM EMPRESA ------------------
with st.form("form_empresa"):
    st.subheader("Datos de la Empresa")
    col1, col2 = st.columns(2)
    nombre = col1.text_input("Nombre de la empresa")
    ruc = col2.text_input("RUC o ID fiscal")
    if ruc and not ruc.isdigit():
        st.warning("El RUC debe ser numérico")
    pais = col1.selectbox("País", sorted(["Argentina", "Bolivia", "Chile", "Colombia", "Ecuador", "España", "México", "Paraguay", "Perú", "Uruguay", "Estados Unidos"]))
    responsable = col2.text_input("Responsable")
    email = st.text_input("Email del responsable")
    enviado = st.form_submit_button("Iniciar")

if enviado and nombre and ruc.isdigit() and responsable and es_email_valido(email):
    st.session_state.datos_empresa = {
        "Empresa": nombre,
        "RUC": ruc,
        "País": pais,
        "Responsable": responsable,
        "Email": email
    }
    st.success("Datos validados. Continúa con el registro.")

# ------------------ CATEGORÍAS ------------------
if st.session_state.datos_empresa:
    st.selectbox("Selecciona categoría para llenar", list(estructura.keys()), key="categoria_actual")

    hoja = st.session_state.categoria_actual
    with st.form(f"form_{hoja}"):
        st.subheader(f"Categoría: {hoja}")
        datos = {}
        for campo, tipo in estructura[hoja].items():
            if tipo == "numeric":
                datos[campo] = st.number_input(campo, step=0.01, format="%.2f", key=f"{hoja}_{campo}")
            elif tipo:
                datos[campo] = st.selectbox(campo, tipo.split(","), key=f"{hoja}_{campo}")
            else:
                datos[campo] = st.text_input(campo, key=f"{hoja}_{campo}")
        evidencias = st.file_uploader("Subir evidencias", accept_multiple_files=True, key=f"{hoja}_files")
        if st.form_submit_button("Agregar entrada"):
            st.session_state.entradas.setdefault(hoja, []).append({"datos": datos, "evidencias": evidencias})
            st.success("Entrada registrada.")

    if st.button("📤 Finalizar y Enviar"):
        nombre_archivo = f"SUMAC_{st.session_state.datos_empresa['Empresa'].strip().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        excel_filename = f"datos/{nombre_archivo}.xlsx"
        zip_filename = f"datos/{nombre_archivo}.zip"
        writer = pd.ExcelWriter(excel_filename, engine="openpyxl")
        adjuntos = [excel_filename]

        for hoja, registros in st.session_state.entradas.items():
            if registros:
                df = pd.DataFrame([r["datos"] for r in registros])
                df.to_excel(writer, sheet_name=hoja[:31], index=False)
                carpeta = f"evidencias/{hoja}"
                os.makedirs(carpeta, exist_ok=True)
                for i, reg in enumerate(registros):
                    for file in reg["evidencias"] or []:
                        with open(os.path.join(carpeta, f"{i+1}_{file.name}"), "wb") as f:
                            f.write(file.read())
        writer.close()

        with ZipFile(zip_filename, "w") as z:
            z.write(excel_filename, arcname=os.path.basename(excel_filename))
            for root, _, files in os.walk("evidencias"):
                for f in files:
                    path = os.path.join(root, f)
                    z.write(path, arcname=os.path.relpath(path, "evidencias"))
        adjuntos.append(zip_filename)

        enviar_correo("avilchez@sumacinc.com", "Formulario CHC recibido", f"Formulario enviado por: {st.session_state.datos_empresa['Responsable']}", adjuntos)
        st.success("Formulario enviado correctamente.")
