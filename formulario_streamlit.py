import streamlit as st
import pandas as pd
import os, re, shutil, yagmail
from datetime import datetime
from zipfile import ZipFile

# ------------------ CONFIGURACIÓN ------------------
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
    password = "xbna iizl vhta aync"
    yag = yagmail.SMTP(user=remitente, password=password)
    yag.send(to=destinatario, subject=asunto, contents=cuerpo, attachments=archivos)

# ------------------ APP ------------------
st.title("Formulario de Huella de Carbono - SUMAC")

with st.expander("📘 Instrucciones de uso"):
    st.markdown("""
**Este programa registra información por alcance y fuente de emisión.**  
La estructura se organiza así:

- **A1:** Combustión móvil (vehículos, maquinarias, equipos)
- **A2:** Electricidad adquirida
- **A3:** Consumo indirecto (papelería, agua, transporte, residuos, etc.)

**Pasos para usar:**
1. Ingresa los datos de la empresa.
2. Selecciona una categoría (pestaña) y llena los datos.
3. Adjunta evidencias si aplica.
4. Repite el paso 2 para varias categorías si lo deseas.
5. Presiona **📤 Finalizar y Enviar** para completar.
""")

if "entradas" not in st.session_state:
    st.session_state.entradas = {}

with st.form("form_empresa"):
    st.subheader("Datos de la Empresa")
    col1, col2 = st.columns(2)
    nombre = col1.text_input("Nombre de la empresa")
    ruc = col2.number_input("RUC o ID fiscal", step=1, format="%d")
    pais = col1.selectbox("País", sorted(["Argentina", "Bolivia", "Chile", "Colombia", "Ecuador", "España", "México", "Paraguay", "Perú", "Uruguay", "Estados Unidos"]))
    responsable = col2.text_input("Responsable")
    email = st.text_input("Email del responsable")
    enviado = st.form_submit_button("Iniciar")

if enviado and nombre and responsable and es_email_valido(email):
    st.success("Datos validados. Continúa llenando la información.")
    nombre_archivo = f"SUMAC_{nombre.strip().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    hoja_sel = st.selectbox("Selecciona categoría a llenar", list(estructura.keys()))
    st.markdown(f"### Categoría: {hoja_sel}")

    with st.form(f"form_{hoja_sel}"):
        datos = {}
        for campo, opciones in estructura[hoja_sel].items():
            if opciones:
                datos[campo] = st.selectbox(campo, opciones.split(","), key=f"{hoja_sel}_{campo}")
            else:
                datos[campo] = st.text_input(campo, key=f"{hoja_sel}_{campo}")
        evidencias = st.file_uploader("Subir evidencias", accept_multiple_files=True, key=f"{hoja_sel}_files")
        enviar_fila = st.form_submit_button("Agregar entrada")
        if enviar_fila:
            st.session_state.entradas.setdefault(hoja_sel, []).append({"datos": datos, "evidencias": evidencias})
            st.success("Entrada agregada.")

    if st.button("📤 Finalizar y Enviar"):
        excel_filename = f"datos/{nombre_archivo}.xlsx"
        zip_filename = f"datos/{nombre_archivo}.zip"
        writer = pd.ExcelWriter(excel_filename, engine="openpyxl")
        adjuntos = [excel_filename]

        for hoja, registros in st.session_state.entradas.items():
            df = pd.DataFrame([r["datos"] for r in registros])
            df.to_excel(writer, sheet_name=hoja[:31], index=False)
            hoja_dir = f"evidencias/{hoja}"
            os.makedirs(hoja_dir, exist_ok=True)
            for i, reg in enumerate(registros):
                for file in reg["evidencias"]:
                    ruta = os.path.join(hoja_dir, f"{i+1}_{file.name}")
                    with open(ruta, "wb") as f:
                        f.write(file.read())
        writer.close()

        with ZipFile(zip_filename, "w") as z:
            z.write(excel_filename, arcname=os.path.basename(excel_filename))
            for root, _, files in os.walk("evidencias"):
                for f in files:
                    path = os.path.join(root, f)
                    z.write(path, arcname=os.path.relpath(path, "evidencias"))
        adjuntos.append(zip_filename)

        enviar_correo("avilchez@sumacinc.com", "Formulario CHC recibido", f"Formulario enviado por: {responsable}", adjuntos)
        st.success("Formulario enviado correctamente.")
elif enviado:
    st.error("Por favor completa todos los campos obligatorios correctamente.")
