# @copyright: Martin Rayan Rojas Ksiri


# Importamos las librerías necesarias

import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import streamlit as st

def main():
    
    st.markdown("<h1 style='text-align: center; color: #4CAF50;'>📋 Registro de Datos</h1>", unsafe_allow_html=True)
    st.write("### Por favor, completa los siguientes datos:")

    with st.form(key="seleccion_datos"):
            
            global dni, nombre, conductas, articulos, correccion

            dni = st.text_input("🔹 Ingresa tu DNI")
            nombre = st.text_input("🔹 Ingresa tu Nombre")
            conductas = st.selectbox("⚠️ Elige la Conducta", ["Conducta 1", "Conducta 2", "Conducta 3"])
            articulos = st.selectbox("📜 Elige el Artículo", ["Artículo 1", "Artículo 2", "Artículo 3"])
            correccion = st.selectbox("✅ Elige la Corrección", ["Corrección 1", "Corrección 2", "Corrección 3"])
            boton_enviar = st.form_submit_button(label="📩 Enviar")

    if boton_enviar:
        st.success(f"✅ ¡Hola **{nombre}**! Tus datos han sido registrados correctamente.")
        st.balloons()  

if __name__ == '__main__':
    main()

# Cargamos la plantilla de Word

doc = DocxTemplate("Plantilla.docx")


# Obtenemos la fecha actual en formato día/mes/año

fecha = datetime.today().strftime("%d/%m/%Y")


# Leemos los datos del archivo Excel donde están almacenados los detalles de los alumnos

df = pd.read_excel('datos_alumnos.xlsx')


# Buscamos en el archivo Excel el alumno cuyo DNI coincida con el introducido

for indice, fila in df.iterrows():
    if fila["DNI"] == dni:

        # Extraemos los datos relevantes del tutor y del alumno
        nombre_tutor = fila["NOMBRE_TUTOR"]
        apellido1_tutor = fila["APELLIDO1_TUTOR"]
        apellido2_tutor = fila["APELLIDO2_TUTOR"]
        direccion = fila["DIRECCION"]
        localidad = fila["LOCALIDAD"]
        nombre_alumno = fila["NOMBRE_ALUMNO"]
        apellido1_alumno = fila["APELLIDO1_ALUMNO"]
        apellido2_alumno = fila["APELLIDO2_ALUMNO"]


# Creamos un diccionario con todos los datos que vamos a insertar en la plantilla

constantes = {"fecha":fecha, 
              "nombre_tutor":nombre_tutor, 
              "apellido1_tutor":apellido1_tutor, 
              "apellido2_tutor":apellido2_tutor,
              "dirección":direccion, 
              "localidad":localidad, 
              "nombre_alumno":nombre_alumno,
              "apellido1_alumno":apellido1_alumno,
              "apellido2_alumno":apellido2_alumno,
              "correcion":correccion, 
              "conductas":conductas, 
              "artículos":articulos, 
              "nombre":nombre}


# Rellenamos la plantilla de Word con los datos del diccionario
doc.render(constantes)

# Guardamos el documento con los datos ya completados
doc.save("prueba.docx")