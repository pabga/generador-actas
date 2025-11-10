#streamlit run generar_acta.py
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import sys
from num2words import num2words
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# import pydrive2 - Ya no se necesita
# from pydrive2.auth import GoogleAuth - Ya no se necesita
# from pydrive2.drive import GoogleDrive - Ya no se necesita
import json

# --- 1. CONFIGURACIÓN INICIAL ---
NOMBRE_GOOGLE_SHEET = "base_datos_cursos" 
ARCHIVO_PLANTILLA = "plantilla_acta.docx"
# ID_CARPETA_DRIVE_SALIDA = "..." - Ya no se necesita

# --- 2. AUTENTICACIÓN (SOLO PARA LEER SHEETS) ---
@st.cache_resource
def autorizar_google_sheets():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        # "https://www.googleapis.com/auth/drive" - Ya no se necesita
    ]
    creds_json_string = st.secrets["google_credentials"]
    creds_dict = json.loads(creds_json_string) 
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(creds)
    return gc

try:
    gc = autorizar_google_sheets()
except Exception as e:
    st.error(f"Error al conectar con Google APIs. ¿Configuraste los 'Secrets'? Detalle: {e}")
    st.stop()
    
# --- FUNCIÓN DE FORMATEO (sin cambios) ---
def formatear_nota_especial(nota_str):
    if not nota_str or nota_str.strip() == "":
        return "AUSENTE"
    nota_limpia = nota_str.strip().replace(",", ".")
    try:
        nota_num = float(nota_limpia)
    except ValueError:
        return nota_str.upper()

    parte_entera = int(nota_num)
    parte_decimal = int(round((nota_num - parte_entera) * 100))
    try:
        palabra_entera = num2words(parte_entera, lang='es').capitalize()
    except Exception:
        palabra_entera = str(parte_entera)

    if parte_decimal == 0:
        return f"{parte_entera} ({palabra_entera})"
    else:
        nota_formateada_coma = f"{nota_num:.2f}".replace(".", ",")
        decimal_dos_digitos = f"{parte_decimal:02d}"
        return f"{nota_formateada_coma} ({palabra_entera}/{decimal_dos_digitos})"

# --- Cargar plantilla (sin cambios) ---
try:
    doc = DocxTemplate(ARCHIVO_PLANTILLA)
except Exception as e:
    st.error(f"ERROR: No se pudo cargar la plantilla '{ARCHIVO_PLANTILLA}'. {e}")
    st.stop()

# --- 3. LEER DATOS DESDE GOOGLE SHEETS (sin cambios) ---
@st.cache_data(ttl=600) 
def cargar_datos_google_sheets():
    try:
        sh = gc.open(NOMBRE_GOOGLE_SHEET)
        df_cursos = pd.DataFrame(sh.worksheet("Cursos").get_all_records())
        df_alumnos = pd.DataFrame(sh.worksheet("Alumnos").get_all_records())
        df_inscripciones = pd.DataFrame(sh.worksheet("Inscripciones").get_all_records())
        df_cursos = df_cursos.astype(str)
        df_alumnos = df_alumnos.astype(str)
        df_inscripciones = df_inscripciones.astype(str)
        return df_cursos, df_alumnos, df_inscripciones
    except Exception as e:
        st.error(f"ERROR: No se pudo leer el Google Sheet '{NOMBRE_GOOGLE_SHEET}'. ¿Lo compartiste con el robot? ¿Están limpias las columnas? Detalle: {e}")
        st.stop()

df_cursos, df_alumnos, df_inscripciones = cargar_datos_google_sheets()

# --- 4. INTERFAZ DE STREAMLIT (sin cambios) ---
st.title("🚀 Generador de Actas de Examen")

st.sidebar.markdown("## Datos del Acta")
tipo_seleccionado = st.sidebar.radio(
    "1. Seleccione el tipo de acta:",
    ("Final", "Parcial") 
)
fecha_examen_seleccionada = st.sidebar.date_input(
    "2. Seleccione la Fecha del Examen",
    datetime.date.today()
)

lista_nombres_cursos = df_cursos['NombreCurso'].unique() 
curso_seleccionado_nombre = st.selectbox(
    "3. Seleccione el Curso:", 
    lista_nombres_cursos
)

if curso_seleccionado_nombre:
    materias_del_curso = df_cursos[
        df_cursos['NombreCurso'] == curso_seleccionado_nombre
    ]['Asignatura'].unique()
    
    asignatura_seleccionada = st.selectbox(
        "4. Seleccione la Asignatura:",
        materias_del_curso
    )

# --- 5. FILTRAR ALUMNOS (sin cambios) ---
if curso_seleccionado_nombre and asignatura_seleccionada:
    
    try:
        curso_final_serie = df_cursos[
            (df_cursos['NombreCurso'] == curso_seleccionado_nombre) &
            (df_cursos['Asignatura'] == asignatura_seleccionada)
        ].iloc[0] 
    except IndexError:
        st.error("Error: No se encontró esa combinación de Curso y Asignatura.")
        st.stop()
    
    info_curso_dict = curso_final_serie.to_dict()
    id_curso_seleccionado = info_curso_dict['ID_CURSO'] 

    st.subheader(f"Cargar notas para: {asignatura_seleccionada} ({tipo_seleccionado})")
    st.caption(f"Curso: {curso_seleccionado_nombre} | ID: {id_curso_seleccionado}")
    
    grupos_inscriptos_df = df_inscripciones[df_inscripciones['ID_CURSO'] == id_curso_seleccionado]
    lista_grupos = grupos_inscriptos_df['Grupo'].unique()
    
    if len(lista_grupos) == 0:
        st.warning(f"No hay ningún 'Grupo' inscripto a este curso (ID: {id_curso_seleccionado}) en la hoja 'Inscripciones'.")
        st.stop()

    alumnos_del_curso = df_alumnos[df_alumnos['Grupo'].isin(lista_grupos)].copy()
    alumnos_del_curso = alumnos_del_curso.drop_duplicates(subset=['DNI'])

    if alumnos_del_curso.empty:
        st.warning(f"Se encontraron grupos ({', '.join(lista_grupos)}) pero no hay alumnos en la hoja 'Alumnos' que pertenezcan a ellos.")
    else:
        with st.form("notas_form"):
            notas_ingresadas = {}
            st.write("**5. Ingrese la nota numérica (ej: 9,50 o 7):**")
            
            for index, alumno in alumnos_del_curso.iterrows():
                dni = alumno['DNI']
                nombre = alumno['NombreApellido']
                nota = st.text_input(f"Nota para: **{nombre}** (DNI: {dni})", key=dni)
                notas_ingresadas[dni] = nota

            # Botón modificado
            submitted = st.form_submit_button("Generar Acta para Descargar")

# --- 6. LÓGICA DE GENERACIÓN (MODIFICADA) ---
if 'submitted' in locals() and submitted:
    
    context = info_curso_dict
    context['TipodeExamen'] = tipo_seleccionado
    context['FechaExamen'] = fecha_examen_seleccionada.strftime("%d/%m/%Y")

    lista_alumnos_para_plantilla = []
    for index, alumno in alumnos_del_curso.iterrows():
        alumno_dict = alumno.to_dict()
        nota_ingresada_str = notas_ingresadas.get(alumno['DNI'], "") 
        nota_transformada = formatear_nota_especial(nota_ingresada_str) 
        alumno_dict['resultado'] = nota_transformada
        lista_alumnos_para_plantilla.append(alumno_dict)
        
    context['alumnos'] = lista_alumnos_para_plantilla

    try:
        # Renderizar el documento en memoria
        doc.render(context)
        file_buffer = io.BytesIO()
        doc.save(file_buffer)
        file_buffer.seek(0) # Rebobinar el buffer al inicio
        
        nombre_archivo = f"ACTA_{info_curso_dict.get('Asignatura', 'CURSO')}_{id_curso_seleccionado}.docx"
        
        # --- 1. Subir a Drive (ELIMINADO) ---
        # (Se eliminó todo el bloque 'with st.spinner(...)')

        # --- 2. Ofrecer la descarga local ---
        st.success(f"✅ ¡Acta generada! Haz clic para descargar.")
        
        st.download_button(
            label=f"Descargar '{nombre_archivo}' a tu PC",
            data=file_buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.balloons() # ¡Celebración!

    except Exception as e:
        st.error(f"ERROR: Ocurrió un problema al 'renderizar' el archivo.")
        st.error(f"Detalle: {e}")