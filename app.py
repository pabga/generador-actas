#streamlit run generar_acta.py
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import sys
from num2words import num2words  # Necesario para la conversión
import datetime

# --- 1. CONFIGURACIÓN INICIAL ---
ARCHIVO_EXCEL = "base_datos_cursos.xlsx"
ARCHIVO_PLANTILLA = "plantilla_acta.docx"


# --- 2. FUNCIÓN DE FORMATEO (ACTUALIZADA) ---
def formatear_nota_especial(nota_str):
    """
    Toma un string como "4" o "4,25" y lo convierte a:
    - 4 (Cuatro) si es entero
    - 4,25 (Cuatro/25) si tiene decimales
    """
    if not nota_str or nota_str.strip() == "":
        return "AUSENTE"

    # Unificar la entrada (reemplazar coma por punto)
    nota_limpia = nota_str.strip().replace(",", ".")

    try:
        nota_num = float(nota_limpia)
    except ValueError:
        # Si el usuario escribe "AUSENTE" o algo que no es un número
        return nota_str.upper()

    # --- Aquí ocurre la nueva magia ---

    # 1. Separar parte entera y decimal
    parte_entera = int(nota_num)
    # (Usamos round() para evitar problemas de precisión)
    parte_decimal = int(round((nota_num - parte_entera) * 100))

    # 2. Convertir la parte entera a palabras en español
    try:
        palabra_entera = num2words(parte_entera, lang='es').capitalize()
    except Exception:
        palabra_entera = str(parte_entera)  # Fallback

    # 3. Decidir el formato final
    if parte_decimal == 0:
        # Si es un número entero (ej: 4.0)
        # Devolvemos "4 (Cuatro)"
        return f"{parte_entera} ({palabra_entera})"
    else:
        # Si tiene decimales (ej: 4.25)
        # Formateamos el número a "4,25"
        nota_formateada_coma = f"{nota_num:.2f}".replace(".", ",")
        # Formateamos los decimales a "25"
        decimal_dos_digitos = f"{parte_decimal:02d}"
        # Devolvemos "4,25 (Cuatro/25)"
        return f"{nota_formateada_coma} ({palabra_entera}/{decimal_dos_digitos})"


# --- Cargar plantilla (sin cambios) ---
try:
    doc = DocxTemplate(ARCHIVO_PLANTILLA)
except Exception as e:
    st.error(f"ERROR: No se pudo cargar la plantilla '{ARCHIVO_PLANTILLA}'. {e}")
    st.stop()

# --- Cargar las 3 hojas del Excel (sin cambios) ---
try:
    df_cursos = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Cursos", dtype=str)
    df_alumnos = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Alumnos", dtype=str)
    df_inscripciones = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Inscripciones", dtype=str)
except FileNotFoundError:
    st.error(f"ERROR: No se encontró el archivo Excel '{ARCHIVO_EXCEL}'.")
    st.stop()
except Exception as e:
    st.error(
        f"ERROR: No se pudo leer el Excel. Asegúrate de tener las hojas 'Cursos', 'Alumnos' e 'Inscripciones'. Detalle: {e}")
    st.stop()

# --- 2. INTERFAZ DE STREAMLIT (NUEVA LÓGICA) ---
st.title("🚀 Generador de Actas de Examen")

# --- PASO 1: Preguntar el Tipo de Examen ---
st.sidebar.markdown("## Datos del Acta")
tipo_seleccionado = st.sidebar.radio(
    "1. Seleccione el tipo de acta:",
    ("FINAL", "PARCIAL")
)

# --- PASO 2: Selector de Fecha Manual ---
fecha_examen_seleccionada = st.sidebar.date_input(
    "2. Seleccione la Fecha del Examen",
    datetime.date.today()
)

# --- PASO 3: Selector de Curso ---
lista_nombres_cursos = df_cursos['NombreCurso'].unique()
curso_seleccionado_nombre = st.selectbox(
    "3. Seleccione el Curso:",
    lista_nombres_cursos
)

# --- PASO 4: Selector de Asignatura ---
if curso_seleccionado_nombre:
    materias_del_curso = df_cursos[
        df_cursos['NombreCurso'] == curso_seleccionado_nombre
        ]['Asignatura'].unique()

    asignatura_seleccionada = st.selectbox(
        "4. Seleccione la Asignatura:",
        materias_del_curso
    )

# --- 5. FILTRAR ALUMNOS (Lógica de Grupo) ---
if curso_seleccionado_nombre and asignatura_seleccionada:

    # a. Obtener el ID_CURSO
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

    # b. Encontrar los GRUPOS inscriptos a este ID_CURSO
    grupos_inscriptos_df = df_inscripciones[df_inscripciones['ID_CURSO'] == id_curso_seleccionado]
    lista_grupos = grupos_inscriptos_df['Grupo'].unique()

    if len(lista_grupos) == 0:
        st.warning(
            f"No hay ningún 'Grupo' inscripto a este curso (ID: {id_curso_seleccionado}) en la hoja 'Inscripciones'.")
        st.stop()

    # c. Traer a TODOS los alumnos que pertenecen a esos grupos
    alumnos_del_curso = df_alumnos[df_alumnos['Grupo'].isin(lista_grupos)].copy()
    alumnos_del_curso = alumnos_del_curso.drop_duplicates(subset=['DNI'])

    if alumnos_del_curso.empty:
        st.warning(
            f"Se encontraron grupos ({', '.join(lista_grupos)}) pero no hay alumnos en la hoja 'Alumnos' que pertenezcan a ellos.")
    else:
        # --- Formulario de Notas (sin cambios) ---
        with st.form("notas_form"):
            notas_ingresadas = {}
            st.write("**5. Ingrese la nota numérica (ej: 9,50 o 7):**")

            for index, alumno in alumnos_del_curso.iterrows():
                dni = alumno['DNI']
                nombre = alumno['NombreApellido']
                nota = st.text_input(f"Nota para: **{nombre}** (DNI: {dni})", key=dni)
                notas_ingresadas[dni] = nota

            submitted = st.form_submit_button("Generar Acta")

# --- 6. LÓGICA DE GENERACIÓN (MODIFICADA) ---
if 'submitted' in locals() and submitted:

    # a. Preparar el contexto (datos del curso)
    context = info_curso_dict

    # b. Agregar los datos manuales (Tipo y Fecha)
    context['TipodeExamen'] = tipo_seleccionado  # (Corregido)
    context['FechaExamen'] = fecha_examen_seleccionada.strftime("%d/%m/%Y")

    # c. Preparar la lista de alumnos (con la nueva función)
    lista_alumnos_para_plantilla = []

    for index, alumno in alumnos_del_curso.iterrows():
        alumno_dict = alumno.to_dict()
        nota_ingresada_str = notas_ingresadas.get(alumno['DNI'], "")

        # Aplicamos la nueva función de formato
        nota_transformada = formatear_nota_especial(nota_ingresada_str)

        alumno_dict['resultado'] = nota_transformada
        lista_alumnos_para_plantilla.append(alumno_dict)

    context['alumnos'] = lista_alumnos_para_plantilla

    # --- Renderizar y Descargar (sin cambios) ---
    try:
        doc.render(context)
        file_buffer = io.BytesIO()
        doc.save(file_buffer)
        file_buffer.seek(0)

        nombre_archivo = f"ACTA_{info_curso_dict.get('Asignatura', 'CURSO')}_{id_curso_seleccionado}.docx"
        st.success(f"✅ ¡Acta generada con éxito!")

        st.download_button(
            label=f"Descargar {nombre_archivo}",
            data=file_buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"ERROR: Ocurrió un problema al 'renderizar' la plantilla Word.")
        st.error(f"Detalle: {e}")