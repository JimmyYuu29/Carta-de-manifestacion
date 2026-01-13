"""
Main Streamlit Application - Carta de Manifestacion Generator
Aplicacion principal de Streamlit - Generador de Cartas de Manifestacion
"""

import streamlit as st
from datetime import datetime
from pathlib import Path
import sys
import io
import pandas as pd
from docx import Document

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from modules.plugin_loader import load_plugin
from modules.generate import generate_from_form
from modules.context_builder import format_spanish_date, parse_date_string

from ui.streamlit_app.state_store import (
    init_session_state,
    get_field_value,
    set_field_value,
    get_list_items,
    add_list_item,
    remove_list_item,
    get_all_form_data,
    clear_form_data,
    set_imported_data,
    get_stable_key,
)
from ui.streamlit_app.components import (
    render_header,
    render_section_header,
    render_success_message,
    render_error_message,
    render_divider,
)
from ui.streamlit_app.form_renderer import FormRenderer


# Plugin configuration
PLUGIN_ID = "carta_manifestacion"


def process_uploaded_file(uploaded_file, file_type: str) -> dict:
    """
    Process uploaded Excel or Word file
    Procesar archivo Excel o Word cargado
    """
    extracted_data = {}

    try:
        if file_type == "excel":
            df = pd.read_excel(uploaded_file, header=None)

            if df.shape[1] >= 2:
                for index, row in df.iterrows():
                    if pd.notna(row[0]) and pd.notna(row[1]):
                        var_name = str(row[0]).strip()
                        var_value = row[1]

                        if pd.api.types.is_datetime64_any_dtype(type(var_value)) or isinstance(var_value, datetime):
                            var_value = var_value.strftime("%d/%m/%Y")
                        else:
                            var_value = str(var_value).strip()

                        # Normalize boolean values
                        if var_value.upper() in ['SI', 'SÍ'] or var_value == '1':
                            var_value = True
                        elif var_value.upper() == 'NO' or var_value == '0':
                            var_value = False

                        extracted_data[var_name] = var_value

        elif file_type == "word":
            doc = Document(uploaded_file)

            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text and ':' in text:
                    parts = text.split(':', 1)
                    if len(parts) == 2:
                        var_name = parts[0].strip()
                        var_value = parts[1].strip()

                        if var_value.upper() in ['SI', 'SÍ'] or var_value == '1':
                            var_value = True
                        elif var_value.upper() == 'NO' or var_value == '0':
                            var_value = False

                        extracted_data[var_name] = var_value

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return {}

    return extracted_data


def main():
    """Main application entry point / Punto de entrada principal"""

    # Page configuration
    st.set_page_config(
        page_title="Generador de Cartas de Manifestacion",
        page_icon="page_facing_up",
        layout="wide"
    )

    # Initialize session state
    init_session_state(PLUGIN_ID)

    # Load plugin
    try:
        plugin = load_plugin(PLUGIN_ID)
    except Exception as e:
        st.error(f"Error loading plugin: {e}")
        return

    # Create form renderer
    form_renderer = FormRenderer(plugin)

    # Header
    render_header(
        "Generador de Cartas de Manifestacion",
        "Forvis Mazars"
    )

    # Get template path
    template_path = PROJECT_ROOT / "Modelo de plantilla.docx"
    if not template_path.exists():
        # Try config path
        template_path = plugin.get_template_path()

    if not template_path.exists():
        st.error(f"No se encontro el archivo de plantilla")
        st.info("Por favor, asegurate de que el archivo de plantilla este en la carpeta correcta.")
        return

    st.success("Plantilla cargada correctamente")

    # Import section
    render_divider()
    render_section_header("Importar datos desde archivo", "file_folder")

    col_import1, col_import2 = st.columns(2)

    with col_import1:
        uploaded_excel = st.file_uploader(
            "Cargar archivo Excel (.xlsx, .xls)",
            type=['xlsx', 'xls'],
            help="Formato: Columna 1 = Nombre variable, Columna 2 = Valor",
            key="excel_upload"
        )

    with col_import2:
        uploaded_word = st.file_uploader(
            "Cargar archivo Word (.docx)",
            type=['docx'],
            help="Formato: nombre_variable: valor (una por linea)",
            key="word_upload"
        )

    # Process uploaded files
    if uploaded_excel is not None:
        with st.spinner("Procesando archivo Excel..."):
            imported_data = process_uploaded_file(uploaded_excel, "excel")
            if imported_data:
                set_imported_data(imported_data)
                st.success(f"Se importaron {len(imported_data)} valores desde Excel")

    elif uploaded_word is not None:
        with st.spinner("Procesando archivo Word..."):
            imported_data = process_uploaded_file(uploaded_word, "word")
            if imported_data:
                set_imported_data(imported_data)
                st.success(f"Se importaron {len(imported_data)} valores desde Word")

    render_divider()

    # Form sections in columns
    col1, col2 = st.columns(2)

    # Get current form data
    var_values = dict(st.session_state.form_data)
    cond_values = {}

    with col1:
        # Office section
        render_section_header("Informacion de la Oficina", "clipboard")
        var_values = form_renderer.render_oficina_section(var_values)

        # Client section
        render_section_header("Nombre de cliente", "building")
        var_values['Nombre_Cliente'] = st.text_input(
            "Nombre del Cliente *",
            value=var_values.get('Nombre_Cliente', ''),
            key="nombre_cliente"
        )

        # Dates section
        render_section_header("Fechas", "calendar")

        fecha_hoy = parse_date_string(var_values.get('Fecha_de_hoy', ''))
        if not fecha_hoy:
            fecha_hoy = datetime.now().date()
        var_values['Fecha_de_hoy'] = format_spanish_date(
            st.date_input("Fecha de Hoy", value=fecha_hoy, key="fecha_hoy")
        )

        fecha_encargo = parse_date_string(var_values.get('Fecha_encargo', ''))
        if not fecha_encargo:
            fecha_encargo = datetime.now().date()
        var_values['Fecha_encargo'] = format_spanish_date(
            st.date_input("Fecha del Encargo", value=fecha_encargo, key="fecha_encargo")
        )

        fecha_ff = parse_date_string(var_values.get('FF_Ejecicio', ''))
        if not fecha_ff:
            fecha_ff = datetime.now().date()
        var_values['FF_Ejecicio'] = format_spanish_date(
            st.date_input("Fecha Fin del Ejercicio", value=fecha_ff, key="ff_ejercicio")
        )

        fecha_cierre = parse_date_string(var_values.get('Fecha_cierre', ''))
        if not fecha_cierre:
            fecha_cierre = datetime.now().date()
        var_values['Fecha_cierre'] = format_spanish_date(
            st.date_input("Fecha de Cierre", value=fecha_cierre, key="fecha_cierre")
        )

        # General info section
        render_section_header("Informacion General", "memo")
        var_values['Lista_Abogados'] = st.text_area(
            "Lista de abogados y asesores fiscales",
            value=var_values.get('Lista_Abogados', ''),
            placeholder="Ej: Despacho ABC - Asesoria fiscal\nDespacho XYZ - Asesoria legal",
            key="abogados"
        )
        var_values['anexo_partes'] = st.text_input(
            "Numero anexo partes vinculadas",
            value=var_values.get('anexo_partes', '2'),
            key="anexo_partes"
        )
        var_values['anexo_proyecciones'] = st.text_input(
            "Numero anexo proyecciones",
            value=var_values.get('anexo_proyecciones', '3'),
            key="anexo_proyecciones"
        )

    with col2:
        # Administration organ section
        render_section_header("Organo de Administracion", "busts_in_silhouette")
        organo_options = ['consejo', 'administrador_unico', 'administradores']
        organo_labels = {
            'consejo': 'Consejo de Administracion',
            'administrador_unico': 'Administrador Unico',
            'administradores': 'Administradores'
        }
        organo_default = var_values.get('organo', 'consejo')
        if organo_default not in organo_options:
            organo_default = 'consejo'

        cond_values['organo'] = st.selectbox(
            "Tipo de Organo de Administracion",
            options=organo_options,
            index=organo_options.index(organo_default),
            format_func=lambda x: organo_labels.get(x, x),
            key="organo"
        )

        # Conditional options section
        render_section_header("Opciones Condicionales", "ballot_box_with_check")

        cond_values['comision'] = st.checkbox(
            "Existe Comision de Auditoria?",
            value=var_values.get('comision', False),
            key="comision"
        )

        cond_values['junta'] = st.checkbox(
            "Incluir Junta de Accionistas?",
            value=var_values.get('junta', False),
            key="junta"
        )

        cond_values['comite'] = st.checkbox(
            "Incluir Comite?",
            value=var_values.get('comite', False),
            key="comite"
        )

        cond_values['incorreccion'] = st.checkbox(
            "Hay incorrecciones no corregidas?",
            value=var_values.get('incorreccion', False),
            key="incorreccion"
        )

        if cond_values['incorreccion']:
            with st.container():
                st.markdown("##### Detalles de incorrecciones")
                var_values['Anio_incorreccion'] = st.text_input(
                    "Anio de la incorreccion",
                    value=var_values.get('Anio_incorreccion', ''),
                    key="anio_inc"
                )
                var_values['Epigrafe'] = st.text_input(
                    "Epigrafe afectado",
                    value=var_values.get('Epigrafe', ''),
                    key="epigrafe"
                )
                cond_values['limitacion_alcance'] = st.checkbox(
                    "Hay limitacion al alcance?",
                    value=var_values.get('limitacion_alcance', False),
                    key="limitacion"
                )
                if cond_values['limitacion_alcance']:
                    var_values['detalle_limitacion'] = st.text_area(
                        "Detalle de la limitacion",
                        value=var_values.get('detalle_limitacion', ''),
                        key="det_limitacion"
                    )

        cond_values['dudas'] = st.checkbox(
            "Existen dudas sobre empresa en funcionamiento?",
            value=var_values.get('dudas', False),
            key="dudas"
        )

        cond_values['rent'] = st.checkbox(
            "Incluir parrafo sobre arrendamientos?",
            value=var_values.get('rent', False),
            key="rent"
        )

        cond_values['A_coste'] = st.checkbox(
            "Hay activos valorados a coste en vez de valor razonable?",
            value=var_values.get('A_coste', False),
            key="a_coste"
        )

        cond_values['experto'] = st.checkbox(
            "Se utilizo un experto independiente?",
            value=var_values.get('experto', False),
            key="experto"
        )

        if cond_values['experto']:
            with st.container():
                st.markdown("##### Informacion del experto")
                var_values['nombre_experto'] = st.text_input(
                    "Nombre del experto",
                    value=var_values.get('nombre_experto', ''),
                    key="experto_nombre"
                )
                var_values['experto_valoracion'] = st.text_input(
                    "Elemento valorado por experto",
                    value=var_values.get('experto_valoracion', ''),
                    key="experto_val"
                )

        cond_values['unidad_decision'] = st.checkbox(
            "Bajo la misma unidad de decision?",
            value=var_values.get('unidad_decision', False),
            key="unidad_decision"
        )

        if cond_values['unidad_decision']:
            with st.container():
                st.markdown("##### Informacion de la unidad de decision")
                var_values['nombre_unidad'] = st.text_input(
                    "Nombre de la unidad",
                    value=var_values.get('nombre_unidad', ''),
                    key="nombre_unidad"
                )
                var_values['nombre_mayor_sociedad'] = st.text_input(
                    "Nombre de la mayor sociedad",
                    value=var_values.get('nombre_mayor_sociedad', ''),
                    key="nombre_mayor_sociedad"
                )
                var_values['localizacion_mer'] = st.text_input(
                    "Localizacion o domiciliacion mercantil",
                    value=var_values.get('localizacion_mer', ''),
                    key="localizacion_mer"
                )

        cond_values['activo_impuesto'] = st.checkbox(
            "Hay activos por impuestos diferidos?",
            value=var_values.get('activo_impuesto', False),
            key="activo_impuesto"
        )

        if cond_values['activo_impuesto']:
            with st.container():
                st.markdown("##### Recuperacion de activos")
                var_values['ejercicio_recuperacion_inicio'] = st.text_input(
                    "Ejercicio inicio recuperacion",
                    value=var_values.get('ejercicio_recuperacion_inicio', ''),
                    key="rec_inicio"
                )
                var_values['ejercicio_recuperacion_fin'] = st.text_input(
                    "Ejercicio fin recuperacion",
                    value=var_values.get('ejercicio_recuperacion_fin', ''),
                    key="rec_fin"
                )

        cond_values['operacion_fiscal'] = st.checkbox(
            "Operaciones en paraisos fiscales?",
            value=var_values.get('operacion_fiscal', False),
            key="operacion_fiscal"
        )

        if cond_values['operacion_fiscal']:
            with st.container():
                st.markdown("##### Detalle operaciones")
                var_values['detalle_operacion_fiscal'] = st.text_area(
                    "Detalle operaciones paraisos fiscales",
                    value=var_values.get('detalle_operacion_fiscal', ''),
                    key="det_fiscal"
                )

        cond_values['compromiso'] = st.checkbox(
            "Compromisos por pensiones?",
            value=var_values.get('compromiso', False),
            key="compromiso"
        )

        cond_values['gestion'] = st.checkbox(
            "Incluir informe de gestion?",
            value=var_values.get('gestion', False),
            key="gestion"
        )

    # Directors section
    render_divider()
    render_section_header("Alta Direccion", "necktie")

    st.info("Introduce los nombres y cargos de los altos directivos.")

    num_directivos = st.number_input(
        "Numero de altos directivos",
        min_value=0,
        max_value=10,
        value=2,
        key="num_directivos"
    )

    directivos_list = []
    indent = "                                  "

    for i in range(num_directivos):
        col_nombre, col_cargo = st.columns(2)
        with col_nombre:
            nombre = st.text_input(f"Nombre completo {i+1}", key=f"dir_nombre_{i}")
        with col_cargo:
            cargo = st.text_input(f"Cargo {i+1}", key=f"dir_cargo_{i}")
        if nombre and cargo:
            directivos_list.append(f"{indent} D. {nombre} - {cargo}")

    var_values['lista_alto_directores'] = "\n".join(directivos_list)

    if directivos_list:
        st.markdown("#### Vista previa de la lista de directivos:")
        st.code("\n".join(directivos_list))

    # Signature section
    render_divider()
    render_section_header("Persona de firma", "busts_in_silhouette")

    var_values['Nombre_Firma'] = st.text_input(
        "Nombre del firmante",
        value=var_values.get('Nombre_Firma', ''),
        key="nombre_firma"
    )
    var_values['Cargo_Firma'] = st.text_input(
        "Cargo del firmante",
        value=var_values.get('Cargo_Firma', ''),
        key="cargo_firma"
    )

    # Update session state
    st.session_state.form_data = {**var_values, **cond_values}

    # Generate button
    render_divider()

    if st.button("Generar Carta de Manifestacion", type="primary"):
        # Validate required fields
        required_fields = ['Nombre_Cliente', 'Direccion_Oficina', 'CP', 'Ciudad_Oficina']
        missing_fields = [f for f in required_fields if not var_values.get(f)]

        if missing_fields:
            st.error(f"Por favor completa los siguientes campos obligatorios: {', '.join(missing_fields)}")
        else:
            with st.spinner("Generando carta..."):
                try:
                    # Combine all data
                    all_data = {**var_values, **cond_values}

                    # Generate document
                    result = generate_from_form(
                        plugin_id=PLUGIN_ID,
                        form_data=all_data,
                        list_data={},
                        output_dir=PROJECT_ROOT / "output",
                        template_path=template_path
                    )

                    if result.success and result.output_path:
                        st.success("Carta generada exitosamente!")

                        # Read generated file
                        with open(result.output_path, 'rb') as f:
                            doc_bytes = f.read()

                        filename = f"Carta_Manifestacion_{var_values['Nombre_Cliente'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx"

                        st.download_button(
                            label="Descargar Carta de Manifestacion",
                            data=doc_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"Error al generar la carta: {result.error}")
                        if result.validation_errors:
                            for err in result.validation_errors:
                                st.warning(err)

                except Exception as e:
                    st.error(f"Error al generar la carta: {str(e)}")
                    st.exception(e)


if __name__ == "__main__":
    main()
