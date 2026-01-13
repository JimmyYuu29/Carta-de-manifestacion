"""
Main Streamlit Application - Carta de Manifestacion Generator
Aplicacion principal de Streamlit - Generador de Cartas de Manifestacion
"""

import streamlit as st
from datetime import datetime, date
from pathlib import Path
import sys
import json
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
    set_imported_data,
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


def process_json_file(uploaded_file) -> dict:
    """
    Process uploaded JSON file
    Procesar archivo JSON cargado
    """
    try:
        content = uploaded_file.read().decode('utf-8')
        data = json.loads(content)

        # Normalize boolean values
        for key, value in data.items():
            if isinstance(value, str):
                if value.upper() in ['SI', 'SÍ', 'TRUE', 'YES']:
                    data[key] = True
                elif value.upper() in ['NO', 'FALSE']:
                    data[key] = False

        return data
    except Exception as e:
        st.error(f"Error al procesar el archivo JSON: {str(e)}")
        return {}


def serialize_for_export(data: dict) -> dict:
    """
    Serialize data for JSON/Excel export, converting date objects to strings
    Serializar datos para exportacion JSON/Excel, convirtiendo fechas a strings
    """
    result = {}
    for key, value in data.items():
        if isinstance(value, (date, datetime)):
            result[key] = value.strftime("%d/%m/%Y")
        elif isinstance(value, list):
            # Handle list of dicts (like directors)
            result[key] = value
        else:
            result[key] = value
    return result


def export_to_json(data: dict) -> str:
    """Export data to JSON string"""
    serialized = serialize_for_export(data)
    return json.dumps(serialized, indent=2, ensure_ascii=False)


def export_to_excel(data: dict) -> bytes:
    """Export data to Excel bytes"""
    serialized = serialize_for_export(data)

    # Flatten the data for Excel
    rows = []
    for key, value in serialized.items():
        if isinstance(value, list):
            # For lists like directors, create a JSON string representation
            rows.append({"Variable": key, "Valor": json.dumps(value, ensure_ascii=False)})
        elif isinstance(value, bool):
            rows.append({"Variable": key, "Valor": "SI" if value else "NO"})
        else:
            rows.append({"Variable": key, "Valor": str(value) if value else ""})

    df = pd.DataFrame(rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Metadatos')
    output.seek(0)

    return output.getvalue()


def main():
    """Main application entry point / Punto de entrada principal"""

    # Page configuration / Configuración de página
    st.set_page_config(
        page_title="Generador de Cartas de Manifestación",
        page_icon="📄",
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

    # Main title / Título principal
    st.title("🏢 Generador de Cartas de Manifestación - Forvis Mazars")
    st.markdown("---")

    # Get template path
    template_path = PROJECT_ROOT / "Modelo de plantilla.docx"
    if not template_path.exists():
        # Try config path
        template_path = plugin.get_template_path()

    if not template_path.exists():
        st.error(f"⚠️ No se encontró el archivo de plantilla")
        st.info("Por favor, asegúrate de que el archivo de plantilla esté en la carpeta correcta.")
        return

    # Template analysis message
    st.success(f"✅ Plantilla analizada correctamente.")

    # Form subtitle / Subtítulo del formulario
    st.subheader("📝 Información de la Carta")

    # Import section / Sección de importación
    st.markdown("---")
    st.subheader("📁 Importar Metadatos")

    col_import1, col_import2, col_import3 = st.columns(3)

    with col_import1:
        uploaded_json = st.file_uploader(
            "Cargar archivo JSON (.json)",
            type=['json'],
            help="Archivo JSON con metadatos exportados previamente",
            key="json_upload"
        )

    with col_import2:
        uploaded_excel = st.file_uploader(
            "Cargar archivo Excel (.xlsx, .xls)",
            type=['xlsx', 'xls'],
            help="Formato: Columna 1 = Nombre variable, Columna 2 = Valor",
            key="excel_upload"
        )

    with col_import3:
        uploaded_word = st.file_uploader(
            "Cargar archivo Word (.docx)",
            type=['docx'],
            help="Formato: nombre_variable: valor (una por línea)",
            key="word_upload"
        )

    # Process uploaded files
    imported_data = {}

    if uploaded_json is not None:
        with st.spinner("Procesando archivo JSON..."):
            imported_data = process_json_file(uploaded_json)
            if imported_data:
                set_imported_data(imported_data)
                st.success(f"✅ Se importaron {len(imported_data)} valores desde JSON")

    elif uploaded_excel is not None:
        with st.spinner("Procesando archivo Excel..."):
            imported_data = process_uploaded_file(uploaded_excel, "excel")
            if imported_data:
                set_imported_data(imported_data)
                st.success(f"✅ Se importaron {len(imported_data)} valores desde Excel")

    elif uploaded_word is not None:
        with st.spinner("Procesando archivo Word..."):
            imported_data = process_uploaded_file(uploaded_word, "word")
            if imported_data:
                set_imported_data(imported_data)
                st.success(f"✅ Se importaron {len(imported_data)} valores desde Word")

    st.markdown("---")

    # Form sections in columns / Secciones del formulario en columnas
    col1, col2 = st.columns(2)

    # Get current form data
    var_values = dict(st.session_state.form_data)
    cond_values = {}

    with col1:
        # Office section / Sección de oficina
        st.markdown("### 📋 Información de la Oficina")
        var_values = form_renderer.render_oficina_section(var_values)

        # Client section / Sección de cliente
        st.markdown("### 🏢 Nombre de cliente")
        var_values['Nombre_Cliente'] = st.text_input(
            "Nombre del Cliente",
            value=var_values.get('Nombre_Cliente', ''),
            key="nombre_cliente"
        )

        # Dates section / Sección de fechas
        st.markdown("### 📅 Fechas")

        # Store dates as date objects for validation, formatting happens in context_builder
        fecha_hoy = parse_date_string(var_values.get('Fecha_de_hoy', ''))
        if not fecha_hoy:
            fecha_hoy = datetime.now().date()
        var_values['Fecha_de_hoy'] = st.date_input("Fecha de Hoy", value=fecha_hoy, key="fecha_hoy")

        fecha_encargo = parse_date_string(var_values.get('Fecha_encargo', ''))
        if not fecha_encargo:
            fecha_encargo = datetime.now().date()
        var_values['Fecha_encargo'] = st.date_input("Fecha del Encargo", value=fecha_encargo, key="fecha_encargo")

        fecha_ff = parse_date_string(var_values.get('FF_Ejecicio', ''))
        if not fecha_ff:
            fecha_ff = datetime.now().date()
        var_values['FF_Ejecicio'] = st.date_input("Fecha Fin del Ejercicio", value=fecha_ff, key="ff_ejercicio")

        fecha_cierre = parse_date_string(var_values.get('Fecha_cierre', ''))
        if not fecha_cierre:
            fecha_cierre = datetime.now().date()
        var_values['Fecha_cierre'] = st.date_input("Fecha de Cierre", value=fecha_cierre, key="fecha_cierre")

        # General info section / Sección de información general
        st.markdown("### 📝 Información General")
        var_values['Lista_Abogados'] = st.text_area(
            "Lista de abogados y asesores fiscales",
            value=var_values.get('Lista_Abogados', ''),
            placeholder="Ej: Despacho ABC - Asesoría fiscal\nDespacho XYZ - Asesoría legal",
            key="abogados"
        )
        var_values['anexo_partes'] = st.text_input(
            "Número anexo partes vinculadas",
            value=var_values.get('anexo_partes', '2'),
            key="anexo_partes"
        )
        var_values['anexo_proyecciones'] = st.text_input(
            "Número anexo proyecciones",
            value=var_values.get('anexo_proyecciones', '3'),
            key="anexo_proyecciones"
        )

    with col2:
        # Administration organ section / Sección órgano de administración
        st.markdown("### 👥 Órgano de Administración")
        organo_options = ['consejo', 'administrador_unico', 'administradores']
        organo_labels = {
            'consejo': 'Consejo de Administración',
            'administrador_unico': 'Administrador Único',
            'administradores': 'Administradores'
        }
        organo_default = var_values.get('organo', 'consejo')
        if organo_default not in organo_options:
            organo_default = 'consejo'

        cond_values['organo'] = st.selectbox(
            "Tipo de Órgano de Administración",
            options=organo_options,
            index=organo_options.index(organo_default),
            format_func=lambda x: organo_labels.get(x, x),
            key="organo"
        )

        # Conditional options section / Sección opciones condicionales
        st.markdown("### ✅ Opciones Condicionales")

        cond_values['comision'] = 'sí' if st.checkbox(
            "¿Existe Comisión de Auditoría?",
            value=var_values.get('comision', False) if isinstance(var_values.get('comision'), bool) else var_values.get('comision') == 'sí',
            key="comision"
        ) else 'no'

        cond_values['junta'] = 'sí' if st.checkbox(
            "¿Incluir Junta de Accionistas?",
            value=var_values.get('junta', False) if isinstance(var_values.get('junta'), bool) else var_values.get('junta') == 'sí',
            key="junta"
        ) else 'no'

        cond_values['comite'] = 'sí' if st.checkbox(
            "¿Incluir Comité?",
            value=var_values.get('comite', False) if isinstance(var_values.get('comite'), bool) else var_values.get('comite') == 'sí',
            key="comite"
        ) else 'no'

        cond_values['incorreccion'] = 'sí' if st.checkbox(
            "¿Hay incorrecciones no corregidas?",
            value=var_values.get('incorreccion', False) if isinstance(var_values.get('incorreccion'), bool) else var_values.get('incorreccion') == 'sí',
            key="incorreccion"
        ) else 'no'

        if cond_values['incorreccion'] == 'sí':
            with st.container():
                st.markdown("##### 📌 Detalles de incorrecciones")
                var_values['Anio_incorreccion'] = st.text_input(
                    "Año de la incorrección",
                    value=var_values.get('Anio_incorreccion', ''),
                    key="anio_inc"
                )
                var_values['Epigrafe'] = st.text_input(
                    "Epígrafe afectado",
                    value=var_values.get('Epigrafe', ''),
                    key="epigrafe"
                )
                cond_values['limitacion_alcance'] = 'sí' if st.checkbox(
                    "¿Hay limitación al alcance?",
                    value=var_values.get('limitacion_alcance', False) if isinstance(var_values.get('limitacion_alcance'), bool) else var_values.get('limitacion_alcance') == 'sí',
                    key="limitacion"
                ) else 'no'
                if cond_values['limitacion_alcance'] == 'sí':
                    var_values['detalle_limitacion'] = st.text_area(
                        "Detalle de la limitación",
                        value=var_values.get('detalle_limitacion', ''),
                        key="det_limitacion"
                    )

        cond_values['dudas'] = 'sí' if st.checkbox(
            "¿Existen dudas sobre empresa en funcionamiento?",
            value=var_values.get('dudas', False) if isinstance(var_values.get('dudas'), bool) else var_values.get('dudas') == 'sí',
            key="dudas"
        ) else 'no'

        cond_values['rent'] = 'sí' if st.checkbox(
            "¿Incluir párrafo sobre arrendamientos?",
            value=var_values.get('rent', False) if isinstance(var_values.get('rent'), bool) else var_values.get('rent') == 'sí',
            key="rent"
        ) else 'no'

        cond_values['A_coste'] = 'sí' if st.checkbox(
            "¿Hay activos valorados a coste en vez de valor razonable?",
            value=var_values.get('A_coste', False) if isinstance(var_values.get('A_coste'), bool) else var_values.get('A_coste') == 'sí',
            key="a_coste"
        ) else 'no'

        cond_values['experto'] = 'sí' if st.checkbox(
            "¿Se utilizó un experto independiente?",
            value=var_values.get('experto', False) if isinstance(var_values.get('experto'), bool) else var_values.get('experto') == 'sí',
            key="experto"
        ) else 'no'

        if cond_values['experto'] == 'sí':
            with st.container():
                st.markdown("##### 📌 Información del experto")
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

        cond_values['unidad_decision'] = 'sí' if st.checkbox(
            "¿Bajo la misma unidad de decisión?",
            value=var_values.get('unidad_decision', False) if isinstance(var_values.get('unidad_decision'), bool) else var_values.get('unidad_decision') == 'sí',
            key="unidad_decision"
        ) else 'no'

        if cond_values['unidad_decision'] == 'sí':
            with st.container():
                st.markdown("##### 📌 Información de la unidad de decisión")
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
                    "Localización o domiciliación mercantil",
                    value=var_values.get('localizacion_mer', ''),
                    key="localizacion_mer"
                )

        cond_values['activo_impuesto'] = 'sí' if st.checkbox(
            "¿Hay activos por impuestos diferidos?",
            value=var_values.get('activo_impuesto', False) if isinstance(var_values.get('activo_impuesto'), bool) else var_values.get('activo_impuesto') == 'sí',
            key="activo_impuesto"
        ) else 'no'

        if cond_values['activo_impuesto'] == 'sí':
            with st.container():
                st.markdown("##### 📌 Recuperación de activos")
                var_values['ejercicio_recuperacion_inicio'] = st.text_input(
                    "Ejercicio inicio recuperación",
                    value=var_values.get('ejercicio_recuperacion_inicio', ''),
                    key="rec_inicio"
                )
                var_values['ejercicio_recuperacion_fin'] = st.text_input(
                    "Ejercicio fin recuperación",
                    value=var_values.get('ejercicio_recuperacion_fin', ''),
                    key="rec_fin"
                )

        cond_values['operacion_fiscal'] = 'sí' if st.checkbox(
            "¿Operaciones en paraísos fiscales?",
            value=var_values.get('operacion_fiscal', False) if isinstance(var_values.get('operacion_fiscal'), bool) else var_values.get('operacion_fiscal') == 'sí',
            key="operacion_fiscal"
        ) else 'no'

        if cond_values['operacion_fiscal'] == 'sí':
            with st.container():
                st.markdown("##### 📌 Detalle operaciones")
                var_values['detalle_operacion_fiscal'] = st.text_area(
                    "Detalle operaciones paraísos fiscales",
                    value=var_values.get('detalle_operacion_fiscal', ''),
                    key="det_fiscal"
                )

        cond_values['compromiso'] = 'sí' if st.checkbox(
            "¿Compromisos por pensiones?",
            value=var_values.get('compromiso', False) if isinstance(var_values.get('compromiso'), bool) else var_values.get('compromiso') == 'sí',
            key="compromiso"
        ) else 'no'

        cond_values['gestion'] = 'sí' if st.checkbox(
            "¿Incluir informe de gestión?",
            value=var_values.get('gestion', False) if isinstance(var_values.get('gestion'), bool) else var_values.get('gestion') == 'sí',
            key="gestion"
        ) else 'no'

    # Directors section / Sección alta dirección
    st.markdown("---")
    st.markdown("### 👔 Alta Dirección")

    st.info("Introduce los nombres y cargos de los altos directivos. Estos reemplazarán completamente el ejemplo en la plantilla.")

    num_directivos = st.number_input(
        "Número de altos directivos",
        min_value=0,
        max_value=10,
        value=2,
        key="num_directivos"
    )

    # Store directors as list of dicts for validation, formatting happens in context_builder
    directivos_list = []
    directivos_display = []
    indent = "                                  "

    for i in range(num_directivos):
        col_nombre, col_cargo = st.columns(2)
        with col_nombre:
            nombre = st.text_input(f"Nombre completo {i+1}", key=f"dir_nombre_{i}")
        with col_cargo:
            cargo = st.text_input(f"Cargo {i+1}", key=f"dir_cargo_{i}")
        if nombre and cargo:
            directivos_list.append({"nombre": nombre, "cargo": cargo})
            directivos_display.append(f"{indent} D. {nombre} - {cargo}")

    var_values['lista_alto_directores'] = directivos_list

    # Signature section / Sección persona de firma
    st.markdown("---")
    st.markdown("### 👥 Persona de firma")

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

    # Preview directors list
    if directivos_display:
        st.markdown("#### Vista previa de la lista de directivos:")
        st.code("\n".join(directivos_display))

    # Update session state
    st.session_state.form_data = {**var_values, **cond_values}

    # Automatic review section / Sección de revisión automática
    st.markdown("---")
    st.header("🔍 Revisión automática")

    # Required fields validation
    required_fields = ['Nombre_Cliente', 'Direccion_Oficina', 'CP', 'Ciudad_Oficina']
    missing_fields = [f for f in required_fields if not var_values.get(f)]

    # Show import summary if data was imported
    if imported_data:
        st.info(f"📊 Datos importados: {len(imported_data)} valores")

    # Inform user about validation status
    if not missing_fields:
        st.success("✅ Todas las variables y condiciones están completas.")
    else:
        st.warning(f"⚠️ Faltan {len(missing_fields)} campos obligatorios: {', '.join(missing_fields)}")

    # Export metadata section / Sección exportar metadatos
    st.markdown("---")
    st.subheader("💾 Exportar Metadatos")
    st.info("Exporta los datos del formulario para usarlos posteriormente o compartirlos.")

    # Combine all current data for export
    all_current_data = {**var_values, **cond_values}

    col_export1, col_export2 = st.columns(2)

    with col_export1:
        # Export to JSON
        json_data = export_to_json(all_current_data)
        client_name_safe = var_values.get('Nombre_Cliente', 'documento').replace(' ', '_').replace('/', '_')
        json_filename = f"metadatos_{client_name_safe}_{datetime.now().strftime('%Y%m%d')}.json"

        st.download_button(
            label="📄 Exportar a JSON",
            data=json_data,
            file_name=json_filename,
            mime="application/json",
            help="Descarga los metadatos en formato JSON para importarlos posteriormente"
        )

    with col_export2:
        # Export to Excel
        excel_data = export_to_excel(all_current_data)
        excel_filename = f"metadatos_{client_name_safe}_{datetime.now().strftime('%Y%m%d')}.xlsx"

        st.download_button(
            label="📊 Exportar a Excel",
            data=excel_data,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga los metadatos en formato Excel"
        )

    # Generate button / Botón de generación
    st.markdown("---")

    if st.button("🚀 Generar Carta de Manifestación", type="primary"):
        if missing_fields:
            st.error(f"⚠️ Por favor completa los siguientes campos obligatorios: {', '.join(missing_fields)}")
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
                        st.success("✅ Carta generada exitosamente!")

                        # Display trace code
                        st.markdown("### 🔖 Código de Traza")
                        st.code(result.trace_id, language=None)
                        st.caption("Este código identifica de forma única este documento generado. Guárdelo para referencia futura.")

                        # Display generation info
                        st.info(f"⏱️ Tiempo de generación: {result.duration_ms}ms")

                        # Read generated file
                        with open(result.output_path, 'rb') as f:
                            doc_bytes = f.read()

                        filename = f"Carta_Manifestacion_{var_values['Nombre_Cliente'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}_{result.trace_id[:8]}.docx"

                        st.download_button(
                            label="📥 Descargar Carta de Manifestación",
                            data=doc_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error(f"❌ Error al generar la carta: {result.error}")
                        if result.validation_errors:
                            st.markdown("### Errores de validación:")
                            for err in result.validation_errors:
                                st.warning(err)
                        # Also show trace code for failed generations
                        st.caption(f"Código de traza: {result.trace_id}")

                except Exception as e:
                    st.error(f"❌ Error al generar la carta: {str(e)}")
                    st.exception(e)


if __name__ == "__main__":
    main()
