"""
Form Renderer - Dynamic form rendering from YAML definitions
Renderizador de formularios dinamico desde definiciones YAML
"""

import streamlit as st
from typing import Any, Dict, List, Optional, Callable
from datetime import datetime, date
import sys
from pathlib import Path

# Add modules to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from modules.plugin_loader import PluginPack
from modules.dsl_evaluator import evaluate_condition
from modules.context_builder import format_spanish_date, parse_date_string

from .state_store import (
    get_stable_key,
    get_field_value,
    set_field_value,
    get_list_items,
    add_list_item,
    remove_list_item,
)


class FormRenderer:
    """
    Renders forms dynamically from plugin field definitions
    Renderiza formularios dinamicamente desde definiciones de campos
    """

    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self.fields = plugin.fields.get("fields", {})
        self.oficinas = plugin.get_oficinas()

    def render_form(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render the complete form and return updated data
        Renderizar el formulario completo y devolver datos actualizados

        Args:
            data: Current form data

        Returns:
            Updated form data dictionary
        """
        result = dict(data)
        sections = self.plugin.get_sections()

        # Render by sections
        for section in sections:
            section_id = section.get("id")
            section_title = section.get("title", section_id)
            section_icon = section.get("icon", "")

            # Get fields for this section
            section_fields = self._get_fields_for_section(section_id)

            if section_fields:
                st.markdown(f"### {section_title}")

                for field_name in section_fields:
                    field_spec = self.fields.get(field_name, {})

                    # Check visibility condition
                    if not self._should_show_field(field_spec, result):
                        continue

                    # Render field
                    value = self._render_field(field_name, field_spec, result)
                    result[field_name] = value

        return result

    def render_section(self, section_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render a specific section
        Renderizar una seccion especifica

        Args:
            section_id: Section ID
            data: Current form data

        Returns:
            Updated form data for this section
        """
        result = dict(data)
        section_fields = self._get_fields_for_section(section_id)

        for field_name in section_fields:
            field_spec = self.fields.get(field_name, {})

            if not self._should_show_field(field_spec, result):
                continue

            value = self._render_field(field_name, field_spec, result)
            result[field_name] = value

        return result

    def _get_fields_for_section(self, section_id: str) -> List[str]:
        """Get fields belonging to a section / Obtener campos de una seccion"""
        return [
            name for name, spec in self.fields.items()
            if spec.get("section") == section_id
        ]

    def _should_show_field(self, field_spec: dict, data: dict) -> bool:
        """Check if field should be visible / Verificar si campo debe ser visible"""
        condition = field_spec.get("condition")
        if not condition:
            return True
        return evaluate_condition(condition, data)

    def _render_field(self, field_name: str, field_spec: dict, data: dict) -> Any:
        """
        Render a single field
        Renderizar un campo individual

        Args:
            field_name: Field name
            field_spec: Field specification
            data: Current form data

        Returns:
            Field value
        """
        field_type = field_spec.get("type", "text")
        label = field_spec.get("label", field_name)
        required = field_spec.get("required", False)
        default = field_spec.get("default")

        # Add required indicator
        if required:
            label = f"{label} *"

        current_value = data.get(field_name, default)
        key = get_stable_key(field_name)

        # Check if field is disabled
        disabled = False
        editable_when = field_spec.get("editable_when")
        if editable_when:
            disabled = not evaluate_condition(editable_when, data)

        # Render based on type
        if field_type == "text":
            if field_spec.get("multiline"):
                return st.text_area(
                    label,
                    value=current_value or "",
                    key=key,
                    placeholder=field_spec.get("placeholder", ""),
                    disabled=disabled
                )
            else:
                return st.text_input(
                    label,
                    value=current_value or "",
                    key=key,
                    placeholder=field_spec.get("placeholder", ""),
                    disabled=disabled
                )

        elif field_type == "date":
            if isinstance(current_value, str):
                parsed = parse_date_string(current_value)
                current_value = parsed if parsed else datetime.now().date()
            elif current_value is None or current_value == "today":
                current_value = datetime.now().date()

            selected_date = st.date_input(
                label,
                value=current_value,
                key=key,
                disabled=disabled
            )
            # Return formatted date string
            return format_spanish_date(selected_date)

        elif field_type == "bool":
            current_bool = False
            if isinstance(current_value, bool):
                current_bool = current_value
            elif isinstance(current_value, str):
                current_bool = current_value.lower() in ('true', 'si', 'yes', '1', 'sí')

            return st.checkbox(
                label,
                value=current_bool,
                key=key,
                disabled=disabled
            )

        elif field_type == "enum":
            values = field_spec.get("values", [])
            options = [v.get("value") for v in values]
            labels = {v.get("value"): v.get("label") for v in values}

            # Find current index
            current_index = 0
            if current_value and current_value in options:
                current_index = options.index(current_value)

            return st.selectbox(
                label,
                options=options,
                index=current_index,
                key=key,
                format_func=lambda x: labels.get(x, x),
                disabled=disabled
            )

        elif field_type == "int":
            current_int = 0
            if current_value:
                try:
                    current_int = int(current_value)
                except (ValueError, TypeError):
                    pass

            return st.number_input(
                label,
                value=current_int,
                key=key,
                disabled=disabled
            )

        elif field_type == "list":
            return self._render_list_field(field_name, field_spec, data)

        # Default to text
        return st.text_input(
            label,
            value=str(current_value) if current_value else "",
            key=key,
            disabled=disabled
        )

    def _render_list_field(self, field_name: str, field_spec: dict, data: dict) -> List[Dict]:
        """
        Render a list field (like lista_alto_directores)
        Renderizar un campo de lista

        Args:
            field_name: Field name
            field_spec: Field specification
            data: Current form data

        Returns:
            List of item dictionaries
        """
        label = field_spec.get("label", field_name)
        item_schema = field_spec.get("item_schema", {})

        st.markdown(f"**{label}**")
        st.info("Introduce los datos. Estos se incluiran en el documento.")

        # Get current items
        items = get_list_items(field_name)

        # Number input for items
        num_items = st.number_input(
            f"Numero de {label.lower()}",
            min_value=0,
            max_value=20,
            value=len(items) if items else 2,
            key=f"num_{field_name}"
        )

        # Adjust list size
        while len(items) < num_items:
            add_list_item(field_name, {})
            items = get_list_items(field_name)

        while len(items) > num_items:
            if items:
                remove_list_item(field_name, items[-1].get("_id", ""))
                items = get_list_items(field_name)

        # Render item inputs
        result_items = []
        for i, item in enumerate(items):
            cols = st.columns(len(item_schema))

            item_data = {"_id": item.get("_id", str(i))}

            for j, (sub_field, sub_spec) in enumerate(item_schema.items()):
                with cols[j]:
                    sub_label = sub_spec.get("label", sub_field)
                    sub_value = item.get(sub_field, "")
                    key = get_stable_key(field_name, i, sub_field)

                    new_value = st.text_input(
                        f"{sub_label} {i+1}",
                        value=sub_value,
                        key=key
                    )
                    item_data[sub_field] = new_value

            result_items.append(item_data)

        # Update session state
        st.session_state.list_items[field_name] = result_items

        return result_items

    def render_oficina_section(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render office selection section with auto-fill
        Renderizar seccion de seleccion de oficina con auto-relleno

        Args:
            data: Current form data

        Returns:
            Updated data with office fields
        """
        result = dict(data)

        # Office selection
        oficina_options = list(self.oficinas.keys())
        display_names = {
            k: v.get("display_name", k)
            for k, v in self.oficinas.items()
        }

        current_oficina = result.get("Oficina_Seleccionada", "BARCELONA")
        if current_oficina not in oficina_options:
            current_oficina = oficina_options[0] if oficina_options else "BARCELONA"

        oficina_sel = st.selectbox(
            "Selecciona la oficina",
            oficina_options,
            index=oficina_options.index(current_oficina) if current_oficina in oficina_options else 0,
            format_func=lambda x: display_names.get(x, x)
        )
        result["Oficina_Seleccionada"] = oficina_sel

        # Auto-fill office data
        oficina_data = self.oficinas.get(oficina_sel, {})
        is_custom = oficina_data.get("editable", False) or oficina_sel == "PERSONALIZADA"

        # Address fields
        for campo in ["Direccion_Oficina", "CP", "Ciudad_Oficina"]:
            # Auto-fill if not custom
            if not is_custom and campo not in result:
                result[campo] = oficina_data.get(campo, "")

            field_spec = self.fields.get(campo, {})
            label = field_spec.get("label", campo)

            result[campo] = st.text_input(
                label,
                value=result.get(campo, oficina_data.get(campo, "")),
                key=get_stable_key(campo),
                disabled=not is_custom
            )

        return result

    def render_conditional_section(
        self,
        condition_field: str,
        dependent_fields: List[str],
        data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Render a conditional section that shows/hides based on a condition
        Renderizar seccion condicional que muestra/oculta basado en condicion

        Args:
            condition_field: Field that controls visibility
            dependent_fields: Fields to show when condition is true
            data: Current form data

        Returns:
            Updated form data
        """
        result = dict(data)

        # Check if condition is met
        if result.get(condition_field):
            with st.container():
                st.markdown("##### Detalles adicionales")
                for field_name in dependent_fields:
                    field_spec = self.fields.get(field_name, {})
                    value = self._render_field(field_name, field_spec, result)
                    result[field_name] = value

        return result
