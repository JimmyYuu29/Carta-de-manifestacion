"""
State Store - Session state management for Streamlit
Gestion del estado de sesion para Streamlit
"""

import streamlit as st
from typing import Any, Dict, List, Optional
import uuid


def init_session_state(plugin_id: str) -> None:
    """
    Initialize session state
    Inicializar estado de sesion

    Args:
        plugin_id: ID of the plugin being used
    """
    if "plugin_id" not in st.session_state:
        st.session_state.plugin_id = plugin_id

    if "form_data" not in st.session_state:
        st.session_state.form_data = {}

    if "list_items" not in st.session_state:
        st.session_state.list_items = {}

    if "generation_result" not in st.session_state:
        st.session_state.generation_result = None

    if "imported_data" not in st.session_state:
        st.session_state.imported_data = {}

    if "field_visibility" not in st.session_state:
        st.session_state.field_visibility = {}


def get_stable_key(field_name: str, index: Optional[int] = None, sub_field: Optional[str] = None) -> str:
    """
    Generate stable widget key
    Generar clave estable de widget

    Args:
        field_name: Name of the field
        index: Optional index for list items
        sub_field: Optional sub-field name

    Returns:
        Stable key string
    """
    key = f"field_{field_name}"
    if index is not None:
        key += f"_{index}"
    if sub_field:
        key += f"_{sub_field}"
    return key


def get_field_value(field_name: str, default: Any = None) -> Any:
    """
    Get field value from session state
    Obtener valor de campo del estado de sesion

    Args:
        field_name: Name of the field
        default: Default value if not found

    Returns:
        Field value or default
    """
    return st.session_state.form_data.get(field_name, default)


def set_field_value(field_name: str, value: Any) -> None:
    """
    Set field value in session state
    Establecer valor de campo en estado de sesion

    Args:
        field_name: Name of the field
        value: Value to set
    """
    st.session_state.form_data[field_name] = value


def get_list_items(field_name: str) -> List[Dict[str, Any]]:
    """
    Get list items from session state
    Obtener elementos de lista del estado de sesion

    Args:
        field_name: Name of the list field

    Returns:
        List of item dictionaries
    """
    if field_name not in st.session_state.list_items:
        st.session_state.list_items[field_name] = []
    return st.session_state.list_items[field_name]


def add_list_item(field_name: str, item: Dict[str, Any]) -> None:
    """
    Add item to list in session state
    Anadir elemento a lista en estado de sesion

    Args:
        field_name: Name of the list field
        item: Item dictionary to add
    """
    items = get_list_items(field_name)
    item["_id"] = str(uuid.uuid4())
    items.append(item)
    st.session_state.list_items[field_name] = items


def remove_list_item(field_name: str, item_id: str) -> None:
    """
    Remove item from list in session state
    Eliminar elemento de lista en estado de sesion

    Args:
        field_name: Name of the list field
        item_id: ID of item to remove
    """
    items = get_list_items(field_name)
    st.session_state.list_items[field_name] = [
        item for item in items if item.get("_id") != item_id
    ]


def update_list_item(field_name: str, item_id: str, updates: Dict[str, Any]) -> None:
    """
    Update item in list
    Actualizar elemento en lista

    Args:
        field_name: Name of the list field
        item_id: ID of item to update
        updates: Dictionary of updates
    """
    items = get_list_items(field_name)
    for item in items:
        if item.get("_id") == item_id:
            item.update(updates)
            break


def get_all_form_data() -> Dict[str, Any]:
    """
    Collect all form data including lists
    Recopilar todos los datos del formulario incluyendo listas

    Returns:
        Complete form data dictionary
    """
    data = dict(st.session_state.form_data)

    # Add list data
    for field_name, items in st.session_state.list_items.items():
        # Clean internal IDs
        clean_items = [
            {k: v for k, v in item.items() if not k.startswith("_")}
            for item in items
        ]
        data[field_name] = clean_items

    return data


def clear_form_data() -> None:
    """
    Clear all form data
    Limpiar todos los datos del formulario
    """
    st.session_state.form_data = {}
    st.session_state.list_items = {}
    st.session_state.generation_result = None
    st.session_state.imported_data = {}


def set_imported_data(data: Dict[str, Any]) -> None:
    """
    Set imported data and merge with form data
    Establecer datos importados y combinar con datos del formulario

    Args:
        data: Imported data dictionary
    """
    st.session_state.imported_data = data

    # Merge with form data
    for key, value in data.items():
        if key not in st.session_state.form_data or st.session_state.form_data[key] == "":
            st.session_state.form_data[key] = value


def get_imported_data() -> Dict[str, Any]:
    """
    Get imported data
    Obtener datos importados

    Returns:
        Imported data dictionary
    """
    return st.session_state.get("imported_data", {})


def update_field_visibility(visibility: Dict[str, bool]) -> None:
    """
    Update field visibility map
    Actualizar mapa de visibilidad de campos

    Args:
        visibility: Dictionary mapping field names to visibility
    """
    st.session_state.field_visibility = visibility


def is_field_visible(field_name: str) -> bool:
    """
    Check if field should be visible
    Verificar si campo debe ser visible

    Args:
        field_name: Name of the field

    Returns:
        True if field is visible
    """
    return st.session_state.field_visibility.get(field_name, True)
