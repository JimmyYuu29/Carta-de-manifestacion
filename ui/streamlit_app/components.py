"""
Components - Reusable UI components
Componentes de UI reutilizables
"""

import streamlit as st
from typing import Optional, Callable
from pathlib import Path
import io


def render_header(title: str, subtitle: Optional[str] = None, icon: str = "page_facing_up") -> None:
    """
    Render page header
    Renderizar cabecera de pagina

    Args:
        title: Main title
        subtitle: Optional subtitle
        icon: Emoji icon
    """
    st.title(f"{title}")
    if subtitle:
        st.markdown(f"**{subtitle}**")
    st.markdown("---")


def render_section_header(title: str, icon: Optional[str] = None) -> None:
    """
    Render section header
    Renderizar cabecera de seccion

    Args:
        title: Section title
        icon: Optional emoji icon
    """
    if icon:
        st.markdown(f"### {icon} {title}")
    else:
        st.markdown(f"### {title}")


def render_success_message(message: str) -> None:
    """
    Render success message
    Renderizar mensaje de exito

    Args:
        message: Success message
    """
    st.success(f"Completado: {message}")


def render_error_message(message: str) -> None:
    """
    Render error message
    Renderizar mensaje de error

    Args:
        message: Error message
    """
    st.error(f"Error: {message}")


def render_warning_message(message: str) -> None:
    """
    Render warning message
    Renderizar mensaje de advertencia

    Args:
        message: Warning message
    """
    st.warning(f"Advertencia: {message}")


def render_info_message(message: str) -> None:
    """
    Render info message
    Renderizar mensaje informativo

    Args:
        message: Info message
    """
    st.info(message)


def render_download_button(
    label: str,
    data: bytes,
    file_name: str,
    mime_type: str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
) -> bool:
    """
    Render download button
    Renderizar boton de descarga

    Args:
        label: Button label
        data: File data bytes
        file_name: Name for downloaded file
        mime_type: MIME type of file

    Returns:
        True if button was clicked
    """
    return st.download_button(
        label=f"Descargar {label}",
        data=data,
        file_name=file_name,
        mime=mime_type
    )


def render_file_uploader(
    label: str,
    file_types: list,
    key: str,
    help_text: Optional[str] = None
) -> Optional[io.BytesIO]:
    """
    Render file uploader
    Renderizar cargador de archivos

    Args:
        label: Uploader label
        file_types: List of allowed file extensions
        key: Unique key for widget
        help_text: Optional help text

    Returns:
        Uploaded file or None
    """
    return st.file_uploader(
        label,
        type=file_types,
        key=key,
        help=help_text
    )


def render_text_input(
    label: str,
    key: str,
    value: str = "",
    placeholder: str = "",
    disabled: bool = False,
    on_change: Optional[Callable] = None
) -> str:
    """
    Render text input
    Renderizar entrada de texto

    Args:
        label: Input label
        key: Unique key
        value: Initial value
        placeholder: Placeholder text
        disabled: Whether input is disabled
        on_change: Optional callback

    Returns:
        Input value
    """
    return st.text_input(
        label,
        value=value,
        key=key,
        placeholder=placeholder,
        disabled=disabled,
        on_change=on_change
    )


def render_text_area(
    label: str,
    key: str,
    value: str = "",
    placeholder: str = "",
    height: int = 100
) -> str:
    """
    Render text area
    Renderizar area de texto

    Args:
        label: Input label
        key: Unique key
        value: Initial value
        placeholder: Placeholder text
        height: Height in pixels

    Returns:
        Input value
    """
    return st.text_area(
        label,
        value=value,
        key=key,
        placeholder=placeholder,
        height=height
    )


def render_selectbox(
    label: str,
    options: list,
    key: str,
    index: int = 0,
    format_func: Optional[Callable] = None
) -> str:
    """
    Render select box
    Renderizar caja de seleccion

    Args:
        label: Select label
        options: List of options
        key: Unique key
        index: Default selected index
        format_func: Optional display formatter

    Returns:
        Selected value
    """
    if format_func:
        return st.selectbox(label, options, index=index, key=key, format_func=format_func)
    return st.selectbox(label, options, index=index, key=key)


def render_checkbox(
    label: str,
    key: str,
    value: bool = False
) -> bool:
    """
    Render checkbox
    Renderizar casilla de verificacion

    Args:
        label: Checkbox label
        key: Unique key
        value: Initial value

    Returns:
        Checkbox state
    """
    return st.checkbox(label, value=value, key=key)


def render_date_input(
    label: str,
    key: str,
    value=None
):
    """
    Render date input
    Renderizar entrada de fecha

    Args:
        label: Input label
        key: Unique key
        value: Initial date value

    Returns:
        Selected date
    """
    from datetime import datetime
    if value is None:
        value = datetime.now().date()
    return st.date_input(label, value=value, key=key)


def render_number_input(
    label: str,
    key: str,
    value: int = 0,
    min_value: int = 0,
    max_value: int = 100,
    step: int = 1
) -> int:
    """
    Render number input
    Renderizar entrada numerica

    Args:
        label: Input label
        key: Unique key
        value: Initial value
        min_value: Minimum value
        max_value: Maximum value
        step: Step size

    Returns:
        Input value
    """
    return st.number_input(
        label,
        value=value,
        min_value=min_value,
        max_value=max_value,
        step=step,
        key=key
    )


def render_button(
    label: str,
    key: Optional[str] = None,
    type: str = "secondary",
    use_container_width: bool = False
) -> bool:
    """
    Render button
    Renderizar boton

    Args:
        label: Button label
        key: Optional unique key
        type: Button type (primary/secondary)
        use_container_width: Use full container width

    Returns:
        True if clicked
    """
    return st.button(
        label,
        key=key,
        type=type,
        use_container_width=use_container_width
    )


def render_columns(num_columns: int = 2):
    """
    Create columns layout
    Crear disposicion de columnas

    Args:
        num_columns: Number of columns

    Returns:
        Tuple of column objects
    """
    return st.columns(num_columns)


def render_expander(label: str, expanded: bool = False):
    """
    Create expander
    Crear expansor

    Args:
        label: Expander label
        expanded: Whether expanded by default

    Returns:
        Expander context manager
    """
    return st.expander(label, expanded=expanded)


def render_container():
    """
    Create container
    Crear contenedor

    Returns:
        Container context manager
    """
    return st.container()


def render_divider() -> None:
    """
    Render horizontal divider
    Renderizar divisor horizontal
    """
    st.markdown("---")


def render_spinner(text: str = "Procesando..."):
    """
    Create spinner context
    Crear contexto de spinner

    Args:
        text: Spinner text

    Returns:
        Spinner context manager
    """
    return st.spinner(text)


def render_progress_bar(value: float, text: str = "") -> None:
    """
    Render progress bar
    Renderizar barra de progreso

    Args:
        value: Progress value (0-1)
        text: Optional text
    """
    st.progress(value, text=text)


def render_metric(label: str, value: str, delta: Optional[str] = None) -> None:
    """
    Render metric
    Renderizar metrica

    Args:
        label: Metric label
        value: Metric value
        delta: Optional delta value
    """
    st.metric(label=label, value=value, delta=delta)


def render_code(code: str, language: str = "python") -> None:
    """
    Render code block
    Renderizar bloque de codigo

    Args:
        code: Code string
        language: Programming language
    """
    st.code(code, language=language)
