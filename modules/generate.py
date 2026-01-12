"""
Generate - Unified document generation entry point
Punto de entrada unificado para generacion de documentos
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Any, Dict
from datetime import date, datetime
import uuid
import time

from .plugin_loader import load_plugin, PluginPack
from .contract_validator import validate_input, ValidationResult
from .renderer_docx import DocxRenderer
from .rule_engine import EvaluationTrace


@dataclass
class GenerationResult:
    """Result of document generation / Resultado de la generacion de documento"""
    success: bool
    output_path: Optional[Path]
    trace_id: str
    validation_errors: List[str] = field(default_factory=list)
    evaluation_traces: List[EvaluationTrace] = field(default_factory=list)
    error: Optional[str] = None
    duration_ms: int = 0


def generate(
    plugin_id: str,
    data: dict,
    output_dir: Path = Path("output"),
    template_path: Optional[Path] = None,
    should_validate: bool = True,
    filename_prefix: Optional[str] = None
) -> GenerationResult:
    """
    Unified entry point for document generation
    Punto de entrada unificado para generacion de documentos

    Args:
        plugin_id: ID of the plugin to use
        data: Input data dictionary
        output_dir: Directory for output files
        template_path: Optional custom template path
        should_validate: Whether to validate input before generation
        filename_prefix: Optional prefix for output filename

    Returns:
        GenerationResult with success status and details
    """
    start_time = time.time()
    trace_id = str(uuid.uuid4())

    try:
        # 1. Load plugin
        plugin = load_plugin(plugin_id)

        # 2. Preprocess input
        data = preprocess_input(data, plugin)

        # 3. Validate (optional)
        if should_validate:
            validation_result = validate_input(plugin, data)
            if not validation_result.is_valid:
                error_messages = [f"{e.field}: {e.message}" for e in validation_result.errors]
                return GenerationResult(
                    success=False,
                    output_path=None,
                    trace_id=trace_id,
                    validation_errors=error_messages,
                    evaluation_traces=[],
                    error="Validation failed / Validacion fallida",
                    duration_ms=int((time.time() - start_time) * 1000)
                )

        # 4. Render document
        renderer = DocxRenderer(plugin)

        # Generate output filename
        if filename_prefix:
            filename = f"{filename_prefix}_{trace_id[:8]}.docx"
        else:
            client_name = data.get("Nombre_Cliente", "documento")
            client_name = client_name.replace(" ", "_").replace("/", "_")
            timestamp = datetime.now().strftime("%Y%m%d")
            filename = f"Carta_Manifestacion_{client_name}_{timestamp}.docx"

        output_path = output_dir / filename

        # Render
        output_path, traces = renderer.render(data, output_path, template_path)

        return GenerationResult(
            success=True,
            output_path=output_path,
            trace_id=trace_id,
            validation_errors=[],
            evaluation_traces=traces,
            error=None,
            duration_ms=int((time.time() - start_time) * 1000)
        )

    except FileNotFoundError as e:
        return GenerationResult(
            success=False,
            output_path=None,
            trace_id=trace_id,
            validation_errors=[],
            evaluation_traces=[],
            error=f"Template not found / Plantilla no encontrada: {e}",
            duration_ms=int((time.time() - start_time) * 1000)
        )

    except Exception as e:
        return GenerationResult(
            success=False,
            output_path=None,
            trace_id=trace_id,
            validation_errors=[],
            evaluation_traces=[],
            error=str(e),
            duration_ms=int((time.time() - start_time) * 1000)
        )


def preprocess_input(data: dict, plugin: PluginPack) -> dict:
    """
    Preprocess input data: type conversions
    Preprocesar datos de entrada: conversiones de tipo

    Args:
        data: Raw input data
        plugin: PluginPack instance

    Returns:
        Preprocessed data dictionary
    """
    fields = plugin.fields.get("fields", {})
    result = dict(data)

    for field_name, spec in fields.items():
        if field_name not in result:
            continue

        value = result[field_name]
        field_type = spec.get("type")

        # Date string conversion
        if field_type == "date" and isinstance(value, str):
            parsed = parse_date_value(value)
            if parsed:
                result[field_name] = parsed

        # Integer conversion
        elif field_type == "int" and isinstance(value, str):
            try:
                result[field_name] = int(value.replace(",", "").replace(".", ""))
            except ValueError:
                pass

        # Currency conversion
        elif field_type == "currency" and isinstance(value, str):
            try:
                clean_value = value.replace(",", "").replace(".", "").replace(" ", "").replace("EUR", "").replace("€", "")
                result[field_name] = int(clean_value)
            except ValueError:
                pass

        # Boolean conversion
        elif field_type == "bool":
            if isinstance(value, str):
                result[field_name] = value.lower() in ('true', 'si', 'yes', '1', 'sí')
            elif isinstance(value, (int, float)):
                result[field_name] = bool(value)

    return result


def parse_date_value(value: str) -> Optional[date]:
    """Parse date from string / Parsear fecha desde string"""
    if not value:
        return None

    formats = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None


def generate_from_form(
    plugin_id: str,
    form_data: dict,
    list_data: dict,
    output_dir: Path = Path("output"),
    template_path: Optional[Path] = None
) -> GenerationResult:
    """
    Generate document from Streamlit form data
    Generar documento desde datos de formulario Streamlit

    This is a convenience wrapper that combines form_data and list_data
    before calling the main generate function.

    Args:
        plugin_id: ID of the plugin
        form_data: Simple form field values
        list_data: List field values (like lista_alto_directores)
        output_dir: Output directory
        template_path: Optional template path

    Returns:
        GenerationResult
    """
    # Combine data
    data = dict(form_data)

    # Add list data, cleaning internal IDs
    for field_name, items in list_data.items():
        clean_items = []
        for item in items:
            if isinstance(item, dict):
                clean_item = {k: v for k, v in item.items() if not k.startswith("_")}
                clean_items.append(clean_item)
            else:
                clean_items.append(item)
        data[field_name] = clean_items

    return generate(
        plugin_id=plugin_id,
        data=data,
        output_dir=output_dir,
        template_path=template_path
    )
