"""
Contract Models - Dynamic Pydantic model generation
Modelos dinamicos de Pydantic
"""

from typing import Any, Dict, List, Optional, Type
from datetime import date
from decimal import Decimal

try:
    from pydantic import BaseModel, create_model, Field
    PYDANTIC_AVAILABLE = True
except ImportError:
    PYDANTIC_AVAILABLE = False
    BaseModel = object

from .plugin_loader import PluginPack


# Type mapping from YAML to Python types
TYPE_MAP = {
    "text": str,
    "date": date,
    "currency": Decimal,
    "int": int,
    "decimal": float,
    "bool": bool,
    "enum": str,
    "list": list,
}


def build_pydantic_model(plugin: PluginPack, model_name: str = "DynamicInputModel") -> Type[BaseModel]:
    """
    Dynamically generate a Pydantic model from YAML field definitions
    Generar dinamicamente un modelo Pydantic desde definiciones YAML

    Args:
        plugin: PluginPack instance with field definitions
        model_name: Name for the generated model class

    Returns:
        Dynamically created Pydantic model class

    Raises:
        ImportError: If Pydantic is not installed
    """
    if not PYDANTIC_AVAILABLE:
        raise ImportError("Pydantic is required for model generation. Install with: pip install pydantic")

    fields = plugin.fields.get("fields", {})
    field_definitions: Dict[str, Any] = {}

    for name, spec in fields.items():
        field_type = TYPE_MAP.get(spec.get("type", "text"), str)
        required = spec.get("required", False)
        default = spec.get("default")
        description = spec.get("label", name)

        # Handle list types
        if spec.get("type") == "list":
            item_schema = spec.get("item_schema", {})
            field_type = List[Dict[str, Any]]
            if not required:
                default = default if default is not None else []

        # Handle enum types
        if spec.get("type") == "enum":
            field_type = str
            values = spec.get("values", [])
            if values:
                valid_values = [v.get("value") for v in values]
                # Could add validation here

        # Build field definition
        if required:
            field_definitions[name] = (field_type, Field(..., description=description))
        else:
            if default is None and field_type in (str,):
                default = ""
            elif default is None and field_type == bool:
                default = False
            elif default is None and field_type == list:
                default = []
            field_definitions[name] = (Optional[field_type], Field(default=default, description=description))

    return create_model(model_name, **field_definitions)


def create_field_schema(plugin: PluginPack) -> Dict[str, Dict[str, Any]]:
    """
    Create a JSON-schema-like representation of fields
    Crear una representacion tipo JSON-schema de los campos

    Args:
        plugin: PluginPack instance

    Returns:
        Dictionary with field schemas
    """
    fields = plugin.fields.get("fields", {})
    schema = {}

    for name, spec in fields.items():
        field_schema = {
            "type": spec.get("type", "text"),
            "label": spec.get("label", name),
            "required": spec.get("required", False),
            "section": spec.get("section", "default"),
        }

        if "default" in spec:
            field_schema["default"] = spec["default"]

        if "validation" in spec:
            field_schema["validation"] = spec["validation"]

        if "values" in spec:
            field_schema["enum"] = [v.get("value") for v in spec["values"]]
            field_schema["enum_labels"] = {v.get("value"): v.get("label") for v in spec["values"]}

        if "condition" in spec:
            field_schema["condition"] = spec["condition"]

        if "item_schema" in spec:
            field_schema["item_schema"] = spec["item_schema"]

        schema[name] = field_schema

    return schema


def get_default_values(plugin: PluginPack) -> Dict[str, Any]:
    """
    Get default values for all fields
    Obtener valores por defecto para todos los campos

    Args:
        plugin: PluginPack instance

    Returns:
        Dictionary with default values
    """
    fields = plugin.fields.get("fields", {})
    defaults = {}

    for name, spec in fields.items():
        field_type = spec.get("type", "text")
        default = spec.get("default")

        if default is not None:
            if default == "today" and field_type == "date":
                defaults[name] = date.today()
            else:
                defaults[name] = default
        else:
            # Set type-appropriate defaults
            if field_type == "bool":
                defaults[name] = False
            elif field_type == "list":
                defaults[name] = []
            elif field_type in ("text", "enum"):
                defaults[name] = ""
            elif field_type in ("int", "currency", "decimal"):
                defaults[name] = 0

    return defaults
