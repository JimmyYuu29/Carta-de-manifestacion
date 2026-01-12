"""
Contract Validator - Multi-layer data validation
Validador de datos multicapa
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from datetime import date
import re

from .plugin_loader import PluginPack
from .dsl_evaluator import evaluate_condition


@dataclass
class ValidationError:
    """Single validation error / Error de validacion individual"""
    field: str
    message: str
    code: str
    value: Any = None


@dataclass
class ValidationResult:
    """Result of validation / Resultado de la validacion"""
    is_valid: bool
    errors: List[ValidationError] = field(default_factory=list)
    warnings: List[ValidationError] = field(default_factory=list)

    def add_error(self, field_name: str, message: str, code: str, value: Any = None):
        self.errors.append(ValidationError(field=field_name, message=message, code=code, value=value))
        self.is_valid = False

    def add_warning(self, field_name: str, message: str, code: str, value: Any = None):
        self.warnings.append(ValidationError(field=field_name, message=message, code=code, value=value))


class ContractValidator:
    """
    Multi-layer validator for input data
    Validador multicapa para datos de entrada
    """

    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self.fields = plugin.fields.get("fields", {})

    def validate(self, data: dict, check_required: bool = True) -> ValidationResult:
        """
        Validate input data against field definitions
        Validar datos de entrada contra definiciones de campos

        Args:
            data: Input data dictionary
            check_required: Whether to check required fields

        Returns:
            ValidationResult with errors and warnings
        """
        result = ValidationResult(is_valid=True)

        for field_name, field_spec in self.fields.items():
            # Skip validation for hidden fields
            condition = field_spec.get("condition")
            if condition and not evaluate_condition(condition, data):
                continue

            value = data.get(field_name)

            # Required check
            if check_required and field_spec.get("required", False):
                if value is None or (isinstance(value, str) and value.strip() == ""):
                    result.add_error(
                        field_name,
                        f"El campo '{field_spec.get('label', field_name)}' es requerido",
                        "required"
                    )
                    continue

            # Type validation
            if value is not None:
                self._validate_type(field_name, field_spec, value, result)

            # Custom validation rules
            validation = field_spec.get("validation", {})
            if validation and value is not None:
                self._validate_rules(field_name, field_spec, value, validation, result)

        return result

    def _validate_type(self, field_name: str, field_spec: dict, value: Any, result: ValidationResult):
        """Validate field type / Validar tipo de campo"""
        field_type = field_spec.get("type", "text")
        label = field_spec.get("label", field_name)

        if field_type == "text":
            if not isinstance(value, str):
                result.add_error(field_name, f"'{label}' debe ser texto", "type_error", value)

        elif field_type == "int":
            if not isinstance(value, int) and value != "":
                try:
                    int(value)
                except (ValueError, TypeError):
                    result.add_error(field_name, f"'{label}' debe ser un numero entero", "type_error", value)

        elif field_type == "decimal" or field_type == "currency":
            if not isinstance(value, (int, float)) and value != "":
                try:
                    float(str(value).replace(",", "."))
                except (ValueError, TypeError):
                    result.add_error(field_name, f"'{label}' debe ser un numero", "type_error", value)

        elif field_type == "bool":
            if not isinstance(value, bool) and value not in ('si', 'no', 'yes', 'no', True, False):
                result.add_error(field_name, f"'{label}' debe ser verdadero/falso", "type_error", value)

        elif field_type == "date":
            if not isinstance(value, date) and value != "":
                if isinstance(value, str):
                    if not self._is_valid_date_string(value):
                        result.add_error(field_name, f"'{label}' debe ser una fecha valida", "type_error", value)

        elif field_type == "enum":
            valid_values = [v.get("value") for v in field_spec.get("values", [])]
            if value and value not in valid_values:
                result.add_error(
                    field_name,
                    f"'{label}' debe ser uno de: {', '.join(valid_values)}",
                    "enum_error",
                    value
                )

        elif field_type == "list":
            if not isinstance(value, list):
                result.add_error(field_name, f"'{label}' debe ser una lista", "type_error", value)
            else:
                # Validate each item in the list
                item_schema = field_spec.get("item_schema", {})
                for i, item in enumerate(value):
                    if not isinstance(item, dict):
                        continue
                    for item_field, item_spec in item_schema.items():
                        item_value = item.get(item_field)
                        if item_spec.get("required", False) and not item_value:
                            result.add_error(
                                f"{field_name}[{i}].{item_field}",
                                f"'{item_spec.get('label', item_field)}' es requerido",
                                "required"
                            )

    def _validate_rules(self, field_name: str, field_spec: dict, value: Any, validation: dict, result: ValidationResult):
        """Validate custom rules / Validar reglas personalizadas"""
        label = field_spec.get("label", field_name)

        # Max length
        max_length = validation.get("max_length")
        if max_length and isinstance(value, str) and len(value) > max_length:
            result.add_error(
                field_name,
                f"'{label}' no puede exceder {max_length} caracteres",
                "max_length",
                value
            )

        # Min length
        min_length = validation.get("min_length")
        if min_length and isinstance(value, str) and len(value) < min_length:
            result.add_error(
                field_name,
                f"'{label}' debe tener al menos {min_length} caracteres",
                "min_length",
                value
            )

        # Pattern (regex)
        pattern = validation.get("pattern")
        if pattern and isinstance(value, str):
            if not re.match(pattern, value):
                result.add_error(
                    field_name,
                    f"'{label}' no tiene el formato correcto",
                    "pattern",
                    value
                )

        # Min value
        min_val = validation.get("min")
        if min_val is not None:
            try:
                if float(value) < float(min_val):
                    result.add_error(
                        field_name,
                        f"'{label}' debe ser al menos {min_val}",
                        "min_value",
                        value
                    )
            except (ValueError, TypeError):
                pass

        # Max value
        max_val = validation.get("max")
        if max_val is not None:
            try:
                if float(value) > float(max_val):
                    result.add_error(
                        field_name,
                        f"'{label}' no puede exceder {max_val}",
                        "max_value",
                        value
                    )
            except (ValueError, TypeError):
                pass

    def _is_valid_date_string(self, value: str) -> bool:
        """Check if string is a valid date / Verificar si string es fecha valida"""
        formats = [
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d de %B de %Y",
            "%Y/%m/%d",
        ]

        for fmt in formats:
            try:
                from datetime import datetime
                datetime.strptime(value, fmt)
                return True
            except ValueError:
                continue

        return False


def validate_input(plugin: PluginPack, data: dict, check_required: bool = True) -> ValidationResult:
    """
    Convenience function to validate input data
    Funcion de conveniencia para validar datos de entrada

    Args:
        plugin: PluginPack instance
        data: Input data dictionary
        check_required: Whether to check required fields

    Returns:
        ValidationResult
    """
    validator = ContractValidator(plugin)
    return validator.validate(data, check_required)
