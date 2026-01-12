"""
DSL Evaluator - Safe conditional expression evaluator
Evaluador seguro de expresiones condicionales DSL
"""

from typing import Any, Dict, Optional

# Allowed operators (whitelist) / Operadores permitidos (lista blanca)
ALLOWED_OPERATORS = frozenset({
    "and", "or", "not",
    "equals", "not_equals", "gt", "gte", "lt", "lte",
    "in", "not_in",
    "exists", "not_exists", "is_empty", "not_empty",
    "contains", "not_contains",
})

# Maximum nesting depth to prevent stack overflow
MAX_NESTING_DEPTH = 5


class DSLEvaluationError(Exception):
    """Exception raised for DSL evaluation errors"""
    pass


def evaluate_condition(condition: dict, data: dict, depth: int = 0) -> bool:
    """
    Safe condition evaluation (no eval/exec)
    Evaluacion segura de condiciones (sin eval/exec)

    Args:
        condition: Condition dictionary with operator, field, value, etc.
        data: Data dictionary to evaluate against
        depth: Current nesting depth (internal use)

    Returns:
        bool: Result of condition evaluation

    Raises:
        DSLEvaluationError: If condition is invalid or too deeply nested
    """
    if not condition:
        return True

    if depth > MAX_NESTING_DEPTH:
        raise DSLEvaluationError(f"Condition nesting too deep (max {MAX_NESTING_DEPTH})")

    operator = condition.get("operator")
    if not operator:
        return True

    if operator not in ALLOWED_OPERATORS:
        raise DSLEvaluationError(f"Operator not allowed: {operator}")

    # Logical operators / Operadores logicos
    if operator == "and":
        conditions = condition.get("conditions", [])
        if not conditions:
            return True
        return all(evaluate_condition(c, data, depth + 1) for c in conditions)

    if operator == "or":
        conditions = condition.get("conditions", [])
        if not conditions:
            return False
        return any(evaluate_condition(c, data, depth + 1) for c in conditions)

    if operator == "not":
        inner_condition = condition.get("condition")
        if not inner_condition:
            return True
        return not evaluate_condition(inner_condition, data, depth + 1)

    # Comparison operators / Operadores de comparacion
    field = condition.get("field")
    value = condition.get("value")
    field_value = get_nested_value(data, field) if field else None

    # Normalize boolean values
    if isinstance(value, str):
        if value.lower() in ('true', 'si', 'yes'):
            value = True
        elif value.lower() in ('false', 'no'):
            value = False

    if isinstance(field_value, str):
        if field_value.lower() in ('true', 'si', 'yes'):
            field_value = True
        elif field_value.lower() in ('false', 'no'):
            field_value = False

    if operator == "equals":
        return field_value == value

    if operator == "not_equals":
        return field_value != value

    if operator == "gt":
        try:
            return float(field_value) > float(value) if field_value is not None else False
        except (ValueError, TypeError):
            return False

    if operator == "gte":
        try:
            return float(field_value) >= float(value) if field_value is not None else False
        except (ValueError, TypeError):
            return False

    if operator == "lt":
        try:
            return float(field_value) < float(value) if field_value is not None else False
        except (ValueError, TypeError):
            return False

    if operator == "lte":
        try:
            return float(field_value) <= float(value) if field_value is not None else False
        except (ValueError, TypeError):
            return False

    if operator == "in":
        values = condition.get("values", [])
        return field_value in values

    if operator == "not_in":
        values = condition.get("values", [])
        return field_value not in values

    if operator == "exists":
        return field_value is not None

    if operator == "not_exists":
        return field_value is None

    if operator == "is_empty":
        if field_value is None:
            return True
        if isinstance(field_value, (str, list, dict)):
            return len(field_value) == 0
        return False

    if operator == "not_empty":
        if field_value is None:
            return False
        if isinstance(field_value, (str, list, dict)):
            return len(field_value) > 0
        return True

    if operator == "contains":
        if field_value is None:
            return False
        if isinstance(field_value, str):
            return str(value) in field_value
        if isinstance(field_value, (list, tuple)):
            return value in field_value
        return False

    if operator == "not_contains":
        if field_value is None:
            return True
        if isinstance(field_value, str):
            return str(value) not in field_value
        if isinstance(field_value, (list, tuple)):
            return value not in field_value
        return True

    return False


def get_nested_value(data: dict, path: str) -> Any:
    """
    Support dot-notation path access: 'servicio.enabled'
    Soporte para acceso con notacion de punto: 'servicio.enabled'

    Args:
        data: Dictionary to access
        path: Dot-separated path string

    Returns:
        Value at the path, or None if not found
    """
    if not path or not data:
        return None

    keys = path.split(".")
    value = data

    for key in keys:
        if isinstance(value, dict):
            value = value.get(key)
        elif isinstance(value, list):
            try:
                index = int(key)
                value = value[index] if 0 <= index < len(value) else None
            except ValueError:
                return None
        else:
            return None

        if value is None:
            return None

    return value


def set_nested_value(data: dict, path: str, value: Any) -> None:
    """
    Set a value at a nested path
    Establecer un valor en una ruta anidada

    Args:
        data: Dictionary to modify
        path: Dot-separated path string
        value: Value to set
    """
    if not path:
        return

    keys = path.split(".")
    current = data

    for key in keys[:-1]:
        if key not in current:
            current[key] = {}
        current = current[key]

    current[keys[-1]] = value


def evaluate_simple_condition(condition_str: str, data: dict) -> bool:
    """
    Evaluate a simple condition string like "tipo_informe == 'completo'"
    Evaluar una cadena de condicion simple como "tipo_informe == 'completo'"

    This is a simplified parser for basic conditions used in field visibility.

    Args:
        condition_str: Simple condition string
        data: Data dictionary

    Returns:
        bool: Result of evaluation
    """
    if not condition_str:
        return True

    # Handle simple equality: field == 'value'
    if "==" in condition_str:
        parts = condition_str.split("==")
        if len(parts) == 2:
            field = parts[0].strip()
            value = parts[1].strip().strip("'\"")
            return get_nested_value(data, field) == value

    # Handle inequality: field != 'value'
    if "!=" in condition_str:
        parts = condition_str.split("!=")
        if len(parts) == 2:
            field = parts[0].strip()
            value = parts[1].strip().strip("'\"")
            return get_nested_value(data, field) != value

    # Handle boolean field: just the field name
    field_value = get_nested_value(data, condition_str)
    if isinstance(field_value, bool):
        return field_value
    if isinstance(field_value, str):
        return field_value.lower() in ('true', 'si', 'yes', '1')

    return bool(field_value)
