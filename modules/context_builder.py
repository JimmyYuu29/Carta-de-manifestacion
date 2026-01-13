"""
Context Builder - Template context construction with derived fields and formatting
Constructor de contexto de plantilla con campos derivados y formateo
"""

from datetime import date, datetime
from decimal import Decimal
from typing import Any, Dict, List, Optional, Callable
import re

from .plugin_loader import PluginPack


# Spanish month names
SPANISH_MONTHS = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]


def format_spanish_date(d: Any) -> str:
    """
    Format date as Spanish: 31 de diciembre de 2025
    Formatear fecha en espanol: 31 de diciembre de 2025

    Args:
        d: date object or string

    Returns:
        Formatted date string
    """
    if d is None:
        return ""

    if isinstance(d, str):
        # Try to parse common formats
        d = parse_date_string(d)
        if d is None:
            return ""

    if isinstance(d, datetime):
        d = d.date()

    if not isinstance(d, date):
        return str(d)

    return f"{d.day} de {SPANISH_MONTHS[d.month - 1]} de {d.year}"


def format_currency_eur(value: Any) -> str:
    """
    Format as Euro currency: 1.500.000 EUR
    Formatear como moneda Euro: 1.500.000 EUR

    Args:
        value: Numeric value

    Returns:
        Formatted currency string
    """
    if value is None:
        return ""

    try:
        if isinstance(value, str):
            value = float(value.replace(",", ".").replace(" ", ""))
        num_str = f"{int(value):,}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{num_str} EUR"
    except (ValueError, TypeError):
        return str(value)


def format_percentage(value: Any) -> str:
    """
    Format as percentage: 15,00 %
    Formatear como porcentaje: 15,00 %

    Args:
        value: Numeric value

    Returns:
        Formatted percentage string
    """
    if value is None:
        return ""

    try:
        if isinstance(value, str):
            value = float(value.replace(",", "."))
        return f"{value:.2f} %".replace(".", ",")
    except (ValueError, TypeError):
        return str(value)


def parse_date_string(date_string: Any) -> Optional[date]:
    """
    Parse date string to date object
    Parsear string de fecha a objeto date

    Args:
        date_string: Date string in various formats, or date/datetime object

    Returns:
        date object or None
    """
    if not date_string:
        return None

    # If already a date object, return it directly
    if isinstance(date_string, date) and not isinstance(date_string, datetime):
        return date_string

    # If datetime object, extract the date
    if isinstance(date_string, datetime):
        return date_string.date()

    # Must be a string from this point
    if not isinstance(date_string, str):
        return None

    formats = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d de %B de %Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
        "%Y.%m.%d",
        "%m/%d/%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_string, fmt).date()
        except ValueError:
            continue

    return None


class ContextBuilder:
    """
    Builds complete template context with derived fields and formatting
    Construye el contexto completo de la plantilla con campos derivados y formateo
    """

    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self._formula_functions: Dict[str, Callable] = {
            "extract_year": self._extract_year,
            "format_directors_list": self._format_directors_list,
            "bool_to_sino": self._bool_to_sino,
            "sum": self._sum_list,
        }

    def build_context(self, data: dict) -> dict:
        """
        Build complete template context
        Construir contexto completo de la plantilla

        Args:
            data: Input data dictionary

        Returns:
            Complete context dictionary for template rendering
        """
        context = dict(data)

        # 1. Calculate derived fields
        context = self._calculate_derived_fields(context)

        # 2. Apply formatting
        context = self._apply_formatting(context)

        # 3. Add text blocks library
        context["texts"] = self.plugin.texts.get("text_blocks", {})

        # 4. Sanitize values (clean None, etc.)
        context = self._sanitize_values(context)

        # 5. Add formatted lists
        context = self._format_lists(context)

        return context

    def _calculate_derived_fields(self, context: dict) -> dict:
        """Calculate derived fields / Calcular campos derivados"""
        derived = self.plugin.derived.get("derived_fields", {})

        for field_name, spec in derived.items():
            formula = spec.get("formula", "")
            deps = spec.get("dependencies", [])

            # Check all dependencies exist
            if all(context.get(d) is not None for d in deps):
                try:
                    context[field_name] = self._evaluate_formula(formula, context)
                except Exception:
                    context[field_name] = None

        return context

    def _evaluate_formula(self, formula: str, context: dict) -> Any:
        """
        Evaluate a formula safely
        Evaluar una formula de forma segura

        Supported formulas:
        - extract_year(field)
        - format_directors_list(field)
        - bool_to_sino(field)
        - sum(list_field.property)
        - Simple arithmetic with fields
        """
        # Match function call pattern: func_name(arg1, arg2, ...)
        func_match = re.match(r'(\w+)\(([^)]+)\)', formula)
        if func_match:
            func_name = func_match.group(1)
            args_str = func_match.group(2)
            args = [arg.strip() for arg in args_str.split(",")]

            if func_name in self._formula_functions:
                # Get actual values for arguments
                arg_values = []
                for arg in args:
                    if arg in context:
                        arg_values.append(context[arg])
                    else:
                        arg_values.append(arg)

                return self._formula_functions[func_name](*arg_values)

        # Simple arithmetic: field1 - 1, field1 + field2, etc.
        # Handle subtraction
        if " - " in formula:
            parts = formula.split(" - ")
            if len(parts) == 2:
                left = self._get_value(parts[0].strip(), context)
                right = self._get_value(parts[1].strip(), context)
                try:
                    return int(left) - int(right)
                except (ValueError, TypeError):
                    return None

        # Handle addition
        if " + " in formula:
            parts = formula.split(" + ")
            if len(parts) == 2:
                left = self._get_value(parts[0].strip(), context)
                right = self._get_value(parts[1].strip(), context)
                try:
                    return int(left) + int(right)
                except (ValueError, TypeError):
                    return None

        # Handle multiplication/division for percentages
        if " * " in formula or " / " in formula:
            # Basic eval with only numbers - unsafe but controlled
            try:
                # Replace field names with values
                expr = formula
                for key, val in context.items():
                    if key in expr and isinstance(val, (int, float, Decimal)):
                        expr = expr.replace(key, str(val))
                # Only allow safe characters
                if re.match(r'^[\d\s\.\+\-\*\/\(\)]+$', expr):
                    return eval(expr)
            except Exception:
                return None

        return None

    def _get_value(self, key: str, context: dict) -> Any:
        """Get value from context or return as literal"""
        if key in context:
            return context[key]
        try:
            return int(key)
        except ValueError:
            try:
                return float(key)
            except ValueError:
                return key

    def _extract_year(self, date_value: Any) -> Optional[int]:
        """Extract year from date / Extraer anio de fecha"""
        if isinstance(date_value, date):
            return date_value.year
        if isinstance(date_value, datetime):
            return date_value.year
        if isinstance(date_value, str):
            parsed = parse_date_string(date_value)
            if parsed:
                return parsed.year
            # Try to extract year from formatted date
            match = re.search(r'\d{4}', date_value)
            if match:
                return int(match.group())
        return None

    def _format_directors_list(self, directors: Any) -> str:
        """Format directors list with indentation / Formatear lista de directores con indentacion"""
        if not directors:
            return ""

        if isinstance(directors, str):
            return directors

        if isinstance(directors, list):
            indent = "                                  "
            lines = []
            for director in directors:
                if isinstance(director, dict):
                    nombre = director.get("nombre", "")
                    cargo = director.get("cargo", "")
                    if nombre and cargo:
                        lines.append(f"{indent} D. {nombre} - {cargo}")
                elif isinstance(director, str):
                    lines.append(director)
            return "\n".join(lines)

        return str(directors)

    def _bool_to_sino(self, value: Any) -> str:
        """Convert boolean to si/no / Convertir booleano a si/no"""
        if isinstance(value, bool):
            return "si" if value else "no"
        if isinstance(value, str):
            if value.lower() in ("true", "si", "yes", "1"):
                return "si"
            return "no"
        return "no"

    def _sum_list(self, list_field_path: str) -> Any:
        """Sum values from a list field / Sumar valores de un campo de lista"""
        # This would need the actual context to work
        # For now, return 0 as placeholder
        return 0

    def _apply_formatting(self, context: dict) -> dict:
        """Apply formatting rules / Aplicar reglas de formateo"""
        formatting = self.plugin.formatting
        field_formats = formatting.get("fields", {})

        for field, fmt_spec in field_formats.items():
            if field in context and context[field] is not None:
                fmt_type = fmt_spec.get("type")
                value = context[field]

                if fmt_type == "date":
                    # Format date to Spanish and replace the original value
                    if isinstance(value, (date, datetime)):
                        formatted_date = format_spanish_date(value)
                        context[field] = formatted_date
                        context[f"{field}_formatted"] = formatted_date
                    elif isinstance(value, str) and "de" not in value:
                        formatted_date = format_spanish_date(value)
                        context[field] = formatted_date
                        context[f"{field}_formatted"] = formatted_date
                    else:
                        context[f"{field}_formatted"] = value

                elif fmt_type == "currency":
                    context[f"{field}_formatted"] = format_currency_eur(value)

                elif fmt_type == "percentage":
                    context[f"{field}_formatted"] = format_percentage(value)

        return context

    def _sanitize_values(self, context: dict) -> dict:
        """Sanitize values: replace None with empty strings / Sanear valores"""
        for key, value in context.items():
            if value is None:
                context[key] = ""
        return context

    def _format_lists(self, context: dict) -> dict:
        """Format list fields for template / Formatear listas para la plantilla"""
        # Format lista_alto_directores if it's a list of dicts
        if "lista_alto_directores" in context:
            directors = context["lista_alto_directores"]
            if isinstance(directors, list):
                context["lista_alto_directores"] = self._format_directors_list(directors)

        return context

    def get_conditional_values(self, data: dict) -> Dict[str, str]:
        """
        Convert boolean fields to 'si'/'no' for template compatibility
        Convertir campos booleanos a 'si'/'no' para compatibilidad de plantilla
        """
        result = {}
        bool_fields = [
            'comision', 'junta', 'comite', 'incorreccion', 'limitacion_alcance',
            'dudas', 'rent', 'A_coste', 'experto', 'unidad_decision',
            'activo_impuesto', 'operacion_fiscal', 'compromiso', 'gestion'
        ]

        for field_name in bool_fields:
            value = data.get(field_name)
            result[field_name] = self._bool_to_sino(value)

        return result
