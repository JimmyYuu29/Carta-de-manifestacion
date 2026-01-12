# Modules - Core engine components
# Componentes del motor central

from .plugin_loader import PluginPack, load_plugin, load_yaml_file
from .dsl_evaluator import evaluate_condition, get_nested_value
from .rule_engine import RuleEngine, RuleHit, EvaluationTrace
from .context_builder import ContextBuilder, format_spanish_date, format_currency_eur
from .renderer_docx import DocxRenderer
from .generate import generate, GenerationResult, preprocess_input
from .contract_validator import validate_input, ValidationResult

__all__ = [
    'PluginPack',
    'load_plugin',
    'load_yaml_file',
    'evaluate_condition',
    'get_nested_value',
    'RuleEngine',
    'RuleHit',
    'EvaluationTrace',
    'ContextBuilder',
    'format_spanish_date',
    'format_currency_eur',
    'DocxRenderer',
    'generate',
    'GenerationResult',
    'preprocess_input',
    'validate_input',
    'ValidationResult',
]
