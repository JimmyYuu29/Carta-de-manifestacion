"""
Rule Engine - Rule evaluation and visibility calculation
Motor de reglas y calculo de visibilidad
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any

from .dsl_evaluator import evaluate_condition
from .plugin_loader import PluginPack


@dataclass
class RuleHit:
    """Result of evaluating a single rule / Resultado de evaluar una regla"""
    rule_id: str
    rule_name: str
    condition_met: bool
    action_type: str
    affected_elements: List[str]
    text_key: Optional[str] = None


@dataclass
class EvaluationTrace:
    """Trace of a decision evaluation / Traza de una evaluacion de decision"""
    decision_id: str
    description: str
    rule_hits: List[RuleHit]
    outcome: str


class RuleEngine:
    """
    Rule engine for evaluating conditions and computing visibility
    Motor de reglas para evaluar condiciones y calcular visibilidad
    """

    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self.logic = plugin.logic
        self.decision_map = plugin.decision_map

    def evaluate_all_rules(self, data: dict) -> Tuple[Dict[str, Any], List[EvaluationTrace]]:
        """
        Evaluate all rules and return visibility map and traces
        Evaluar todas las reglas y devolver mapa de visibilidad y trazas

        Args:
            data: Input data dictionary

        Returns:
            Tuple of (visibility_map, traces)
        """
        visibility_map: Dict[str, Any] = {}
        traces: List[EvaluationTrace] = []

        decisions = self.decision_map.get("decisions", {})
        rules = self.logic.get("rules", {})

        for decision_id, decision in decisions.items():
            rule_hits: List[RuleHit] = []
            is_exclusive = decision.get("exclusive", False)
            exclusive_hit = False

            for rule_id in decision.get("rules", []):
                rule = rules.get(rule_id)
                if not rule:
                    continue

                # Skip if exclusive decision already has a hit
                if is_exclusive and exclusive_hit:
                    continue

                hit = self._evaluate_rule(rule, data)
                rule_hits.append(hit)

                if hit.condition_met:
                    if is_exclusive:
                        exclusive_hit = True

                    # Process action
                    self._process_action(hit, visibility_map)

            # Apply default if no rules matched in exclusive decision
            if is_exclusive and not exclusive_hit:
                default_key = decision.get("default")
                if default_key:
                    visibility_map[f"text_{decision_id}"] = default_key

            traces.append(EvaluationTrace(
                decision_id=decision_id,
                description=decision.get("description", ""),
                rule_hits=rule_hits,
                outcome="exclusive_hit" if exclusive_hit else "evaluated"
            ))

        return visibility_map, traces

    def _evaluate_rule(self, rule: dict, data: dict) -> RuleHit:
        """Evaluate a single rule / Evaluar una regla individual"""
        condition = rule.get("condition", {})
        action = rule.get("action", {})

        met = evaluate_condition(condition, data)

        return RuleHit(
            rule_id=rule.get("rule_id", "unknown"),
            rule_name=rule.get("name", ""),
            condition_met=met,
            action_type=action.get("type", ""),
            affected_elements=action.get("elements", []),
            text_key=action.get("text_key")
        )

    def _process_action(self, hit: RuleHit, visibility_map: Dict[str, Any]) -> None:
        """Process a rule action / Procesar una accion de regla"""
        if hit.action_type == "include_block":
            for element in hit.affected_elements:
                visibility_map[element] = True

        elif hit.action_type == "exclude_block":
            for element in hit.affected_elements:
                visibility_map[element] = False

        elif hit.action_type == "set_text":
            for element in hit.affected_elements:
                visibility_map[element] = hit.text_key

        elif hit.action_type == "include_text":
            for element in hit.affected_elements:
                visibility_map[element] = True

    def get_field_visibility(self, data: dict) -> Dict[str, bool]:
        """
        Calculate which fields should be visible based on conditions
        Calcular que campos deben ser visibles basado en condiciones

        Args:
            data: Current form data

        Returns:
            Dictionary mapping field names to visibility (True/False)
        """
        visibility = {}
        fields = self.plugin.fields.get("fields", {})

        for field_name, field_spec in fields.items():
            condition = field_spec.get("condition")
            if condition:
                visibility[field_name] = evaluate_condition(condition, data)
            else:
                visibility[field_name] = True

        return visibility

    def get_required_fields(self, data: dict) -> List[str]:
        """
        Get list of currently required fields based on conditions
        Obtener lista de campos requeridos actuales basado en condiciones

        Args:
            data: Current form data

        Returns:
            List of required field names
        """
        required = []
        fields = self.plugin.fields.get("fields", {})
        visibility = self.get_field_visibility(data)

        for field_name, field_spec in fields.items():
            if not visibility.get(field_name, True):
                continue
            if field_spec.get("required", False):
                required.append(field_name)

        return required

    def compute_conditional_values(self, data: dict) -> Dict[str, str]:
        """
        Convert boolean values to 'si'/'no' for template compatibility
        Convertir valores booleanos a 'si'/'no' para compatibilidad con plantilla

        Args:
            data: Input data dictionary

        Returns:
            Dictionary with 'si'/'no' string values
        """
        result = {}
        bool_fields = [
            'comision', 'junta', 'comite', 'incorreccion', 'limitacion_alcance',
            'dudas', 'rent', 'A_coste', 'experto', 'unidad_decision',
            'activo_impuesto', 'operacion_fiscal', 'compromiso', 'gestion'
        ]

        for field_name in bool_fields:
            value = data.get(field_name)
            if isinstance(value, bool):
                result[field_name] = 'si' if value else 'no'
            elif isinstance(value, str):
                result[field_name] = value.lower() if value.lower() in ('si', 'no') else 'no'
            else:
                result[field_name] = 'no'

        return result
