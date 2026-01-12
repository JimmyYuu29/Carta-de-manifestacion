"""
Tests for DSL evaluator
Tests para el evaluador DSL
"""

import pytest
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from modules.dsl_evaluator import evaluate_condition, get_nested_value, DSLEvaluationError


class TestGetNestedValue:
    """Tests for get_nested_value function"""

    def test_simple_key(self):
        data = {"name": "test"}
        assert get_nested_value(data, "name") == "test"

    def test_nested_key(self):
        data = {"user": {"name": "test"}}
        assert get_nested_value(data, "user.name") == "test"

    def test_missing_key(self):
        data = {"name": "test"}
        assert get_nested_value(data, "missing") is None

    def test_deep_nesting(self):
        data = {"a": {"b": {"c": "value"}}}
        assert get_nested_value(data, "a.b.c") == "value"

    def test_empty_data(self):
        assert get_nested_value({}, "key") is None

    def test_empty_path(self):
        assert get_nested_value({"key": "value"}, "") is None


class TestEvaluateCondition:
    """Tests for evaluate_condition function"""

    def test_equals(self):
        condition = {"operator": "equals", "field": "status", "value": "active"}
        data = {"status": "active"}
        assert evaluate_condition(condition, data) is True

    def test_equals_false(self):
        condition = {"operator": "equals", "field": "status", "value": "active"}
        data = {"status": "inactive"}
        assert evaluate_condition(condition, data) is False

    def test_not_equals(self):
        condition = {"operator": "not_equals", "field": "status", "value": "inactive"}
        data = {"status": "active"}
        assert evaluate_condition(condition, data) is True

    def test_gt(self):
        condition = {"operator": "gt", "field": "count", "value": 5}
        data = {"count": 10}
        assert evaluate_condition(condition, data) is True

    def test_lt(self):
        condition = {"operator": "lt", "field": "count", "value": 10}
        data = {"count": 5}
        assert evaluate_condition(condition, data) is True

    def test_and_operator(self):
        condition = {
            "operator": "and",
            "conditions": [
                {"operator": "equals", "field": "a", "value": 1},
                {"operator": "equals", "field": "b", "value": 2}
            ]
        }
        data = {"a": 1, "b": 2}
        assert evaluate_condition(condition, data) is True

    def test_or_operator(self):
        condition = {
            "operator": "or",
            "conditions": [
                {"operator": "equals", "field": "a", "value": 1},
                {"operator": "equals", "field": "b", "value": 3}
            ]
        }
        data = {"a": 1, "b": 2}
        assert evaluate_condition(condition, data) is True

    def test_not_operator(self):
        condition = {
            "operator": "not",
            "condition": {"operator": "equals", "field": "status", "value": "inactive"}
        }
        data = {"status": "active"}
        assert evaluate_condition(condition, data) is True

    def test_exists(self):
        condition = {"operator": "exists", "field": "name"}
        data = {"name": "test"}
        assert evaluate_condition(condition, data) is True

    def test_not_exists(self):
        condition = {"operator": "not_exists", "field": "missing"}
        data = {"name": "test"}
        assert evaluate_condition(condition, data) is True

    def test_is_empty(self):
        condition = {"operator": "is_empty", "field": "list"}
        data = {"list": []}
        assert evaluate_condition(condition, data) is True

    def test_not_empty(self):
        condition = {"operator": "not_empty", "field": "list"}
        data = {"list": [1, 2, 3]}
        assert evaluate_condition(condition, data) is True

    def test_contains(self):
        condition = {"operator": "contains", "field": "text", "value": "hello"}
        data = {"text": "hello world"}
        assert evaluate_condition(condition, data) is True

    def test_invalid_operator(self):
        condition = {"operator": "invalid_op", "field": "a", "value": 1}
        with pytest.raises(DSLEvaluationError):
            evaluate_condition(condition, {})

    def test_max_nesting_depth(self):
        # Create deeply nested condition
        condition = {"operator": "not", "condition": None}
        current = condition
        for _ in range(10):
            current["condition"] = {"operator": "not", "condition": None}
            current = current["condition"]
        current["condition"] = {"operator": "equals", "field": "a", "value": 1}

        with pytest.raises(DSLEvaluationError):
            evaluate_condition(condition, {"a": 1})

    def test_empty_condition(self):
        assert evaluate_condition({}, {}) is True
        assert evaluate_condition(None, {}) is True

    def test_boolean_normalization(self):
        condition = {"operator": "equals", "field": "active", "value": True}
        data = {"active": "si"}
        assert evaluate_condition(condition, data) is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
