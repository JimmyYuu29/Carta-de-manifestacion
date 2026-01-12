# Prompt: 将简单Streamlit应用重构为模块化文档生成平台

## 使用说明

将此prompt复制并提供给Claude Code，同时附上你的：
1. 当前的Streamlit脚本文件
2. Word模板文件

---

## Prompt内容

```
我有一个简单的Streamlit文档生成应用，目前所有代码都在一个脚本中运行（前端、计算逻辑、后端）。我需要你帮我将它重构为模块化的企业级架构。

请先分析我提供的文件，然后按照以下参考架构进行重构。

---

## 目标架构概览

### 项目结构
```
project_root/
├── config/                          # 配置与模板
│   ├── yamls/
│   │   └── {plugin_id}/            # 插件配置包
│   │       ├── manifest.yaml       # 元数据（版本、描述、敏感字段）
│   │       ├── config.yaml         # 运行时配置 & UI定义
│   │       ├── fields.yaml         # 输入字段定义（类型、验证、条件可见性）
│   │       ├── texts.yaml          # 固定文本块库
│   │       ├── tables.yaml         # 表格定义
│   │       ├── logic.yaml          # 条件规则（DSL表达式）
│   │       ├── decision_map.yaml   # 决策与规则映射
│   │       ├── derived.yaml        # 派生字段计算公式
│   │       └── formatting.yaml     # 格式化规则（日期、货币、颜色）
│   └── templates/
│       └── {plugin_id}/
│           └── template.docx       # Word模板
│
├── modules/                         # 核心引擎
│   ├── __init__.py
│   ├── plugin_loader.py            # 插件加载器（带LRU缓存）
│   ├── contract_models.py          # 动态Pydantic模型生成
│   ├── contract_validator.py       # 多层数据验证
│   ├── dsl_evaluator.py            # 安全的条件表达式评估器
│   ├── rule_engine.py              # 规则引擎与可见性计算
│   ├── context_builder.py          # 模板上下文构建（派生字段、格式化）
│   ├── renderer_docx.py            # Word渲染器（docxtpl）
│   ├── generate.py                 # 统一生成入口
│   └── audit_logger.py             # 审计与追踪
│
├── ui/                              # 用户界面
│   └── streamlit_app/
│       ├── app.py                  # 主应用入口
│       ├── state_store.py          # 会话状态管理
│       ├── form_renderer.py        # 表单渲染器
│       └── components.py           # 可复用UI组件
│
├── scripts/                         # CLI工具
│   ├── run_validate.py             # 插件验证脚本
│   └── run_generate.py             # 文档生成脚本
│
├── tests/                           # 测试套件
│   └── test_*.py
│
├── output/                          # 生成的文档
├── logs/                            # 审计日志
└── requirements.txt
```

---

## 核心模块设计规范

### 1. Plugin加载器 (plugin_loader.py)

```python
from functools import lru_cache
from pathlib import Path
import yaml

class PluginPack:
    """懒加载的配置容器"""
    def __init__(self, plugin_id: str):
        self.plugin_id = plugin_id
        self.base_path = Path(f"config/yamls/{plugin_id}")
        self._cache = {}

    @property
    def manifest(self) -> dict:
        return self._load("manifest.yaml")

    @property
    def fields(self) -> dict:
        return self._load("fields.yaml")

    # ... 其他配置属性 ...

    def _load(self, filename: str) -> dict:
        if filename not in self._cache:
            self._cache[filename] = load_yaml_file(self.base_path / filename)
        return self._cache[filename]

@lru_cache(maxsize=16)
def load_yaml_file(path: Path) -> dict:
    """缓存的YAML加载"""
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def load_plugin(plugin_id: str) -> PluginPack:
    return PluginPack(plugin_id)
```

### 2. 动态数据模型 (contract_models.py)

```python
from pydantic import BaseModel, create_model
from typing import Any, Optional
from decimal import Decimal
from datetime import date

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

def build_pydantic_model(plugin: PluginPack) -> type[BaseModel]:
    """从YAML字段定义动态生成Pydantic模型"""
    fields = plugin.fields.get("fields", {})
    field_definitions = {}

    for name, spec in fields.items():
        field_type = TYPE_MAP.get(spec.get("type", "text"), str)
        required = spec.get("required", False)
        default = spec.get("default", None)

        if required:
            field_definitions[name] = (field_type, ...)
        else:
            field_definitions[name] = (Optional[field_type], default)

    return create_model("DynamicInputModel", **field_definitions)
```

### 3. 安全DSL评估器 (dsl_evaluator.py)

```python
from typing import Any

ALLOWED_OPERATORS = frozenset({
    "and", "or", "not",
    "equals", "not_equals", "gt", "gte", "lt", "lte",
    "in", "not_in",
    "exists", "not_exists", "is_empty", "not_empty",
    "contains", "not_contains",
})

MAX_NESTING_DEPTH = 3

def evaluate_condition(condition: dict, data: dict, depth: int = 0) -> bool:
    """安全的条件评估（无eval/exec）"""
    if depth > MAX_NESTING_DEPTH:
        raise ValueError("条件嵌套过深")

    operator = condition.get("operator")
    if operator not in ALLOWED_OPERATORS:
        raise ValueError(f"不允许的操作符: {operator}")

    # 逻辑操作符
    if operator == "and":
        return all(evaluate_condition(c, data, depth+1)
                   for c in condition.get("conditions", []))
    if operator == "or":
        return any(evaluate_condition(c, data, depth+1)
                   for c in condition.get("conditions", []))
    if operator == "not":
        return not evaluate_condition(condition.get("condition"), data, depth+1)

    # 比较操作符
    field = condition.get("field")
    value = condition.get("value")
    field_value = get_nested_value(data, field)

    if operator == "equals":
        return field_value == value
    if operator == "not_equals":
        return field_value != value
    if operator == "gt":
        return field_value > value
    # ... 其他操作符 ...

    return False

def get_nested_value(data: dict, path: str) -> Any:
    """支持点号路径访问: 'servicio.enabled'"""
    keys = path.split(".")
    value = data
    for key in keys:
        if isinstance(value, dict):
            value = value.get(key)
        else:
            return None
    return value
```

### 4. 规则引擎 (rule_engine.py)

```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class RuleHit:
    rule_id: str
    condition_met: bool
    action_type: str
    affected_elements: list[str]

@dataclass
class EvaluationTrace:
    decision_id: str
    rule_hits: list[RuleHit]
    outcome: str

class RuleEngine:
    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self.logic = plugin.logic
        self.decision_map = plugin.decision_map

    def evaluate_all_rules(self, data: dict) -> tuple[dict, list[EvaluationTrace]]:
        """评估所有规则，返回可见性映射和追踪"""
        visibility_map = {}
        traces = []

        for decision_id, decision in self.decision_map.items():
            rule_hits = []
            for rule_id in decision.get("rules", []):
                rule = self.logic.get("rules", {}).get(rule_id)
                if rule:
                    hit = self._evaluate_rule(rule, data)
                    rule_hits.append(hit)
                    if hit.condition_met:
                        for element in hit.affected_elements:
                            visibility_map[element] = True

            traces.append(EvaluationTrace(
                decision_id=decision_id,
                rule_hits=rule_hits,
                outcome="evaluated"
            ))

        return visibility_map, traces

    def _evaluate_rule(self, rule: dict, data: dict) -> RuleHit:
        condition = rule.get("condition", {})
        action = rule.get("action", {})

        met = evaluate_condition(condition, data)

        return RuleHit(
            rule_id=rule.get("rule_id"),
            condition_met=met,
            action_type=action.get("type"),
            affected_elements=action.get("elements", [])
        )
```

### 5. 上下文构建器 (context_builder.py)

```python
from datetime import date
from decimal import Decimal

class ContextBuilder:
    def __init__(self, plugin: PluginPack):
        self.plugin = plugin

    def build_context(self, data: dict) -> dict:
        """构建完整的模板上下文"""
        context = dict(data)

        # 1. 计算派生字段
        context = self._calculate_derived_fields(context)

        # 2. 应用格式化
        context = self._apply_formatting(context)

        # 3. 添加文本块库
        context["texts"] = self.plugin.texts

        # 4. 清理值
        context = self._sanitize_values(context)

        return context

    def _calculate_derived_fields(self, context: dict) -> dict:
        """计算派生字段"""
        derived = self.plugin.derived.get("derived_fields", {})

        for field_name, spec in derived.items():
            formula = spec.get("formula")
            deps = spec.get("dependencies", [])

            # 检查依赖是否存在
            if all(context.get(d) is not None for d in deps):
                context[field_name] = self._evaluate_formula(formula, context)

        return context

    def _apply_formatting(self, context: dict) -> dict:
        """应用格式化规则"""
        formatting = self.plugin.formatting

        for field, fmt_spec in formatting.get("fields", {}).items():
            if field in context and context[field] is not None:
                fmt_type = fmt_spec.get("type")
                if fmt_type == "date":
                    context[f"{field}_formatted"] = format_spanish_date(context[field])
                elif fmt_type == "currency":
                    context[f"{field}_formatted"] = format_currency_eur(context[field])
                elif fmt_type == "percentage":
                    context[f"{field}_formatted"] = format_percentage(context[field])

        return context

# 格式化函数
def format_spanish_date(d: date) -> str:
    """格式化为西班牙日期: 31 de diciembre de 2025"""
    months = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
              "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    return f"{d.day} de {months[d.month - 1]} de {d.year}"

def format_currency_eur(value: Decimal) -> str:
    """格式化为欧元: 1.500.000 €"""
    return f"{value:,.0f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def format_percentage(value: float) -> str:
    """格式化为百分比: 15,00 %"""
    return f"{value:.2f} %".replace(".", ",")
```

### 6. Word渲染器 (renderer_docx.py)

```python
from pathlib import Path
from docxtpl import DocxTemplate
from docx.shared import RGBColor

class DocxRenderer:
    def __init__(self, plugin: PluginPack):
        self.plugin = plugin
        self.template_path = Path(plugin.manifest.get("template", {}).get("path"))
        self.context_builder = ContextBuilder(plugin)
        self.rule_engine = RuleEngine(plugin)

    def render(self, data: dict, output_path: Path) -> tuple[Path, list]:
        """渲染Word文档"""
        # 1. 加载模板
        doc = DocxTemplate(self.template_path)

        # 2. 构建上下文
        context = self.context_builder.build_context(data)

        # 3. 评估规则
        visibility_map, traces = self.rule_engine.evaluate_all_rules(data)
        context["visibility"] = visibility_map

        # 4. 渲染
        doc.render(context)

        # 5. 后处理
        self._post_process(doc)

        # 6. 保存
        doc.save(output_path)

        return output_path, traces

    def _post_process(self, doc):
        """后处理：单元格着色、清理空段落等"""
        # 应用表格单元格颜色
        self._apply_cell_colors(doc)
        # 移除空段落
        self._remove_empty_paragraphs(doc)

    def _apply_cell_colors(self, doc):
        """根据条件应用单元格背景色"""
        colors = self.plugin.formatting.get("colors", {})
        for table in doc.docx.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip().lower()
                    if text in colors:
                        self._set_cell_color(cell, colors[text])

    def _set_cell_color(self, cell, hex_color: str):
        """设置单元格背景色"""
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color.replace("#", ""))
        tcPr.append(shd)
```

### 7. 统一生成入口 (generate.py)

```python
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import uuid
import time

@dataclass
class GenerationResult:
    success: bool
    output_path: Optional[Path]
    trace_id: str
    validation_errors: list[str]
    evaluation_traces: list
    error: Optional[str]
    duration_ms: int

def generate(
    plugin_id: str,
    data: dict,
    output_dir: Path = Path("output"),
    should_validate: bool = True
) -> GenerationResult:
    """文档生成的统一入口"""
    start_time = time.time()
    trace_id = str(uuid.uuid4())

    try:
        # 1. 加载插件
        plugin = load_plugin(plugin_id)

        # 2. 预处理输入
        data = preprocess_input(data, plugin)

        # 3. 验证（可选）
        validation_errors = []
        if should_validate:
            validation_result = validate_input(plugin, data)
            if not validation_result.is_valid:
                return GenerationResult(
                    success=False,
                    output_path=None,
                    trace_id=trace_id,
                    validation_errors=validation_result.errors,
                    evaluation_traces=[],
                    error="验证失败",
                    duration_ms=int((time.time() - start_time) * 1000)
                )

        # 4. 渲染
        renderer = DocxRenderer(plugin)
        output_path = output_dir / f"{plugin_id}_{trace_id[:8]}.docx"
        output_path, traces = renderer.render(data, output_path)

        return GenerationResult(
            success=True,
            output_path=output_path,
            trace_id=trace_id,
            validation_errors=[],
            evaluation_traces=traces,
            error=None,
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
    """预处理输入数据：类型转换"""
    fields = plugin.fields.get("fields", {})
    result = dict(data)

    for field_name, spec in fields.items():
        if field_name in result:
            value = result[field_name]
            field_type = spec.get("type")

            # 日期字符串转换
            if field_type == "date" and isinstance(value, str):
                result[field_name] = date.fromisoformat(value)

            # 数字转换
            elif field_type in ("int", "currency") and isinstance(value, str):
                result[field_name] = int(value.replace(",", "").replace(".", ""))

    return result
```

### 8. Streamlit状态管理 (state_store.py)

```python
import streamlit as st
from typing import Any, Optional
import uuid

def init_session_state(plugin_id: str):
    """初始化会话状态"""
    if "plugin_id" not in st.session_state:
        st.session_state.plugin_id = plugin_id
    if "form_data" not in st.session_state:
        st.session_state.form_data = {}
    if "list_items" not in st.session_state:
        st.session_state.list_items = {}
    if "generation_result" not in st.session_state:
        st.session_state.generation_result = None

def get_stable_key(field_name: str, index: Optional[int] = None, sub_field: Optional[str] = None) -> str:
    """生成稳定的widget键"""
    key = f"field_{field_name}"
    if index is not None:
        key += f"_{index}"
    if sub_field:
        key += f"_{sub_field}"
    return key

def get_field_value(field_name: str, default: Any = None) -> Any:
    """获取字段值"""
    return st.session_state.form_data.get(field_name, default)

def set_field_value(field_name: str, value: Any):
    """设置字段值"""
    st.session_state.form_data[field_name] = value

def get_list_items(field_name: str) -> list:
    """获取列表项"""
    if field_name not in st.session_state.list_items:
        st.session_state.list_items[field_name] = []
    return st.session_state.list_items[field_name]

def add_list_item(field_name: str, item: dict):
    """添加列表项"""
    items = get_list_items(field_name)
    item["_id"] = str(uuid.uuid4())
    items.append(item)

def remove_list_item(field_name: str, item_id: str):
    """删除列表项"""
    items = get_list_items(field_name)
    st.session_state.list_items[field_name] = [
        item for item in items if item.get("_id") != item_id
    ]

def get_all_form_data() -> dict:
    """收集所有表单数据"""
    data = dict(st.session_state.form_data)

    # 添加列表数据
    for field_name, items in st.session_state.list_items.items():
        # 清理内部ID
        clean_items = [{k: v for k, v in item.items() if not k.startswith("_")}
                       for item in items]
        data[field_name] = clean_items

    return data

def clear_form_data():
    """清空表单数据"""
    st.session_state.form_data = {}
    st.session_state.list_items = {}
    st.session_state.generation_result = None
```

---

## YAML配置文件示例

### fields.yaml
```yaml
fields:
  empresa_nombre:
    type: text
    required: true
    label: "Nombre de la empresa"
    validation:
      max_length: 200

  fecha_fin_fiscal:
    type: date
    required: true
    label: "Fecha de cierre del ejercicio"

  cifra_negocios:
    type: currency
    required: true
    label: "Cifra de negocios"
    validation:
      min: 0

  tipo_informe:
    type: enum
    required: true
    label: "Tipo de informe"
    values:
      - value: "completo"
        label: "Informe completo"
      - value: "simplificado"
        label: "Informe simplificado"

  incluir_anexos:
    type: bool
    required: false
    label: "Incluir anexos"
    default: false

  servicios:
    type: list
    label: "Servicios vinculados"
    item_schema:
      nombre:
        type: text
        required: true
      importe:
        type: currency
        required: true
```

### logic.yaml
```yaml
rules:
  r001:
    rule_id: r001
    name: "Incluir sección de anexos"
    condition:
      operator: equals
      field: incluir_anexos
      value: true
    action:
      type: include_block
      elements: ["bloque_anexos"]

  r002:
    rule_id: r002
    name: "Informe simplificado"
    condition:
      operator: and
      conditions:
        - operator: equals
          field: tipo_informe
          value: "simplificado"
        - operator: lt
          field: cifra_negocios
          value: 1000000
    action:
      type: include_text
      elements: ["texto_simplificado"]
```

### derived.yaml
```yaml
derived_fields:
  anyo_ejercicio:
    type: int
    formula: "extract_year(fecha_fin_fiscal)"
    dependencies: [fecha_fin_fiscal]

  total_servicios:
    type: currency
    formula: "sum(servicios.importe)"
    dependencies: [servicios]

  porcentaje_servicios:
    type: percentage
    formula: "(total_servicios / cifra_negocios) * 100"
    dependencies: [total_servicios, cifra_negocios]
```

### formatting.yaml
```yaml
fields:
  fecha_fin_fiscal:
    type: date
    format: "spanish"  # "31 de diciembre de 2025"

  cifra_negocios:
    type: currency
    format: "eur"  # "1.500.000 €"

colors:
  si: "#90EE90"     # 绿色
  no: "#FFB6C1"     # 红色
  parcial: "#FFFFE0" # 黄色
```

---

## 重构步骤

请按以下顺序执行重构：

### 阶段1：分析现有代码
1. 阅读我提供的Streamlit脚本
2. 识别所有输入字段及其类型
3. 识别所有计算逻辑和派生字段
4. 识别所有条件逻辑（if语句）
5. 分析Word模板中的占位符

### 阶段2：创建配置文件
1. 创建目录结构
2. 从现有代码提取字段定义 → fields.yaml
3. 从条件逻辑提取规则 → logic.yaml
4. 从计算逻辑提取派生字段 → derived.yaml
5. 创建格式化规则 → formatting.yaml
6. 创建manifest.yaml

### 阶段3：创建核心模块
1. plugin_loader.py - 配置加载
2. dsl_evaluator.py - 条件评估
3. rule_engine.py - 规则引擎
4. context_builder.py - 上下文构建
5. renderer_docx.py - 文档渲染
6. generate.py - 生成入口

### 阶段4：创建UI层
1. state_store.py - 状态管理
2. components.py - UI组件
3. form_renderer.py - 表单渲染
4. app.py - 主应用

### 阶段5：测试与验证
1. 创建测试数据
2. 验证生成的文档与原版一致
3. 测试边界情况

---

## 性能优化要点

1. **LRU缓存**: 对YAML文件加载使用 `@lru_cache`
2. **懒加载**: 配置文件仅在访问时加载
3. **预计算**: 派生字段分组计算，避免重复
4. **并行处理**: 独立规则可并行评估

---

## 安全要点

1. **DSL白名单**: 仅允许预定义操作符，无eval/exec
2. **深度限制**: 条件嵌套最大深度为3
3. **输入验证**: 使用Pydantic进行类型验证
4. **敏感字段**: 在manifest中标记，审计日志中掩码

---

现在请分析我提供的文件，开始重构工作。
```

---

## 附录：关键代码模式速查

### A. Streamlit Widget状态保持

```python
# 问题：Streamlit重新运行时widget值丢失
# 解决：使用session_state作为单一数据源

def render_text_input(field_name: str, label: str):
    key = get_stable_key(field_name)

    # 从session_state获取值
    current_value = get_field_value(field_name, "")

    # 创建widget
    new_value = st.text_input(
        label,
        value=current_value,
        key=key,
        on_change=lambda: set_field_value(field_name, st.session_state[key])
    )
```

### B. 条件字段可见性

```python
# fields.yaml
campos_adicionales:
  type: text
  condition: "tipo_informe == 'completo'"

# form_renderer.py
def should_show_field(field_spec: dict, data: dict) -> bool:
    condition = field_spec.get("condition")
    if not condition:
        return True
    return evaluate_simple_condition(condition, data)
```

### C. 表格单元格着色

```python
def apply_cell_color(cell, color_map: dict):
    text = cell.text.strip().lower()
    if text in color_map:
        hex_color = color_map[text]
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color.lstrip('#'))
        tcPr.append(shd)
```

### D. 西班牙日期格式化

```python
SPANISH_MONTHS = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

def format_spanish_date(d: date) -> str:
    return f"{d.day} de {SPANISH_MONTHS[d.month - 1]} de {d.year}"
```

### E. 欧元货币格式化

```python
def format_currency_eur(value) -> str:
    # 1500000 → "1.500.000 €"
    return f"{int(value):,} €".replace(",", "X").replace(".", ",").replace("X", ".")
```
