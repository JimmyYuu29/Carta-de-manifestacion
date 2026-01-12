"""
Plugin Loader - Configuration loader with LRU cache
Cargador de plugins con cache LRU
"""

from functools import lru_cache
from pathlib import Path
from typing import Dict, Any, Optional
import yaml


class PluginPack:
    """Lazy-loading configuration container / Contenedor de configuracion con carga perezosa"""

    def __init__(self, plugin_id: str, base_path: Optional[Path] = None):
        self.plugin_id = plugin_id
        if base_path:
            self.base_path = base_path
        else:
            self.base_path = Path(__file__).parent.parent / "config" / "yamls" / plugin_id
        self._cache: Dict[str, dict] = {}

    @property
    def manifest(self) -> dict:
        """Plugin metadata"""
        return self._load("manifest.yaml")

    @property
    def config(self) -> dict:
        """Runtime configuration & UI definition"""
        return self._load("config.yaml")

    @property
    def fields(self) -> dict:
        """Input field definitions"""
        return self._load("fields.yaml")

    @property
    def texts(self) -> dict:
        """Fixed text blocks library"""
        return self._load("texts.yaml")

    @property
    def tables(self) -> dict:
        """Table definitions"""
        return self._load("tables.yaml")

    @property
    def logic(self) -> dict:
        """Conditional rules"""
        return self._load("logic.yaml")

    @property
    def decision_map(self) -> dict:
        """Decision and rule mapping"""
        return self._load("decision_map.yaml")

    @property
    def derived(self) -> dict:
        """Derived field calculation formulas"""
        return self._load("derived.yaml")

    @property
    def formatting(self) -> dict:
        """Formatting rules"""
        return self._load("formatting.yaml")

    def _load(self, filename: str) -> dict:
        """Load a YAML file with caching"""
        if filename not in self._cache:
            file_path = self.base_path / filename
            self._cache[filename] = load_yaml_file(file_path)
        return self._cache[filename]

    def get_template_path(self) -> Path:
        """Get the path to the Word template"""
        template_info = self.manifest.get("template", {})
        template_path = template_info.get("path", "")
        if template_path:
            # Path relative to project root
            return Path(__file__).parent.parent / template_path
        # Default path
        return Path(__file__).parent.parent / "config" / "templates" / self.plugin_id / "template.docx"

    def get_oficinas(self) -> dict:
        """Get office configurations"""
        return self.config.get("oficinas", {})

    def get_sections(self) -> list:
        """Get UI sections in order"""
        sections = self.config.get("sections", [])
        return sorted(sections, key=lambda x: x.get("order", 0))

    def get_field_spec(self, field_name: str) -> Optional[dict]:
        """Get specification for a specific field"""
        return self.fields.get("fields", {}).get(field_name)

    def get_conditional_dependencies(self) -> dict:
        """Get mapping of conditional fields to their dependent fields"""
        return self.decision_map.get("conditional_dependencies", {})

    def clear_cache(self):
        """Clear the internal cache"""
        self._cache.clear()


@lru_cache(maxsize=32)
def load_yaml_file(path: Path) -> dict:
    """Cached YAML file loading / Carga de archivo YAML con cache"""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    except FileNotFoundError:
        return {}
    except yaml.YAMLError as e:
        raise ValueError(f"Error parsing YAML file {path}: {e}")


def load_plugin(plugin_id: str) -> PluginPack:
    """Load a plugin by ID / Cargar un plugin por ID"""
    return PluginPack(plugin_id)


def list_available_plugins() -> list:
    """List all available plugins / Listar todos los plugins disponibles"""
    plugins_dir = Path(__file__).parent.parent / "config" / "yamls"
    if not plugins_dir.exists():
        return []
    return [d.name for d in plugins_dir.iterdir() if d.is_dir()]
