"""
Tests for plugin loader
Tests para el cargador de plugins
"""

import pytest
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from modules.plugin_loader import load_plugin, PluginPack, list_available_plugins


def test_load_plugin():
    """Test loading a plugin / Probar carga de plugin"""
    plugin = load_plugin("carta_manifestacion")
    assert plugin is not None
    assert plugin.plugin_id == "carta_manifestacion"


def test_plugin_manifest():
    """Test plugin manifest loading / Probar carga de manifest"""
    plugin = load_plugin("carta_manifestacion")
    manifest = plugin.manifest

    assert manifest is not None
    assert "plugin_id" in manifest
    assert manifest["plugin_id"] == "carta_manifestacion"


def test_plugin_fields():
    """Test plugin fields loading / Probar carga de campos"""
    plugin = load_plugin("carta_manifestacion")
    fields = plugin.fields

    assert fields is not None
    assert "fields" in fields
    assert "Nombre_Cliente" in fields["fields"]


def test_plugin_logic():
    """Test plugin logic loading / Probar carga de logica"""
    plugin = load_plugin("carta_manifestacion")
    logic = plugin.logic

    assert logic is not None
    assert "rules" in logic


def test_list_available_plugins():
    """Test listing available plugins / Probar listado de plugins"""
    plugins = list_available_plugins()
    assert "carta_manifestacion" in plugins


def test_get_oficinas():
    """Test getting oficinas / Probar obtencion de oficinas"""
    plugin = load_plugin("carta_manifestacion")
    oficinas = plugin.get_oficinas()

    assert oficinas is not None
    assert "BARCELONA" in oficinas
    assert "Direccion_Oficina" in oficinas["BARCELONA"]


def test_get_sections():
    """Test getting sections / Probar obtencion de secciones"""
    plugin = load_plugin("carta_manifestacion")
    sections = plugin.get_sections()

    assert sections is not None
    assert len(sections) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
