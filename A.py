import sys
print(f"Python version: {sys.version}")

try:
    import streamlit as st
    print(f"✅ Streamlit instalado - Versión: {st.__version__}")
except ImportError:
    print("❌ Streamlit NO está instalado")

try:
    import docx
    print("✅ python-docx instalado")
except ImportError:
    print("❌ python-docx NO está instalado")