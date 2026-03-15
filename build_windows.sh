#!/bin/bash
# Script de construction pour Windows et Linux

# Vérifier que PyInstaller est installé
echo "📦 Vérification de PyInstaller..."
python3 -m pip list | grep -q pyinstaller || python3 -m pip install pyinstaller

echo "🔨 Construction de l'exécutable..."
echo ""

# Construction pour Windows (si sur Windows) ou simulation Linux
python3 -m PyInstaller \
    --name=CorrectionApp \
    --clean \
    --onefile \
    --windowed \
    --hidden-import=PyPDF2 \
    --hidden-import=pdfplumber \
    --hidden-import=pdf2image \
    --hidden-import=PIL \
    --hidden-import=openpyxl \
    --hidden-import=odf \
    --hidden-import=correction \
    --hidden-import=charset_normalizer \
    --collect-all=charset_normalizer \
    --collect-all=PyQt5 \
    app_correction.py

echo ""
echo "✅ Construction terminée !"
echo "📍 Exécutable disponible dans: ./dist/CorrectionApp.exe (Windows)"
echo ""
echo "Pour distribuer sur Windows:"
echo "  1. Copier le fichier dist/CorrectionApp.exe"
echo "  2. C'est un fichier autonome, aucune dépendance requise"
