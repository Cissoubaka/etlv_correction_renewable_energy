#!/usr/bin/env python3
"""
Script de construction d'un exécutable Windows
Utilise PyInstaller pour créer un .exe autonome
"""

import subprocess
import sys
from pathlib import Path

def build_executable():
    """Construit l'exécutable Windows pour l'application de correction"""
    
    project_dir = Path(__file__).parent
    app_file = project_dir / "app_correction.py"
    output_dir = project_dir / "dist"
    build_dir = project_dir / "build"
    
    print("🔨 Construction de l'exécutable Windows...\n")
    
    # Commande PyInstaller
    pyinstaller_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=CorrectionApp",
        "--clean",  # Évite de réutiliser un cache PyInstaller cassé
        "--onefile",  # Un seul fichier .exe
        "--windowed",  # Pas de console window
        "--icon=app_icon.ico" if (project_dir / "app_icon.ico").exists() else "",
        "--add-data=README.md:.",  # Inclure fichiers de données si nécessaire
        "--hidden-import=PyPDF2",
        "--hidden-import=pdfplumber",
        "--hidden-import=pdf2image",
        "--hidden-import=PIL",
        "--hidden-import=openpyxl",
        "--hidden-import=odf",
        "--hidden-import=charset_normalizer",
        "--collect-all=charset_normalizer",
        str(app_file)
    ]
    
    # Nettoyer les arguments vides
    pyinstaller_cmd = [arg for arg in pyinstaller_cmd if arg]
    
    print(f"📦 Commande: {' '.join(pyinstaller_cmd[:5])}...\n")
    
    try:
        # Lancer PyInstaller
        result = subprocess.run(pyinstaller_cmd, cwd=project_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            exe_path = output_dir / "CorrectionApp.exe"
            print(f"✅ Exécutable créé avec succès !")
            print(f"📍 Localisation: {exe_path}")
            print(f"\n📋 Fichiers générés:")
            print(f"   - {exe_path} (Exécutable)")
            print(f"   - {build_dir}/ (Fichiers de construction)")
            print(f"\n💡 Pour distribuer:")
            print(f"   - Copier le fichier '{exe_path.name}' sur Windows")
            print(f"   - Le fichier est autonome, aucune dépendance requise")
        else:
            print(f"❌ Erreur lors de la construction:")
            print(result.stderr)
            return 1
    
    except FileNotFoundError:
        print("❌ PyInstaller n'est pas installé")
        print("   Installez avec: pip install pyinstaller")
        return 1
    except Exception as e:
        print(f"❌ Erreur: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(build_executable())
