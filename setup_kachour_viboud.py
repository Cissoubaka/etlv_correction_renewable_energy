#!/usr/bin/env python3
"""
Script pour ajouter les fichiers de KACHOUR-VIBOUD à la configuration.
Cet élève a des fichiers dans un répertoire différent du répertoire de travail standard.
"""

import json
import re
from pathlib import Path

# Configuration
CONFIG_FILE = Path.home() / ".correction_app" / "config.json"
STUDENT_NAME = "KACHOUR-VIBOUD"
STUDENT_FOLDER = Path("/media/sauvegarde/Travail eleve/KACHOUR-VIBOUD")

# Mapping des fichiers (worksheet -> fichier PDF)
FILE_MAPPING = {
    1: "sohan activité 1.pdf",
    2: "etlv act 2.pdf",
    3: "ETLV activité  3.pdf",
    4: "activité 4.pdf",  # Ou "activité 4 (sohan).pdf"
}

def setup_student():
    """Configure les fichiers pour l'étudiant KACHOUR-VIBOUD"""
    
    print(f"📝 Configuration de {STUDENT_NAME}")
    print(f"📂 Dossier: {STUDENT_FOLDER}\n")
    
    # Vérifier que le dossier existe
    if not STUDENT_FOLDER.exists():
        print(f"❌ Le dossier {STUDENT_FOLDER} n'existe pas!")
        return False
    
    # Charger ou créer la configuration
    CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        config = {"manual_selections": {}}
    
    # S'assurer que la clé manual_selections existe
    if "manual_selections" not in config:
        config["manual_selections"] = {}
    
    if STUDENT_NAME not in config["manual_selections"]:
        config["manual_selections"][STUDENT_NAME] = {}
    
    # Ajouter les fichiers
    success_count = 0
    for ws_num, filename in FILE_MAPPING.items():
        file_path = STUDENT_FOLDER / filename
        
        if file_path.exists():
            config["manual_selections"][STUDENT_NAME][f"worksheet{ws_num}"] = str(file_path)
            print(f"✅ Worksheet {ws_num}: {filename}")
            success_count += 1
        else:
            print(f"⚠️ Worksheet {ws_num}: Fichier non trouvé: {filename}")
            # Lister les fichiers disponibles
            pdf_files = list(STUDENT_FOLDER.glob("*.pdf"))
            if pdf_files:
                print(f"   Fichiers PDF disponibles:")
                for pdf in sorted(pdf_files):
                    print(f"   - {pdf.name}")
    
    if success_count > 0:
        # Sauvegarder la configuration
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ Configuration sauvegardée pour {success_count}/4 worksheets")
        print(f"📝 Fichier de config: {CONFIG_FILE}")
        return True
    else:
        print(f"\n❌ Aucun fichier ne correspond au mapping")
        return False

if __name__ == "__main__":
    setup_student()
