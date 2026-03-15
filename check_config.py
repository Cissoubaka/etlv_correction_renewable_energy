#!/usr/bin/env python3
"""Diagnostic des deadlines stockées en config"""

import json
from pathlib import Path

config_dir = Path.home() / ".correction_app"
config_file = config_dir / "config.json"

if config_file.exists():
    with open(config_file) as f:
        config = json.load(f)
    
    print("📋 Contenu de la configuration:\n")
    
    if "student_deadlines" in config:
        print(f"✓ Student deadlines trouvées ({len(config['student_deadlines'])} élèves)\n")
        
        # Afficher les 3 premiers
        for i, (student, deadlines) in enumerate(config["student_deadlines"].items()):
            if i >= 3:
                break
            print(f"  {student}:")
            for key, value in deadlines.items():
                print(f"    - {key}: {value}")
        print()
    else:
        print("❌ Pas de 'student_deadlines' en config\n")
    
    if "deadlines" in config:
        print(f"✓ Global deadlines trouvées:")
        for key, value in config["deadlines"].items():
            print(f"  - {key}: {value}")
    else:
        print("❌ Pas de 'deadlines' globales en config")
        
    print(f"\nFichier de config: {config_file}")
else:
    print(f"❌ Fichier config non trouvé: {config_file}")
