#!/usr/bin/env python3
"""Test de récupération des deadlines"""

import sys
sys.path.insert(0, '/home/cissou/etlv_correction_renewable_energy')

from app_correction import ConfigManager

# Créer une instance du ConfigManager
config_mgr = ConfigManager()

# Tester pour ABATE Tom
student = "ABATE Tom"
print(f"🧪 Test pour {student}\n")

for ws in [1, 2, 3, 4]:
    # Récupérer directement de la config
    if "student_deadlines" in config_mgr.config:
        if student in config_mgr.config["student_deadlines"]:
            deadline_str = config_mgr.config["student_deadlines"][student].get(f"worksheet{ws}")
            print(f"  Raw config WS{ws}: {deadline_str}")

# Appeler les méthodes pour voir ce qu'elles retournent
print()
print("Après parse de date:")

from datetime import datetime

for ws in [1, 2, 3, 4]:
    if "student_deadlines" in config_mgr.config:
        if student in config_mgr.config["student_deadlines"]:
            deadline_str = config_mgr.config["student_deadlines"][student].get(f"worksheet{ws}")
            if deadline_str:
                try:
                    parsed = datetime.strptime(deadline_str, "%Y-%m-%d").date()
                    print(f"  WS{ws}: {parsed}")
                except Exception as e:
                    print(f"  WS{ws}: ERREUR - {e}")
