#!/usr/bin/env python3
"""Test rapide du parser de spreadsheet"""

import sys
sys.path.insert(0, '/home/cissou/etlv_correction_renewable_energy')

from app_correction import SpreadsheetParser

file_path = "/home/cissou/etlv_correction_renewable_energy/1T1 ETLV 2025-2026 - 03 Renewable technology.ods"

print(f"🧪 Test du parser pour: {file_path}\n")

try:
    result = SpreadsheetParser.parse_spreadsheet(file_path)
    
    print(f"\n✅ Résultat du parsing:\n")
    
    for student_name in sorted(result.keys())[:5]:  # Afficher les 5 premiers
        print(f"  📝 {student_name}:")
        for ws_num in sorted(result[student_name].keys()):
            deadline = result[student_name][ws_num]
            print(f"     - Worksheet {ws_num}: {deadline}")
    
    print(f"\n📊 Total: {len(result)} élèves importés")
    
except Exception as e:
    print(f"❌ ERREUR: {e}")
    import traceback
    traceback.print_exc()
