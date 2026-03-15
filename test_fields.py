#!/usr/bin/env python3
"""Test d'extraction de champs PDF"""

from pathlib import Path
import pdfplumber
from PyPDF2 import PdfReader

def extract_pdf_fields(pdf_path):
    """
    Extrait les champs d'un formulaire PDF
    Essaie d'abord les champs AcroForm, puis détecte les zones de saisie
    """
    fields = {}
    
    try:
        # Méthode 1: Champs AcroForm
        try:
            reader = PdfReader(str(pdf_path))
            if reader.get_fields():
                field_names = list(reader.get_fields().keys())
                print(f"   📋 Champs AcroForm trouvés: {field_names}")
                for i, field_name in enumerate(field_names):
                    simple_name = field_name.split('.')[-1] if '.' in field_name else field_name
                    fields[f"field_{i+1}"] = {'name': simple_name, 'type': 'AcroForm'}
                
                if fields:
                    print(f"   ✓ {len(fields)} champs AcroForm détectés")
                    return fields
        except Exception as e:
            print(f"   ⚠️ Erreur AcroForm: {e}")
        
        # Méthode 2: Détecter les zones de saisie
        with pdfplumber.open(str(pdf_path)) as pdf:
            if len(pdf.pages) > 0:
                page = pdf.pages[0]
                
                # Chercher les tableaux
                tables = page.find_tables()
                print(f"   📊 Tableaux trouvés: {len(tables) if tables else 0}")
                
                if tables:
                    field_count = 1
                    for table_idx, table in enumerate(tables):
                        for i in range(len(table)):
                            fields[f"field_{field_count}"] = {
                                'name': f'Tableau {table_idx+1} - Cellule {i+1}',
                                'type': 'Table'
                            }
                            field_count += 1
                
                # Chercher les rectangles
                if hasattr(page, 'rects') and page.rects:
                    print(f"   📦 Rectangles trouvés: {len(page.rects)}")
                    rect_count = 1
                    for rect in page.rects[:10]:
                        field_key = f"field_rect_{rect_count}"
                        fields[field_key] = {
                            'name': f'Zone {rect_count}',
                            'type': 'Rectangle'
                        }
                        rect_count += 1
                
                if fields:
                    print(f"   ✓ {len(fields)} zones de saisie détectées")
                    return fields
    
    except Exception as e:
        print(f"   ❌ Erreur: {e}")
    
    print(f"   ⚠️ Aucun champ détecté")
    return {}

# Test sur les fichiers de correction
correction_dir = Path("/home/cissou/etlv_correction_renewable_energy")

for i in range(1, 5):
    pdf_file = correction_dir / f"correction_worksheet{i}.pdf"
    
    if pdf_file.exists():
        print(f"\n🧪 Analyse de {pdf_file.name}:")
        fields = extract_pdf_fields(pdf_file)
        
        if fields:
            print(f"   Détails des champs:")
            for field_key, field_info in list(fields.items())[:5]:
                print(f"     - {field_info['name']} ({field_info['type']})")
            if len(fields) > 5:
                print(f"     ... et {len(fields)-5} autre(s)")
    else:
        print(f"\n❌ {pdf_file.name} n'existe pas")
