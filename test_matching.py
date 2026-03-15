#!/usr/bin/env python3
"""Test du matching de noms"""

def find_matching_folder(tableur_name, folder_names):
    """Essaie de trouver le dossier correspondant au nom du tableur"""
    # Chercher d'abord une correspondance exacte (case-insensitive)
    for folder in folder_names:
        if folder.lower() == tableur_name.lower():
            return folder
    
    # Extraire les parties du nom et chercher des correspondances partielles
    parts = tableur_name.split()
    for part in parts:
        part_lower = part.lower()
        # Chercher un dossier qui contient ce mot-clé (au moins 4 caractères)
        if len(part_lower) >= 4:
            for folder in folder_names:
                if part_lower in folder.lower():
                    return folder
    
    return None

# Noms du tableur et dossiers
tableur_names = [
    "ABATE Tom",
    "ARENE Logan",
    "AZAUBERT Esteban",
    "BATTESTI Léo",
    "BERLENGERO--MURET Dylan",
    "BOUFERAA SAADI Nizar",
    "BRISELET-GIANSANTE Léon",
]

folder_names = [
    "ABATET",
    "ARENE",
    "AZAUBERT1",
    "BATTESTI",
    "BERLENGERO--MURET",
    "BOUFERAASAADI",
    "BRISELET-GIANSANTE",
]

print("🧪 Test du matching de noms:\n")
for tableur in tableur_names:
    match = find_matching_folder(tableur, folder_names)
    status = "✓" if match else "❌"
    print(f"  {status} '{tableur}' → '{match}'")
