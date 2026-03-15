#!/usr/bin/env python3
"""
Analyse les PDFs formulaires pour détecter les doublons et similarités
- Compare les hashes pour les copies exactes
- Extrait le texte UNIQUEMENT des zones de formulaire (pas les en-têtes/pieds de page)
- Compare la similarité du contenu
"""

import os
import hashlib
import tempfile
import sys
from pathlib import Path
from collections import defaultdict
from difflib import SequenceMatcher
import pdfplumber
from datetime import datetime

try:
    import networkx as nx
    import matplotlib.pyplot as plt
    GRAPH_AVAILABLE = True
except ImportError:
    GRAPH_AVAILABLE = False

try:
    from PyPDF2 import PdfReader
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

# Configuration
PDF_DIR = "/media/sauvegarde/etlv1/"
OUTPUT_FILE = "/home/cissou/analyse_pdf.txt"
GRAPH_FILE = "/home/cissou/analyse_pdf_graph.png"
REFERENCE_PDF = "/media/sauvegarde/correction_worksheet1.pdf"  # À adapter si nécessaire
DEBUG_MODE = True  # Afficher les détails d'extraction

class Logger:
    """Classe pour logger à la fois sur stdout et dans un fichier"""
    def __init__(self, filepath):
        self.filepath = filepath
        self.file = open(filepath, 'w', encoding='utf-8')
    
    def log(self, message=""):
        """Écrit le message sur stdout et dans le fichier"""
        print(message)
        self.file.write(message + '\n')
        self.file.flush()
    
    def close(self):
        """Ferme le fichier"""
        self.file.close()

def calculate_hash(filepath):
    """Calcule le hash SHA256 d'un fichier"""
    sha256_hash = hashlib.sha256()
    with open(filepath, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def extract_form_text(pdf_path):
    """
    Extrait le texte UNIQUEMENT des zones de formulaire.
    Stratégie :
    1. Cherche les champs AcroForm (champs interactifs)
    2. Détecte les rectangles/boîtes (zones de saisie)
    3. Extrait le texte DANS ces zones
    """
    text_content = []
    
    try:
        # Essayer d'extraire les champs AcroForm (formulaires interactifs)
        acroform_text = extract_acroform_fields(pdf_path)
        if acroform_text:
            text_content.append(acroform_text)
        
        with pdfplumber.open(pdf_path) as pdf:
            # Ensuite, détecter les zones de saisie par les rectangles
            for page_num, page in enumerate(pdf.pages):
                rect_text = extract_text_in_boxes(page)
                if rect_text:
                    text_content.append(f"\n[PAGE {page_num + 1}]\n{rect_text}")
    
    except Exception as e:
        print(f"⚠️ Erreur lors de la lecture de {pdf_path}: {e}")
        return ""
    
    return "\n".join(text_content).strip()

def extract_acroform_fields(pdf_path):
    """
    Extrait les valeurs des champs AcroForm (formulaires interactifs)
    """
    if not PYPDF2_AVAILABLE:
        return ""
    
    try:
        reader = PdfReader(pdf_path)
        
        # Extraire les champs remplis
        fields_text = []
        if reader.get_fields():
            for field_name, field_data in reader.get_fields().items():
                if field_data and field_data.get('/V'):  # /V est la valeur du champ
                    value = field_data['/V']
                    if isinstance(value, bytes):
                        value = value.decode('utf-8', errors='ignore')
                    fields_text.append(f"{field_name}: {value}")
        
        return "\n".join(fields_text) if fields_text else ""
    except Exception:
        return ""

def extract_text_in_boxes(page):
    """
    Détecte les rectangles/boîtes du PDF et extrait le texte DEDANS
    Les rectangles correspondent généralement aux zones de saisie
    """
    text_content = []
    
    try:
        # Obtenir les tableaux détectés (zones structurées)
        tables = page.find_tables()
        
        if tables:
            # S'il y a des tableaux, extraire le texte de chaque cellule
            for table in tables:
                for row in table:
                    for cell in row:
                        if cell:
                            text = cell.strip()
                            if text:
                                text_content.append(text)
        else:
            # Sinon, extraire le texte par lignes (groupé par position Y)
            chars = page.chars
            if chars:
                lines = {}
                for char in chars:
                    y = round(char['top'])
                    if y not in lines:
                        lines[y] = []
                    lines[y].append(char)
                
                # Extraire le texte par ligne
                for y in sorted(lines.keys()):
                    line_chars = sorted(lines[y], key=lambda c: c['x0'])
                    line_text = "".join([c['text'] for c in line_chars]).strip()
                    if line_text and len(line_text) > 1:  # Ignorer les caractères isolés
                        text_content.append(line_text)
    
    except Exception as e:
        pass
    
    return "\n".join(text_content)

def cleanup_text(text):
    """Nettoie le texte en supprimant les espaces superflus"""
    lines = [line.strip() for line in text.split('\n')]
    return '\n'.join([line for line in lines if line])

def normalize_text(text):
    """
    Normalise le texte pour la comparaison :
    - Casse insensible
    - Nombres en lettres -> chiffres
    - Espaces/accents superflus
    """
    if not text:
        return ""
    
    # Convertir en minuscules
    text = text.lower().strip()
    
    # Dictionnaire des nombres en lettres français
    numbers_map = {
        'zéro': '0', 'zero': '0',
        'un': '1', 'une': '1',
        'deux': '2',
        'trois': '3',
        'quatre': '4',
        'cinq': '5',
        'six': '6',
        'sept': '7',
        'huit': '8',
        'neuf': '9',
        'dix': '10',
        'onze': '11',
        'douze': '12',
        'treize': '13',
        'quatorze': '14',
        'quinze': '15',
        'seize': '16',
        'dix-sept': '17', 'dixsept': '17',
        'dix-huit': '18', 'dixhuit': '18',
        'dix-neuf': '19', 'dixneuf': '19',
        'vingt': '20',
        'trente': '30',
        'quarante': '40',
        'cinquante': '50',
        'soixante': '60',
        'septante': '70', 'soixante-dix': '70',
        'huitante': '80', 'quatre-vingts': '80', 'quatrevingts': '80',
        'nonante': '90', 'quatre-vingt-dix': '90', 'quatrevingtdix': '90',
        'cent': '100', 'cents': '100',
        'mille': '1000', 'mil': '1000',
        'million': '1000000',
    }
    
    # Remplacer les nombres en lettres
    for word, number in numbers_map.items():
        text = text.replace(word, number)
    
    # Supprimer les accents
    import unicodedata
    text = ''.join(c for c in unicodedata.normalize('NFD', text)
                   if unicodedata.category(c) != 'Mn')
    
    # Supprimer les espaces multiples
    text = ' '.join(text.split())
    
    return text

def extract_responses_from_pdf(pdf_path, extract_type='both'):
    """
    Extrait UNIQUEMENT les réponses du formulaire (pas les questions)
    extract_type: 'acroform' (champs interactifs), 'boxes' (boîtes de texte), 'both'
    """
    responses = []
    
    try:
        # Essayer d'abord les champs AcroForm (si disponible)
        if extract_type in ['acroform', 'both'] and PYPDF2_AVAILABLE:
            try:
                reader = PdfReader(pdf_path)
                if reader.get_fields():
                    for field_idx, (field_name, field_data) in enumerate(reader.get_fields().items()):
                        if field_data and field_data.get('/V'):  # /V = valeur remplie
                            value = field_data['/V']
                            if isinstance(value, bytes):
                                value = value.decode('utf-8', errors='ignore')
                            
                            value_str = str(value).strip()
                            if value_str:
                                responses.append({
                                    'id': f"ACRO_{field_idx}",
                                    'field_name': field_name,
                                    'original': value_str,
                                    'normalized': normalize_text(value_str),
                                    'page': 0,
                                    'type': 'acroform'
                                })
            except:
                pass
        
        # Extraire aussi les boîtes de texte (si aucun AcroForm ou extract_type='boxes'/'both')
        if extract_type in ['boxes', 'both']:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    # Chercher les rectangles (boîtes de saisie)
                    rects = page.rects
                    
                    if rects:
                        # Pour chaque rectangle, extraire le texte DEDANS
                        for rect_idx, rect in enumerate(rects):
                            # Ignorer les rectangles trop petits (probablement des traits)
                            width = rect['x1'] - rect['x0']
                            height = rect['y1'] - rect['y0']
                            
                            if width > 20 and height > 10:  # Taille minimale pour une zone de saisie
                                # Crée une bbox correcte (x0, top, x1, bottom) pour page.crop()
                                bbox = (rect['x0'], rect['top'], rect['x1'], rect['bottom'])
                                try:
                                    text_inside = page.crop(bbox).extract_text()
                                except Exception:
                                    # Si crop échoue, essayer avec une approche alternative
                                    continue
                                
                                if text_inside and text_inside.strip():
                                    text = text_inside.strip()
                                    # Filtrer: si c'est une question (généralement > 50 caractères ou avec ponctuation), ignorer
                                    if len(text) < 100:  # Les réponses sont généralement courtes
                                        responses.append({
                                            'id': f"P{page_num+1}_BOX_{rect_idx}",
                                            'original': text,
                                            'normalized': normalize_text(text),
                                            'page': page_num + 1,
                                            'type': 'box'
                                        })
    
    except Exception as e:
        print(f"⚠️ Erreur d'extraction des réponses : {e}")
    
    return responses

def extract_fields_from_pdf(pdf_path):
    """
    Ancien nom : appel à extract_responses_from_pdf pour compatibilité
    """
    return extract_responses_from_pdf(pdf_path, extract_type='both')

def compare_pdf_with_reference(pdf_path, reference_fields, logger=None):
    """
    Compare un PDF avec les champs de référence de manière rigoureuse.
    Compare champ par champ en ordre, affiche les divergences.
    """
    pdf_fields = extract_fields_from_pdf(pdf_path)
    
    if not pdf_fields:
        return 0, len(reference_fields), [], reference_fields, []
    
    # Trier les champs par page et position pour une comparaison cohérente
    pdf_fields_sorted = sorted(pdf_fields, key=lambda f: (f['page'], f['id']))
    ref_fields_sorted = sorted(reference_fields, key=lambda f: (f['page'], f['id']))
    
    matched = []
    unmatched_pdf = []
    unmatched_ref = []
    score = 0
    
    # Créer des ensembles pour la comparaison
    pdf_normalized = {f['normalized']: f for f in pdf_fields_sorted}
    ref_normalized = {f['normalized']: f for f in ref_fields_sorted}
    
    # Champs de référence qui ont une correspondance
    ref_found = set()
    
    # Comparer: pour chaque champ du PDF, chercher s'il existe en référence
    for pdf_field in pdf_fields_sorted:
        normalized_text = pdf_field['normalized']
        
        if normalized_text in ref_normalized:
            score += 1
            ref_found.add(normalized_text)
            matched.append({
                'pdf': pdf_field,
                'ref': ref_normalized[normalized_text],
                'match': True
            })
        else:
            # Ce champ du PDF ne correspond à aucun champ de référence
            unmatched_pdf.append({
                'pdf': pdf_field,
                'type': 'extra_in_pdf'
            })
    
    # Champs de référence qui n'ont pas de correspondance
    for ref_field in ref_fields_sorted:
        normalized_text = ref_field['normalized']
        if normalized_text not in ref_found:
            unmatched_ref.append({
                'ref': ref_field,
                'type': 'missing_in_pdf'
            })
    
    return score, len(ref_fields_sorted), matched, unmatched_pdf, unmatched_ref

def calculate_similarity(text1, text2):
    """Calcule la similarité entre deux textes en pourcentage"""
    if not text1 or not text2:
        return 0
    
    # Utiliser SequenceMatcher pour une comparaison fiable
    matcher = SequenceMatcher(None, text1, text2)
    ratio = matcher.ratio()
    return int(ratio * 100)

def create_similarity_graph(extracted_texts, logger):
    """Crée un graphique des similarités à 100% entre fichiers"""
    if not GRAPH_AVAILABLE:
        logger.log("\n⚠️ networkx/matplotlib non installés - graphique non généré")
        logger.log("   Installez avec : pip install networkx matplotlib")
        return
    
    MIN_TEXT_LENGTH = 100  # Minimum de caractères pour considérer un texte valide
    
    # Filtrer les textes vides ou trop courts
    valid_texts = {}
    for pdf_path, text in extracted_texts.items():
        cleaned = cleanup_text(text)
        if len(cleaned) >= MIN_TEXT_LENGTH:
            valid_texts[pdf_path] = cleaned
    
    logger.log(f"\nFichiers avec texte valide : {len(valid_texts)}/{len(extracted_texts)}")
    
    if len(valid_texts) < 2:
        logger.log("✓ Pas assez de fichiers avec du contenu pour générer un graphique")
        return
    
    # Créer un graphique
    G = nx.Graph()
    
    # Ajouter les nœuds
    for pdf_path in valid_texts.keys():
        node_name = Path(pdf_path).name
        G.add_node(node_name)
    
    # Ajouter les arêtes pour les similarités à 100%
    pdf_list = list(valid_texts.keys())
    edges_100 = 0
    
    for i in range(len(pdf_list)):
        for j in range(i + 1, len(pdf_list)):
            pdf1 = pdf_list[i]
            pdf2 = pdf_list[j]
            
            # S'assurer qu'on compare des fichiers différents
            if pdf1 == pdf2:
                continue
            
            text1 = valid_texts[pdf1]
            text2 = valid_texts[pdf2]
            
            similarity = calculate_similarity(text1, text2)
            
            if similarity == 100:
                edges_100 += 1
                node1 = Path(pdf1).name
                node2 = Path(pdf2).name
                logger.log(f"\n  ⚠️ Similarité EXACTE détectée :")
                logger.log(f"     - {node1} ({len(text1)} chars)")
                logger.log(f"     - {node2} ({len(text2)} chars)")
                G.add_edge(node1, node2, weight=similarity)
    
    # Si pas d'arêtes, pas besoin de créer un graphique
    if edges_100 == 0:
        logger.log("\n✓ Aucune similarité à 100% détectée")
        return
    
    # Créer le graphique avec matplotlib
    try:
        plt.figure(figsize=(14, 10))
        
        # Utiliser spring layout pour une meilleure visualisation
        pos = nx.spring_layout(G, k=2, iterations=50, seed=42)
        
        # Couleurs pour les nœuds
        node_colors = []
        for node in G.nodes():
            if G.degree(node) > 1:
                node_colors.append('#ff6b6b')  # Rouge pour les nœuds avec plusieurs connexions
            else:
                node_colors.append('#4ecdc4')  # Turquoise pour les nœuds isolés
        
        # Dessiner les nœuds
        nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=2000, alpha=0.9)
        
        # Dessiner les arêtes
        nx.draw_networkx_edges(G, pos, width=2, alpha=0.6, edge_color='#666666')
        
        # Dessiner les labels
        labels = {node: node[:20] + "..." if len(node) > 20 else node for node in G.nodes()}
        nx.draw_networkx_labels(G, pos, labels, font_size=8, font_weight='bold')
        
        # Ajouter un titre et des infos
        plt.title("Réseau des Formulaires Identiques (Similarité 100%)", fontsize=16, fontweight='bold')
        plt.text(0.5, -0.05, f"Fichiers analysés : {len(valid_texts)} | Connexions : {edges_100}",
                ha='center', fontsize=10, transform=plt.gca().transAxes)
        plt.axis('off')
        plt.tight_layout()
        
        # Sauvegarder l'image
        plt.savefig(GRAPH_FILE, dpi=150, bbox_inches='tight', facecolor='white')
        plt.close()
        
        logger.log(f"\n✓ Graphique créé : {GRAPH_FILE}")
        logger.log(f"  - {edges_100} connexions (similarité 100%)")
        logger.log(f"  - {len(G.nodes())} fichiers impliqués")
    
    except Exception as e:
        logger.log(f"\n❌ Erreur lors de la création du graphique : {e}")

def main():
    logger = Logger(OUTPUT_FILE)
    
    logger.log("\n" + "=" * 50)
    logger.log("🔍 ANALYSE DES PDF - FORMULAIRES UNIQUEMENT")
    logger.log(f"📁 Dossier : {PDF_DIR}")
    logger.log(f"⏰ Timestamp : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.log("=" * 50)
    
    # Avertissements sur les dépendances optionnelles
    if not PYPDF2_AVAILABLE:
        logger.log("⚠️ PyPDF2 non installé - extraction des champs AcroForm désactivée")
        logger.log("   Installez avec : pip install PyPDF2")
    
    if not GRAPH_AVAILABLE:
        logger.log("⚠️ networkx/matplotlib non installés - graphique désactivé")
        logger.log("   Installez avec : pip install networkx matplotlib")
    
    # Vérifier que le répertoire existe
    if not os.path.isdir(PDF_DIR):
        logger.log(f"❌ Le répertoire {PDF_DIR} n'existe pas")
        logger.close()
        return
    
    # Trouver tous les PDF
    pdf_files = list(Path(PDF_DIR).glob("**/*.pdf"))
    if not pdf_files:
        logger.log("❌ Aucun fichier PDF trouvé")
        logger.close()
        return
    
    logger.log(f"✓ {len(pdf_files)} fichiers PDF trouvés\n")
    
    # ==================================================
    # 1. FICHIERS AVEC MÊME TAILLE
    # ==================================================
    logger.log("\n📊 1. FICHIERS AVEC MÊME TAILLE")
    logger.log("----------------------------------------")
    
    sizes = defaultdict(list)
    for pdf_file in pdf_files:
        size = os.path.getsize(pdf_file)
        sizes[size].append(str(pdf_file))
    
    duplicates_by_size = {s: files for s, files in sizes.items() if len(files) > 1}
    
    if duplicates_by_size:
        for size in sorted(duplicates_by_size.keys()):
            files = duplicates_by_size[size]
            logger.log(f"\n⚠️ Taille : {size} octets ({len(files)} fichiers)")
            for f in files:
                logger.log(f"   - {f}")
    else:
        logger.log("\n✓ Aucun fichier avec la même taille")
    
    # ==================================================
    # 2. FICHIERS STRICTEMENT IDENTIQUES (HASH)
    # ==================================================
    logger.log("\n\n🔐 2. FICHIERS STRICTEMENT IDENTIQUES (HASH)")
    logger.log("----------------------------------------")
    
    hashes = defaultdict(list)
    for pdf_file in pdf_files:
        print(f"  Vérification hash : {pdf_file.name}", end='\r', flush=True)
        hash_val = calculate_hash(str(pdf_file))
        hashes[hash_val].append(str(pdf_file))
    
    print(" " * 80, end='\r', flush=True)  # Nettoyer la ligne de progression
    
    duplicates_by_hash = {h: files for h, files in hashes.items() if len(files) > 1}
    
    if duplicates_by_hash:
        for hash_val in sorted(duplicates_by_hash.keys()):
            files = duplicates_by_hash[hash_val]
            logger.log(f"\n🚨 COPIE EXACTE DÉTECTÉE ({len(files)} fichiers)")
            logger.log(f"Hash : {hash_val}")
            for f in files:
                logger.log(f"   - {f}")
    else:
        logger.log("\n✓ Aucun doublon exact détecté")
    
    # ==================================================
    # 3. EXTRACTION DU TEXTE DES FORMULAIRES
    # ==================================================
    logger.log("\n\n📝 3. EXTRACTION DU TEXTE DES FORMULAIRES")
    logger.log("----------------------------------------")
    logger.log("Extraction en cours...")
    
    extracted_texts = {}
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"  [{i}/{len(pdf_files)}] {pdf_file.name}", end='\r', flush=True)
        text = extract_form_text(str(pdf_file))
        extracted_texts[str(pdf_file)] = text
    
    print(" " * 80, end='\r', flush=True)  # Nettoyer la ligne de progression
    logger.log("✓ Extraction terminée")
    
    # ==================================================
    # 4. COMPARAISON AVEC PDF DE RÉFÉRENCE
    # ==================================================
    logger.log("\n\n📋 4. COMPARAISON AVEC PDF DE RÉFÉRENCE")
    logger.log("----------------------------------------")
    
    if os.path.isfile(REFERENCE_PDF):
        logger.log(f"🔍 Chargement de la référence : {Path(REFERENCE_PDF).name}")
        reference_fields = extract_fields_from_pdf(REFERENCE_PDF)
        logger.log(f"✓ {len(reference_fields)} champs détectés dans la référence")
        
        # Mode débogage : afficher les champs de la référence
        if DEBUG_MODE:
            logger.log("\n" + "=" * 70)
            logger.log("🔎 CHAMPS DÉTECTÉS DANS LA RÉFÉRENCE (MODE DÉBOGAGE)")
            logger.log("=" * 70)
            for i, field in enumerate(reference_fields[:20]):  # Afficher les 20 premiers
                logger.log(f"\n{i+1}. ID: {field['id']} | Page: {field['page']} | Type: {field['type']}")
                logger.log(f"   Original : '{field['original']}'")
                logger.log(f"   Normalisé: '{field['normalized']}'")
                logger.log(f"   Longueur : {len(field['original'])} caractères")
            if len(reference_fields) > 20:
                logger.log(f"\n... et {len(reference_fields) - 20} autres champs")
            logger.log("\n" + "=" * 70)
        
        logger.log("\n📊 Résultats de conformité :\n")
        
        # Mode débogage : comparer le premier PDF avec la référence
        if DEBUG_MODE and pdf_files:
            logger.log("=" * 70)
            logger.log(f"🔎 COMPARAISON DU PREMIER PDF : {pdf_files[0].name} (MODE DÉBOGAGE)")
            logger.log("=" * 70)
            
            first_pdf_fields = extract_fields_from_pdf(str(pdf_files[0]))
            logger.log(f"\n✓ {len(first_pdf_fields)} champs détectés dans le PDF\n")
            
            for i, field in enumerate(first_pdf_fields[:20]):
                logger.log(f"\n{i+1}. ID: {field['id']} | Page: {field['page']} | Type: {field['type']}")
                logger.log(f"   Original : '{field['original']}'")
                logger.log(f"   Normalisé: '{field['normalized']}'")
                logger.log(f"   Longueur : {len(field['original'])} caractères")
            if len(first_pdf_fields) > 20:
                logger.log(f"\n... et {len(first_pdf_fields) - 20} autres champs")
            logger.log("\n" + "=" * 70)
        
        # Créer un tableau de résultats
        results = []
        for pdf_file in pdf_files:
            pdf_name = pdf_file.name
            score, total, matched, unmatched_pdf, unmatched_ref = compare_pdf_with_reference(str(pdf_file), reference_fields, logger)
            
            # Calculer le pourcentage
            percentage = (score / total * 100) if total > 0 else 0
            
            results.append({
                'name': pdf_name,
                'path': str(pdf_file),
                'score': score,
                'total': total,
                'percentage': percentage,
                'matched': matched,
                'unmatched_pdf': unmatched_pdf,
                'unmatched_ref': unmatched_ref
            })
            
            logger.log(f"📄 {pdf_name}")
            logger.log(f"   Score : {score}/{total} champs ({percentage:.1f}%)")
        
        # Afficher les détails des divergences pour chaque fichier
        logger.log("\n" + "=" * 70)
        logger.log("🔍 DÉTAILS DES DIVERGENCES")
        logger.log("=" * 70)
        
        for result in results:
            logger.log(f"\n📋 {result['name']}")
            logger.log("-" * 70)
            
            # Afficher les champs qui ne correspondent pas
            if result['unmatched_pdf']:
                logger.log(f"\n  ❌ Champs EXTRA dans le PDF ({len(result['unmatched_pdf'])}) :")
                for item in result['unmatched_pdf'][:10]:  # Limiter à 10
                    pdf_field = item['pdf']
                    logger.log(f"     - Page {pdf_field['page']}: '{pdf_field['original']}'")
                if len(result['unmatched_pdf']) > 10:
                    logger.log(f"     ... et {len(result['unmatched_pdf']) - 10} autres")
            
            if result['unmatched_ref']:
                logger.log(f"\n  ⚠️  Champs MANQUANTS (attendus mais pas trouvés) ({len(result['unmatched_ref'])}) :")
                for item in result['unmatched_ref'][:10]:  # Limiter à 10
                    ref_field = item['ref']
                    logger.log(f"     - Page {ref_field['page']}: '{ref_field['original']}'")
                if len(result['unmatched_ref']) > 10:
                    logger.log(f"     ... et {len(result['unmatched_ref']) - 10} autres")
            
            if not result['unmatched_pdf'] and not result['unmatched_ref']:
                logger.log("\n  ✅ PARFAIT ! Tous les champs correspondent.")
        
        # Trier par score décroissant et afficher un résumé
        logger.log("\n" + "=" * 50)
        logger.log("📈 RÉSUMÉ GLOBAL")
        logger.log("=" * 50)
        
        results_sorted = sorted(results, key=lambda x: x['percentage'], reverse=True)
        
        for result in results_sorted:
            score_bar = "█" * int(result['percentage'] / 5) + "░" * (20 - int(result['percentage'] / 5))
            logger.log(f"{result['name'][:30]:<30} | {score_bar} | {result['score']}/{result['total']} ({result['percentage']:.1f}%)")
        
        # Moyenne générale
        avg_percentage = sum(r['percentage'] for r in results) / len(results) if results else 0
        logger.log(f"\n📊 Moyenne générale : {avg_percentage:.1f}%")
    else:
        logger.log(f"⚠️ PDF de référence introuvable : {REFERENCE_PDF}")
        logger.log("   Configurez REFERENCE_PDF dans le script")
    
    # ==================================================
    # 5. SIMILARITÉ DES CONTENUS (> 90%)
    # ==================================================
    logger.log("\n\n🔎 5. SIMILARITÉ DES CONTENUS (> 90%)")
    logger.log("----------------------------------------")
    
    similarities_found = False
    
    pdf_list = list(extracted_texts.keys())
    for i in range(len(pdf_list)):
        for j in range(i + 1, len(pdf_list)):
            pdf1 = pdf_list[i]
            pdf2 = pdf_list[j]
            
            text1 = cleanup_text(extracted_texts[pdf1])
            text2 = cleanup_text(extracted_texts[pdf2])
            
            similarity = calculate_similarity(text1, text2)
            
            if similarity > 90:
                similarities_found = True
                logger.log(f"\n⚠️ Similarité : {similarity}%")
                logger.log(f"   - {Path(pdf1).name}")
                logger.log(f"   - {Path(pdf2).name}")
    
    if not similarities_found:
        logger.log("\n✓ Aucune similarité > 90% détectée")
    
    # ==================================================
    # 6. GRAPHIQUE DES SIMILARITÉS 100%
    # ==================================================
    logger.log("\n\n📊 6. GRAPHIQUE DES SIMILARITÉS 100%")
    logger.log("----------------------------------------")
    create_similarity_graph(extracted_texts, logger)
    
    # ==================================================
    # Résumé final
    # ==================================================
    logger.log("\n\n" + "=" * 50)
    logger.log("✅ Analyse terminée")
    logger.log(f"📄 Résultats sauvegardés dans : {OUTPUT_FILE}")
    logger.log("=" * 50 + "\n")
    
    logger.close()

if __name__ == "__main__":
    main()
