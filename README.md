# Correcteur de Formulaires PDF

Application graphique pour corriger rapidement les formulaires PDF en comparant les travaux des étudiants avec les corrections.

## Fonctionnalités

✅ **Sélection des répertoires** : Choisissez le dossier de travail des élèves et celui des corrections
✅ **Liste des élèves** : Affiche automatiquement tous les élèves et leurs worksheets
✅ **Vérification de deadline** : Détecte les travaux rendus en retard
✅ **Affichage des dates** : Voir la date de modification de chaque fichier
✅ **Comparaison side-by-side** : Affichez le travail de l'étudiant et la correction côte à côte
✅ **Aperçu PDF** : Visualisation des PDFs directement dans l'application

## Structure des répertoires attendue

```
/chemin/vers/travaux_eleves/
├── eleve1/
│   ├── worksheet1.pdf
│   ├── worksheet2.pdf
│   ├── worksheet3.pdf
│   └── worksheet4.pdf
├── eleve2/
│   ├── worksheet1.pdf
│   ├── worksheet2.pdf
│   ...
└── ...

/chemin/vers/corrections/
├── correction_worksheet1.pdf
├── correction_worksheet2.pdf
├── correction_worksheet3.pdf
└── correction_worksheet4.pdf
```

## Installation

### 1. Prérequis

- Python 3.7+
- pip

### 2. Installer les dépendances

```bash
pip install -r requirements.txt
```

**Note pour Linux (Ubuntu/Debian)** : Vous devez aussi installer `poppler` :
```bash
sudo apt-get install poppler-utils
```

**Note pour macOS** :
```bash
brew install poppler
```

**Note pour Windows** : poppler est généralement téléchargé automatiquement par `pdf2image`

### 3. Lancer l'application

```bash
python3 app_correction.py
```

## Utilisation

1. **Définir la deadline** : Cliquez sur "📅 Définir deadline" pour indiquer la date limite de rendu
2. **Sélectionner le dossier des élèves** : Cliquez sur "📁 Sélectionner dossier élèves"
3. **Sélectionner le dossier des corrections** : Cliquez sur "📁 Sélectionner dossier corrections"
4. **Corriger rapidement** : 
   - Cliquez sur un élève/worksheet dans la liste de gauche
   - Regardez le travail de l'étudiant et la correction côte à côte
   - Vérifiez si le travail a été remis à temps dans le tableau d'informations

## Informations affichées

Pour chaque fichier sélectionné :
- **Étudiant** : Nom du dossier de l'élève
- **Fichier** : Nom du fichier worksheet
- **Date de modification** : Date et heure du dernier enregistrement
- **Taille** : Taille du fichier en KB
- **Statut deadline** : ✓ À temps ou ⚠️ RETARD (en rouge)

## Troubleshooting

**"pdf2image non installé"** :
- Pour Linux : `sudo apt-get install poppler-utils`
- Pour macOS : `brew install poppler`
- Puis réinstallez : `pip install pdf2image`

**"Le PDF ne s'affiche pas"** :
- Vérifiez que le fichier PDF est valide
- Essayez de l'ouvrir manuellement avec un lecteur PDF

**"Aucun élève trouvé"** :
- Vérifiez que les dossiers d'élèves contiennent des fichiers `worksheet*.pdf`
- Les fichiers doivent être nommés exactement : `worksheet1.pdf`, `worksheet2.pdf`, etc.

## Notes

- L'application détecte automatiquement tous les fichiers `worksheet*.pdf` dans chaque dossier d'élève
- Les corrections doivent être nommées `correction_worksheet1.pdf`, `correction_worksheet2.pdf`, etc.
- Les dates sont basées sur la date de modification du fichier
- L'interface affiche l'aperçu de la première page du PDF
