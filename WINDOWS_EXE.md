# 📦 Créer un Exécutable Windows

## Vue d'ensemble

Vous pouvez créer un exécutable Windows autonome `.exe` à partir de ce script Python en utilisant **PyInstaller**.

## Prérequis

- Python 3.7+
- PyInstaller (`pip install pyinstaller`)
- Toutes les dépendances du projet

## Option 1 : Utiliser le script automatique (Recommandé)

### Sur Linux/Mac :
```bash
bash build_windows.sh
```

### Sur Windows PowerShell :
```powershell
python build_exe.py
```

### Sur Windows CMD :
```cmd
python build_exe.py
```

## Option 2 : Commande manuelle

```bash
python -m PyInstaller \
    --name=CorrectionApp \
    --onefile \
    --windowed \
    --hidden-import=PyPDF2 \
    --hidden-import=pdfplumber \
    --hidden-import=pdf2image \
    --hidden-import=PIL \
    --hidden-import=openpyxl \
    --hidden-import=odf \
    --hidden-import=correction \
    --collect-all=PyQt5 \
    app_correction.py
```

## Résultat

Après la construction, vous obtiendrez :

```
dist/
├── CorrectionApp.exe          ← Cet exécutable sur Windows
build/
    └── [fichiers temporaires]
```

## Utilisation sur Windows

1. **Copier le fichier** `dist/CorrectionApp.exe`
2. **Placer-le n'importe où** sur Windows
3. **Double-cliquer** pour lancer l'application

⚠️ **Note**: Aucune dépendance supplémentaire n'est requise - c'est un fichier autonome !

## Options PyInstaller

| Option | Description |
|--------|-------------|
| `--onefile` | Crée un seul fichier .exe (plus lent au démarrage) |
| `--windowed` | Pas de console (mode GUI uniquement) |
| `--hidden-import` | Ajouter les imports dynamiques |
| `--collect-all` | Collecter tous les fichiers du package |
| `--name` | Nom de l'exécutable |
| `--icon` | Icône personnalisée (.ico) |

## Créer une icône personnalisée

Si vous avez une icône (16x16, 32x32, 64x64px en PNG) :

```bash
# Convertir PNG → ICO (sur Linux)
convert app_icon.png app_icon.ico

# Puis ajouter à la commande:
--icon=app_icon.ico
```

## Réduire la taille de l'exécutable

Pour un fichier plus petit :

```bash
python -m PyInstaller \
    --onefile \
    --windowed \
    --noupx \
    --strip \
    app_correction.py
```

## Dépannage

### Erreur : "Module not found"
Ajouter `--hidden-import=NOM_DU_MODULE` pour chaque module manquant.

### Exécutable trop volumineux
- Utiliser `--onedir` au lieu de `--onefile` (crée un dossier)
- Supprimer les dépendances non-utilisées

### Problèmes au démarrage sur Windows
Vérifier qu'une `Console Python` (Python executable) est correctement installée sur Windows.

## Distribution

### Fichier zip (idéal)
```bash
# Linux/Mac
zip -r CorrectionApp.zip dist/CorrectionApp.exe

# Windows
# Clic droit → Envoyer vers → Dossier compressé
```

### Installer avec un installateur
Pour sophistiqué, utiliser **Inno Setup** ou **NSIS** :
1. Télécharger depuis https://jrsoftware.org/isinfo.php
2. Créer un script `.iss`
3. Compiler en `.msi` ou `.exe` installer

## Support Windows

Cet exécutable fonctionne sur :
- ✅ Windows 10
- ✅ Windows 11
- ✅ Windows Server 2019+

**Architecture** : 64-bit (standard Python 3.10+)

## Exemple complet

```bash
# 1. Installer PyInstaller
pip install pyinstaller

# 2. Construire
python -m PyInstaller --onefile --windowed \
    --hidden-import=PyPDF2 \
    --hidden-import=pdfplumber \
    --collect-all=PyQt5 \
    app_correction.py

# 3. Vérifier
ls -lh dist/CorrectionApp.exe

# 4. Copier sur Windows et tester
```

## Questions fréquentes

**Q: L'exe doit-il être sur le même ordinateur que les PDFs?**
R: Non, l'exe est autonome. Juste accéder aux fichiers via le chemin réseau si besoin.

**Q: Peut-on signer le .exe?**
R: Oui, avec un certificat de signature de code (CodeSign).

**Q: L'exe peut-il être scanné par l'antivirus?**
R: Possiblement au premier démarrage (faux positif commun avec PyInstaller). Ajouter une exception si nécessaire.

**Q: Quelle est la taille typique?**
R: ~200-350 MB pour une app PyQt5 complète.
