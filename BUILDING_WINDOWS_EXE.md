# 🪟 Créer un .exe Windows depuis Linux/Mac

## ⚠️ Le problème

PyInstaller crée des exécutables **natifs** au système d'exploitation :

| OS hôte | Résultat |
|---------|----------|
| Linux | `CorrectionApp` (fichier ELF) |
| macOS | `CorrectionApp.app` (bundle) |
| Windows | `CorrectionApp.exe` ✅ |

Si tu compiles sur Linux, tu obtiendras un exécutable Linux, pas un `.exe` !

---

## ✅ Solutions (de la plus simple à la plus complexe)

### 🥇 **Option 1 : GitHub Actions (RECOMMANDÉ)**

Utilise Git et GitHub pour compiler automatiquement sur Windows.

#### Prérequis
- Compte GitHub (gratuit)
- Ton projet sur GitHub (public ou privé)

#### Étapes

**1. Initialiser Git**
```bash
cd /home/cissou/etlv_correction_renewable_energy
git init
git add .
git commit -m "Initial commit"
```

**2. Créer un repo GitHub**
- Aller sur https://github.com/new
- Créer un repo nommé `etlv_correction_renewable_energy`

**3. Pousser le code**
```bash
git remote add origin https://github.com/TON_USERNAME/etlv_correction_renewable_energy.git
git branch -M main
git push -u origin main
```

**4. Le workflow se déclenche automatiquement**
- Aller sur l'onglet **Actions** de ton repo
- Attendre que "Build Windows Executable" se termine (2-3 minutes)
- Cliquer sur la dernière exécution
- Dans la section **Artifacts**, télécharger `CorrectionApp-Windows`
- Extraire le `.exe` 🎉

#### Déclencher manuellement
```bash
git push  # Relance le build
# OU
# Va sur le repo GitHub → Actions → Build Windows Executable → Run workflow
```

---

### 🥈 **Option 2 : Windows Subsystem for Linux (WSL2) sur Windows**

Si tu as Windows avec WSL2, tu peux installer Python sur Windows natif :

```powershell
# Dans PowerShell (pas WSL)
cd C:\chemin\vers\etlv_correction_renewable_energy
.\build_exe_windows.ps1
```

---

### 🥉 **Option 3 : Machine virtuelle Windows**

1. Créer/accéder à une VM Windows (VirtualBox, VMware, Hyper-V)
2. Y cloner le projet
3. Exécuter :
```cmd
python build_exe.py
```
4. Récupérer le `.exe` depuis `dist/CorrectionApp.exe`

---

### 🔬 **Option 4 : Docker (Advanced)**

Pour compiler dans un conteneur Windows :

```bash
# Nécessite Docker Desktop avec support Windows containers
docker build -f Dockerfile.windows -t correction-builder .
docker run --rm -v $(pwd):/output correction-builder cmd /c "copy C:\app\dist\CorrectionApp.exe C:\output\"
```

---

## 🚀 Méthode recommandée : GitHub Actions

### Pourquoi ?
✅ **Gratuit** - Inclus avec GitHub  
✅ **Automatique** - Crée le .exe à chaque push  
✅ **Pas d'infrastructure** - Pas besoin de Windows  
✅ **Fiable** - Serveurs Microsoft  
✅ **Facile** - Juste copier-coller  

### Résumé des étapes
1. **Push vers GitHub** → Le workflow se déclenche
2. **Attendre 2-3 min** → Compilation sur Windows
3. **Récupérer l'artifact** → `CorrectionApp.exe` prêt

### Structure des fichiers créés

```
.github/workflows/
└── build-windows.yml  ← Workflow GitHub Actions
```

Le workflow est **déjà configuré** pour :
- ✅ Compiler sur Windows
- ✅ Empaqueter dépendances PyQt5
- ✅ Créer un .exe autonome
- ✅ Télécharger comme artifact

---

## ✨ Configuration personnalisée

Tu peux modifier `.github/workflows/build-windows.yml` pour :

### Ajouter une icône personnalisée
```yaml
run: |
  python -m PyInstaller \
    ... \
    --icon=logo.ico \
    app_correction.py
```

### Mettre en auto-release
```yaml
- name: Create Release
  if: startsWith(github.ref, 'refs/tags/')  # À chaque tag
  uses: softprops/action-gh-release@v1
  with:
    files: dist_release/CorrectionApp.exe
```

Puis créer un tag :
```bash
git tag v1.0.0
git push origin v1.0.0
```

---

## 🎯 Commande rapide

### Si tu veux tester le build manuellement sur Linux (simulation)
```bash
# Ceci crée un exécutable Linux, pas Windows
python -m PyInstaller \
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

Résultat : `dist/CorrectionApp` (exécutable Linux)

### Pour le convertir en pseudo-Windows (ne marche pas vraiment)
```bash
# ❌ Ne PAS faire ceci !
mv dist/CorrectionApp dist/CorrectionApp.exe  # Juste change le nom
```

C'est pas un vrai .exe, faut compiler sur Windows !

---

## 📊 Comparaison des méthodes

| Méthode | Difficulté | Coût | Vitesse | Automatisation |
|---------|-----------|------|--------|---|
| GitHub Actions | ⭐ Facile | 💚 Gratuit | ⚡ 3 min | ✅ Auto |
| PowerShell (Windows) | ⭐ Facile | 💚 Gratuit | ⚡ 3 min | ❌ Manuel |
| WSL2 | ⭐⭐ Moyen | 💚 Gratuit | ⚡ 3 min | ❌ Manuel |
| VM Windows | ⭐⭐⭐ Complexe | 💛 Libre | ⚡ 3 min | ❌ Manuel |
| Docker Windows | ⭐⭐⭐ Complexe | 💛 Libre | ⚠️ Lent | ✅ Auto |

---

## 🎬 Prochaines étapes

1. **Créer un compte GitHub** (si pas encore)
2. **Pousser le code** vers GitHub
3. **Vérifier les Actions** → Attendre le build
4. **Télécharger le .exe** 🎉

La documentation `.github/workflows/build-windows.yml` est **déjà prête** !

Besoin d'aide pour configurer GitHub ?
