# build_exe_windows.ps1 - Script PowerShell pour Windows

Write-Host "🔨 Construction de l'exécutable Windows..." -ForegroundColor Cyan

# Vérifier si Python est installé
try {
    $python = python --version
    Write-Host "✓ Python trouvé: $python" -ForegroundColor Green
} catch {
    Write-Host "❌ Python n'est pas installé ou pas dans PATH" -ForegroundColor Red
    Write-Host "Installer Python depuis https://www.python.org/" -ForegroundColor Yellow
    exit 1
}

# Installer PyInstaller si absent
Write-Host "`n📦 Vérification de PyInstaller..." -ForegroundColor Cyan
pip list | Select-String PyInstaller | Out-Null
if ($?) {
    Write-Host "✓ PyInstaller est installé" -ForegroundColor Green
} else {
    Write-Host "Installation de PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
}

# Installer les dépendances
Write-Host "`n📦 Installation des dépendances..." -ForegroundColor Cyan
pip install -r requirements.txt --quiet

# Compiler
Write-Host "`n🔨 Compilation de l'exécutable..." -ForegroundColor Cyan
python -m PyInstaller `
    --name=CorrectionApp `
    --onefile `
    --windowed `
    --hidden-import=PyPDF2 `
    --hidden-import=pdfplumber `
    --hidden-import=pdf2image `
    --hidden-import=PIL `
    --hidden-import=openpyxl `
    --hidden-import=odf `
    --hidden-import=correction `
    --collect-all=PyQt5 `
    app_correction.py

# Vérifier le résultat
if (Test-Path "dist/CorrectionApp.exe") {
    Write-Host "`n✅ Exécutable créé avec succès !" -ForegroundColor Green
    Write-Host "📍 Localisation: $(Resolve-Path 'dist/CorrectionApp.exe')" -ForegroundColor Green
    Write-Host "`n📋 Taille: $((Get-Item 'dist/CorrectionApp.exe').Length / 1MB -as [int]) MB" -ForegroundColor Cyan
    Write-Host "`n💡 Prêt pour distribution !" -ForegroundColor Green
} else {
    Write-Host "`n❌ Erreur: le fichier .exe n'a pas été créé" -ForegroundColor Red
    exit 1
}
