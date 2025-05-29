# Analyseur de Nombres Excel

[![Download Latest Release](https://img.shields.io/github/v/release/anas1412/Analyseur-de-Nombres-Excel?label=Download%20Latest)](https://github.com/anas1412/Analyseur-de-Nombres-Excel/releases/latest/download/Analyseur.de.Nombres.Excel.exe)

Ce programme analyse un fichier Excel (.xls ou .xlsx) contenant une seule colonne de nombres. Il identifie les plages de nombres manquants et compte les occurrences de chaque nombre.

<div align="center">
  <img src="https://github.com/anas1412/Analyseur-de-Nombres-Excel/raw/main/ENA.png" alt="ENA" />
</div>

## Fonctionnalités

- Lit un fichier Excel (.xls ou .xlsx) à colonne unique
- Trie les nombres par ordre croissant
- Identifie les plages de nombres manquants à partir de 1
- Compte les occurrences de chaque nombre
- Interface graphique simple et intuitive

## Téléchargement et Installation

### Télécharger l'exécutable

Simplement clicker ici: [Télécharger](https://github.com/anas1412/Analyseur-de-Nombres-Excel/releases/latest/download/Analyseur.de.Nombres.Excel.exe)

### Construction à partir du code source

Pour construire l'application à partir du code source, vous aurez besoin de Python 3.7.5 (64 bits). Vous pouvez le télécharger [ici](https://www.python.org/downloads/release/python-375/).

1.  **Cloner le dépôt :**
    ```bash
    git clone https://github.com/your-repo/excel-analyzer.git
    cd excel-analyzer
    ```
2.  **Installer les dépendances :**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Construire l'exécutable (Windows) :**
    L'application utilise maintenant Tkinter pour son interface graphique. Pour construire l'exécutable, utilisez la commande suivante :
    ```bash
    pyinstaller excel_analyzer_gui.spec --clean --upx-dir "c:/upx"
    ```
    L'exécutable, nommé `Analyseur.de.Nombres.Excel.exe`, se trouvera dans le dossier `dist`.

    **Installation d'UPX (Optionnel, pour des exécutables plus petits) :**
    Pour réduire davantage la taille de l'exécutable, vous pouvez utiliser UPX. Téléchargez la dernière version d'UPX pour Windows depuis <mcurl name="Versions d'UPX sur GitHub" url="https://github.com/upx/upx/releases/latest"></mcurl> <mcreference link="https://github.com/upx/upx/releases/latest" index="0">0</mcreference>.
    Extrayez l'archive téléchargée et placez le fichier `upx.exe` dans `c:/upx`. Si le dossier extrait contient `upx.exe` dans un sous-dossier (par exemple, `upx-5.0.1-win64`), déplacez `upx.exe` directement dans `c:/upx`.

## Utilisation

1. Lancez l'application en double-cliquant sur `Analyseur.de.Nombres.Excel.exe`.
2. Cliquez sur "Select Excel File" et choisissez votre fichier Excel.
3. Les résultats s'afficheront dans la fenêtre de l'application.

## Format du fichier d'entrée

Le fichier Excel d'entrée doit :

- Être au format .xls ou .xlsx
- Contenir une seule colonne de nombres
- Ne pas avoir de ligne d'en-tête

## Dépannage

- Si l'application ne démarre pas, assurez-vous que votre antivirus ne la bloque pas.
- Pour tout problème avec le fichier Excel, vérifiez qu'il respecte le format requis.

## Contribution

Les contributions, les problèmes et les demandes de fonctionnalités sont les bienvenus. N'hésitez pas à consulter la [page des problèmes](https://github.com/anas1412/Analyseur-de-Nombres-Excel/issues) si vous souhaitez contribuer.

## Licence

[MIT](https://choosealicense.com/licenses/mit/)
