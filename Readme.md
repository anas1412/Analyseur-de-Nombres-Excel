# Analyseur de Nombres Excel

Ce programme analyse un fichier Excel (.xls ou .xlsx) contenant une seule colonne de nombres. Il identifie les plages de nombres manquants et compte les occurrences de chaque nombre.

## Fonctionnalités

- Lit un fichier Excel (.xls ou .xlsx) à colonne unique
- Trie les nombres par ordre croissant
- Identifie les plages de nombres manquants à partir de 1
- Compte les occurrences de chaque nombre
- Interface graphique simple et intuitive

## Téléchargement et Installation

### Option 1 : Télécharger l'exécutable

1. Allez sur la page des [Releases](https://github.com/anas1412/process_excel/releases) de ce projet.
2. Téléchargez la dernière version de `Excel.Number.Analyzer.exe`.
3. Double-cliquez sur le fichier téléchargé pour exécuter le programme.

### Option 2 : Construire à partir du code source

Si vous préférez construire l'application vous-même :

1. Assurez-vous d'avoir Python 3.7 ou supérieur installé.
2. Clonez ce dépôt :
   `git clone https://github.com/anas1412/process_excel.git`
3. Naviguez vers le répertoire du projet :
   `cd process_excel`
4. Installez les dépendances :
   `pip install pandas openpyxl numpy PyQt5 pyinstaller`
5. Construisez l'exécutable :
   `pyinstaller excel_analyzer.spec`
6. L'exécutable sera créé dans le dossier `dist`.

## Utilisation

1. Lancez l'application en double-cliquant sur `Excel Number Analyzer.exe`.
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

Les contributions, les problèmes et les demandes de fonctionnalités sont les bienvenus. N'hésitez pas à consulter la [page des problèmes](https://github.com/anas1412/process_excel/issues) si vous souhaitez contribuer.

## Licence

[MIT](https://choosealicense.com/licenses/mit/)
