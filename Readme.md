# Analyseur de Nombres Excel

Ce script Python analyse un fichier Excel (.xls ou .xlsx) contenant une seule colonne de nombres. Il identifie les plages de nombres manquants et compte les occurrences de chaque nombre.

## Fonctionnalités

- Lit un fichier Excel (.xls) à colonne unique
- Trie les nombres par ordre croissant
- Identifie les plages de nombres manquants à partir de 1
- Compte les occurrences de chaque nombre
- Prend en charge la fonctionnalité glisser-déposer pour une saisie facile des fichiers

## Prérequis

Avant d'exécuter ce script, vous devez avoir Python et pip (installateur de paquets Python) installés sur votre système.

### Installation de Python

1. Visitez le site officiel de Python : https://www.python.org/downloads/
2. Téléchargez la dernière version de Python pour votre système d'exploitation
3. Exécutez l'installateur et suivez l'assistant d'installation
   - Assurez-vous de cocher la case "Ajouter Python au PATH" pendant l'installation
4. Pour vérifier l'installation, ouvrez une invite de commande ou un terminal et tapez :

`python --version`

### Installation des bibliothèques requises

Ce script nécessite plusieurs bibliothèques Python. Installez-les à l'aide de pip :

1. Ouvrez une invite de commande ou un terminal
2. Exécutez la commande suivante :

`pip install pandas openpyxl numpy`

## Installation

1. Clonez ce dépôt ou téléchargez le fichier script :

git clone https://github.com/anas1412/process_excel.git

Ou téléchargez directement `process_excel.py`.

2. Naviguez vers le répertoire du projet :

`cd process_excel`

## Utilisation

Il y a deux façons d'utiliser ce script :

### Méthode 1 : Glisser-déposer

1. Localisez le fichier `process_excel.py` dans votre explorateur de fichiers
2. Faites glisser et déposez votre fichier Excel sur le fichier `process_excel.py`
3. Le script s'exécutera automatiquement, traitera votre fichier et affichera les résultats
4. Appuyez sur Entrée pour fermer la fenêtre de console lorsque vous avez terminé d'examiner les résultats

### Méthode 2 : Exécution depuis la ligne de commande

1. Ouvrez une invite de commande ou un terminal
2. Naviguez vers le répertoire contenant `process_excel.py`
3. Exécutez le script :

`python process_excel.py`

4. Lorsque vous y êtes invité, entrez le chemin complet de votre fichier Excel ou faites glisser et déposez le fichier dans la fenêtre de console
5. Appuyez sur Entrée pour traiter le fichier
6. Examinez les résultats et appuyez à nouveau sur Entrée pour quitter

## Format du fichier d'entrée

Le fichier Excel d'entrée doit :

- Être au format .xls
- Contenir une seule colonne de nombres
- Ne pas avoir de ligne d'en-tête

## Sortie

Le script affichera :

1. Les plages de nombres manquants (à partir de 1)
2. Les occurrences de chaque nombre présent dans le fichier

## Dépannage

- Si vous rencontrez une erreur "Fichier non trouvé", assurez-vous de fournir le bon chemin de fichier
- Pour toute erreur d'importation, assurez-vous d'avoir installé toutes les bibliothèques requises
- Si le script se ferme immédiatement, essayez de l'exécuter à partir de la ligne de commande pour une meilleure visibilité des erreurs

## Contribution

Les contributions, les problèmes et les demandes de fonctionnalités sont les bienvenus. N'hésitez pas à consulter la [page des problèmes](https://github.com/votrenomdutilisateur/analyseur-nombres-excel/issues) si vous souhaitez contribuer.

## Licence

[MIT](https://choosealicense.com/licenses/mit/)
