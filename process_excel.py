import pandas as pd
import numpy as np
from itertools import groupby
from operator import itemgetter
import sys
import os

def process_excel(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path, header=None, names=['Numbers'])

    # Convert to integers and sort
    numbers = sorted(df['Numbers'].astype(int))

    # Find missing ranges
    all_numbers = set(range(1, max(numbers) + 1))
    missing = sorted(all_numbers - set(numbers))

    # Create ranges of missing numbers
    missing_ranges = []
    for k, g in groupby(enumerate(missing), lambda x: x[0] - x[1]):
        group = list(map(itemgetter(1), g))
        missing_ranges.append((group[0], group[-1]))

    # Count occurrences
    occurrences = pd.Series(numbers).value_counts().sort_index()

    # Print results
    print("Plages de numéros manquantes:")
    for start, end in missing_ranges:
        if start == end:
            print(f"{start}")
        else:
            print(" ".join(map(str, range(start, end + 1))))

    print("\nLes occurrences:")
    print(occurrences)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Veuillez glisser et déposer votre fichier Excel ici et appuyez sur Entrée: ").strip()

    # Remove quotes if present (some systems add quotes when dragging files)
    file_path = file_path.strip("\"'")

    if not os.path.exists(file_path):
        print(f"Erreur : le fichier '{file_path}' n'a pas été trouvé.")
    else:
        process_excel(file_path)

    # Keep console open
    input("\nAppuyez sur Entrée pour quitter...")