__version__ = "1.0.2"

import sys
import os
import pandas as pd
import numpy as np
from itertools import groupby
from operator import itemgetter
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, QWidget

class ExcelAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Number Analyzer')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.selectButton = QPushButton('Sélectionner le fichier Excel', self)
        self.selectButton.clicked.connect(self.selectFile)
        layout.addWidget(self.selectButton)

        self.resultText = QTextEdit(self)
        self.resultText.setReadOnly(True)
        layout.addWidget(self.resultText)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def selectFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner le fichier Excel", "", "Fichiers Excel (*.xls *.xlsx)")
        if fileName:
            self.processExcel(fileName)

    def processExcel(self, file_path):
        try:
            # Determine the file extension and choose the appropriate engine
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.xlsx':
                engine = 'openpyxl'
            elif file_extension == '.xls':
                engine = 'xlrd'
            else:
                raise ValueError("Unsupported file format")

            # Read the Excel file with the specified engine
            df = pd.read_excel(file_path, header=None, names=['Numbers'], engine=engine)

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

            # Filter occurrences greater than 1
            occurrences_gt_1 = occurrences[occurrences > 1]

            # Prepare results
            result = "Plages de numéros manquantes:\n"
            for start, end in missing_ranges:
                if start == end:
                    result += f"{start}\n"
                else:
                    result += f"{start}-{end}\n"

            result += "\nLes occurrences (plus d'une fois):\n"
            if not occurrences_gt_1.empty:
                result += str(occurrences_gt_1)
            else:
                result += "Aucun numéro n'apparaît plus d'une fois."

            self.resultText.setText(result)

        except Exception as e:
            self.resultText.setText(f"Une erreur s'est produite: {str(e)}")

def main():
    app = QApplication(sys.argv)
    ex = ExcelAnalyzerGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
