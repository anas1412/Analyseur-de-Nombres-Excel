# -*- mode: python ; coding: utf-8 -*-

__version__ = "2.0.1"

import sys
import os
# import pandas as pd # Removed pandas
# import numpy as np # Removed numpy
import openpyxl # Added openpyxl
from itertools import groupby
from operator import itemgetter
from collections import Counter # Added Counter for occurrences
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit,
    # QVBoxLayout, # Removed unused import
    QWidget, QLabel, QGridLayout, QStatusBar, QHBoxLayout
)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt

class ExcelAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle(f'Analyseur de Nombres Excel v{__version__}')
        self.setGeometry(100, 100, 750, 550)

        # Set window icon using icon.png
        icon_path = os.path.join(os.path.dirname(sys.argv[0]), 'icon.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"Icon not found at {icon_path}") # Fallback message

        # Main widget and layout
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        grid_layout = QGridLayout(main_widget)

        # --- File Selection ---
        self.selectButton = QPushButton('📂 Sélectionner un fichier Excel', self)
        self.selectButton.setFont(QFont('Arial', 10))
        self.selectButton.clicked.connect(self.selectFile)
        grid_layout.addWidget(self.selectButton, 0, 0, 1, 1)

        self.selectedFileLabel = QLabel("Aucun fichier sélectionné.", self)
        self.selectedFileLabel.setFont(QFont('Arial', 9))
        grid_layout.addWidget(self.selectedFileLabel, 0, 1, 1, 2)

        # --- Results Area ---
        results_label = QLabel("📊 Résultats de l'analyse:", self)
        results_label.setFont(QFont('Arial', 11, QFont.Bold))
        grid_layout.addWidget(results_label, 1, 0, 1, 3)

        self.resultText = QTextEdit(self)
        self.resultText.setReadOnly(True)
        self.resultText.setFont(QFont('Courier New', 10))
        grid_layout.addWidget(self.resultText, 2, 0, 1, 3)

        # --- Action Buttons ---
        buttons_layout = QHBoxLayout()

        self.exportButton = QPushButton('💾 Exporter en .txt', self)
        self.exportButton.setFont(QFont('Arial', 10))
        self.exportButton.clicked.connect(self.exportResults)
        self.exportButton.setEnabled(False)
        buttons_layout.addWidget(self.exportButton)

        self.copyButton = QPushButton('📋 Copier les résultats', self)
        self.copyButton.setFont(QFont('Arial', 10))
        self.copyButton.clicked.connect(self.copyResults)
        self.copyButton.setEnabled(False)
        buttons_layout.addWidget(self.copyButton)

        self.clearButton = QPushButton('🧹 Effacer', self)
        self.clearButton.setFont(QFont('Arial', 10))
        self.clearButton.clicked.connect(self.clearResults)
        buttons_layout.addWidget(self.clearButton)

        grid_layout.addLayout(buttons_layout, 3, 0, 1, 3)

        # --- Status Bar ---
        self.statusBar = QStatusBar(self)
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Prêt. Veuillez sélectionner un fichier Excel.")

        # Set column stretch factors for responsiveness
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 2)
        grid_layout.setColumnStretch(2, 1)
        grid_layout.setRowStretch(2, 1)

    def selectFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner un fichier Excel", "", "Fichiers Excel (*.xls *.xlsx)")
        if fileName:
            self.selectedFileLabel.setText(f"Sélectionné: {os.path.basename(fileName)}")
            self.statusBar.showMessage(f"Traitement de {os.path.basename(fileName)}...")
            QApplication.processEvents()
            self.processExcel(fileName)
            self.copyButton.setEnabled(True)
            self.exportButton.setEnabled(True)
        else:
            self.selectedFileLabel.setText("Aucun fichier sélectionné.")
            self.statusBar.showMessage("Sélection de fichier annulée.")
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)

    def processExcel(self, file_path):
        numbers = []
        try:
            # Use openpyxl to read data
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        try:
                            # Attempt to convert to integer
                            num = int(cell.value)
                            numbers.append(num)
                        except (ValueError, TypeError):
                            # Ignore non-integer values
                            pass

            if not numbers:
                self.resultText.setHtml("<p>Aucune donnée numérique trouvée dans le fichier Excel.</p>")
                self.statusBar.showMessage("Aucune donnée numérique trouvée.", 5000)
                self.copyButton.setEnabled(False)
                self.exportButton.setEnabled(False)
                return

            numbers.sort()

            # Start building HTML result string
            result_html = f"<h3>Analyse du fichier: {os.path.basename(file_path)}</h3>"
            result_html += "<hr>"

            # --- Missing Numbers ---
            result_html += "<h2>🔢 Numéros manquants:</h2>"
            if not numbers:
                result_html += "<p><i>(Aucun nombre trouvé pour analyser les plages manquantes)</i></p>"
            else:
                max_number = max(numbers)
                all_present_numbers = set(range(1, max_number + 1))
                missing_numbers_list = sorted(list(all_present_numbers - set(numbers)))

                if not missing_numbers_list:
                    result_html += "<p>Aucun numéro manquant (jusqu'au nombre maximum trouvé).</p>"
                else:
                    missing_ranges_formatted = []
                    for k, g in groupby(enumerate(missing_numbers_list), lambda x: x[0] - x[1]):
                        group = list(map(itemgetter(1), g))
                        if group[0] == group[-1]:
                            missing_ranges_formatted.append(f"{group[0]}")
                        else:
                            missing_ranges_formatted.append(" ".join(map(str, range(group[0], group[-1] + 1))))
                    result_html += "<pre>" + "\n".join(missing_ranges_formatted) + "</pre>"
            result_html += "<br>"

            # --- Occurrences ---
            result_html += "<h2>📊 Occurrences des numéros (Plus d'une fois):</h2>"
            if not numbers:
                result_html += "<p><i>(Aucun nombre trouvé pour compter les occurrences)</i></p>"
            else:
                occurrences = Counter(numbers)
                occurrences_gt_1 = {num: count for num, count in occurrences.items() if count > 1}
                
                if not occurrences_gt_1:
                    result_html += "<p>Aucun numéro n'apparaît plus d'une fois.</p>"
                else:
                    occurrences_list = []
                    # Sort occurrences by number for consistent output
                    for num in sorted(occurrences_gt_1.keys()):
                         count = occurrences_gt_1[num]
                         occurrences_list.append(f"Numéro {num}: {count} fois")
                    result_html += "<pre>" + "\n".join(occurrences_list) + "</pre>"

            self.resultText.setHtml(result_html)
            self.statusBar.showMessage("Analyse terminée.", 5000)
            self.copyButton.setEnabled(True)
            self.exportButton.setEnabled(True)

        except FileNotFoundError:
            self.resultText.setHtml(f"<p style='color: red;'>Erreur: Fichier non trouvé à {file_path}</p>")
            self.statusBar.showMessage("Erreur: Fichier non trouvé.", 5000)
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)
        except openpyxl.utils.exceptions.InvalidFileException:
             self.resultText.setHtml("<p style='color: red;'>Erreur: Le fichier sélectionné n'est pas un fichier Excel valide (.xlsx).</p>")
             self.statusBar.showMessage("Erreur: Fichier Excel invalide.", 5000)
             self.copyButton.setEnabled(False)
             self.exportButton.setEnabled(False)
        except Exception as e:
            self.resultText.setHtml(f"<p style='color: red;'>Une erreur inattendue s'est produite lors de la lecture du fichier: {str(e)}</p>")
            self.statusBar.showMessage(f"Erreur de lecture du fichier: {str(e)}", 5000)
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)

    def exportResults(self):
        if not self.resultText.toPlainText():
            self.statusBar.showMessage("Aucun résultat à exporter.", 3000)
            return

        fileName, _ = QFileDialog.getSaveFileName(self, "Exporter les résultats en .txt", "resultats_analyse_excel.txt", "Fichiers Texte (*.txt);;Tous les fichiers (*)")
        if fileName:
            try:
                with open(fileName, 'w', encoding='utf-8') as f:
                    f.write(self.resultText.toPlainText())
                self.statusBar.showMessage(f"Résultats exportés vers {fileName}", 3000)
            except Exception as e:
                self.statusBar.showMessage(f"Erreur lors de l'exportation: {str(e)}", 5000)

    def copyResults(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.resultText.toPlainText())
        self.statusBar.showMessage("Résultats copiés dans le presse-papiers!", 3000)

    def clearResults(self):
        self.resultText.clear()
        self.selectedFileLabel.setText("Aucun fichier sélectionné.")
        self.statusBar.showMessage("Prêt. Veuillez sélectionner un fichier Excel.")
        self.copyButton.setEnabled(False)
        self.exportButton.setEnabled(False)

def main():
    app = QApplication(sys.argv)
    ex = ExcelAnalyzerGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
