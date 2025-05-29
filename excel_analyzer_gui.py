__version__ = "2.0.0" # Updated to next major version

import sys
import os
import pandas as pd
# numpy is imported but not used directly in this file, pandas uses it.
# import numpy as np 
from itertools import groupby
from operator import itemgetter
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, 
    QVBoxLayout, QWidget, QLabel, QGridLayout, QStatusBar, QHBoxLayout
)
from PyQt5.QtGui import QFont, QIcon # Re-added QIcon
from PyQt5.QtCore import Qt

class ExcelAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle(f'Analyseur de Nombres Excel v{__version__}') # Emoji removed
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
        self.selectButton = QPushButton('📂 Sélectionner un fichier Excel', self) # French Translation
        self.selectButton.setFont(QFont('Arial', 10))
        self.selectButton.clicked.connect(self.selectFile)
        grid_layout.addWidget(self.selectButton, 0, 0, 1, 1) # row, col, rowspan, colspan

        self.selectedFileLabel = QLabel("Aucun fichier sélectionné.", self) # French Translation
        self.selectedFileLabel.setFont(QFont('Arial', 9))
        grid_layout.addWidget(self.selectedFileLabel, 0, 1, 1, 2) # Span 2 columns
        
        # --- Results Area --- 
        results_label = QLabel("📊 Résultats de l'analyse:", self) # French Translation
        results_label.setFont(QFont('Arial', 11, QFont.Bold))
        grid_layout.addWidget(results_label, 1, 0, 1, 3)

        self.resultText = QTextEdit(self)
        self.resultText.setReadOnly(True)
        self.resultText.setFont(QFont('Courier New', 10))
        grid_layout.addWidget(self.resultText, 2, 0, 1, 3) # Span 3 columns

        # --- Action Buttons --- 
        buttons_layout = QHBoxLayout()

        self.exportButton = QPushButton('💾 Exporter en .txt', self) # New Button & French Translation
        self.exportButton.setFont(QFont('Arial', 10))
        self.exportButton.clicked.connect(self.exportResults)
        self.exportButton.setEnabled(False) # Disabled until results are available
        buttons_layout.addWidget(self.exportButton)

        self.copyButton = QPushButton('📋 Copier les résultats', self) # French Translation
        self.copyButton.setFont(QFont('Arial', 10))
        self.copyButton.clicked.connect(self.copyResults)
        self.copyButton.setEnabled(False) # Disabled until results are available
        buttons_layout.addWidget(self.copyButton)

        self.clearButton = QPushButton('🧹 Effacer', self) # French Translation
        self.clearButton.setFont(QFont('Arial', 10))
        self.clearButton.clicked.connect(self.clearResults)
        buttons_layout.addWidget(self.clearButton)

        grid_layout.addLayout(buttons_layout, 3, 0, 1, 3) # Span 3 columns

        # --- Status Bar --- 
        self.statusBar = QStatusBar(self)
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Prêt. Veuillez sélectionner un fichier Excel.") # French Translation

        # Set column stretch factors for responsiveness
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 2)
        grid_layout.setColumnStretch(2, 1)
        grid_layout.setRowStretch(2, 1) # Allow resultText to expand vertically

    def selectFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Sélectionner un fichier Excel", "", "Fichiers Excel (*.xls *.xlsx)") # French Translation
        if fileName:
            self.selectedFileLabel.setText(f"Sélectionné: {os.path.basename(fileName)}") # French Translation
            self.statusBar.showMessage(f"Traitement de {os.path.basename(fileName)}...") # French Translation
            QApplication.processEvents() # Update UI before long task
            self.processExcel(fileName)
            self.copyButton.setEnabled(True)
            self.exportButton.setEnabled(True) # Enable export button
        else:
            self.selectedFileLabel.setText("Aucun fichier sélectionné.") # French Translation
            self.statusBar.showMessage("Sélection de fichier annulée.") # French Translation
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False) # Disable export button

    def processExcel(self, file_path):
        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.xlsx':
                engine = 'openpyxl'
            elif file_extension == '.xls':
                engine = 'xlrd'
            else:
                self.resultText.setText("Erreur: Format de fichier non supporté. Veuillez utiliser des fichiers .xls ou .xlsx.") # French Translation
                self.statusBar.showMessage("Erreur: Format de fichier non supporté.", 5000) # French Translation
                return

            df = pd.read_excel(file_path, header=None, names=['Numbers'], engine=engine)

            if df.empty or 'Numbers' not in df.columns or df['Numbers'].isnull().all():
                self.resultText.setText("Erreur: Le fichier Excel est vide ou la colonne 'Numbers' est manquante/vide.") # French Translation
                self.statusBar.showMessage("Erreur: Contenu du fichier Excel vide ou invalide.", 5000) # French Translation
                return
            
            try:
                numbers_series = df['Numbers'].dropna().astype(int)
            except ValueError:
                self.resultText.setText("Erreur: La colonne 'Numbers' contient des données non numériques.") # French Translation
                self.statusBar.showMessage("Erreur: Données non numériques dans la colonne 'Numbers'.", 5000) # French Translation
                return

            if numbers_series.empty:
                self.resultText.setText("Aucune donnée numérique trouvée dans la colonne 'Numbers' après nettoyage.") # French Translation
                self.statusBar.showMessage("Aucune donnée numérique trouvée.", 5000) # French Translation
                return

            numbers = sorted(numbers_series)

            if not numbers:
                self.resultText.setText("Aucun nombre trouvé dans le fichier Excel.") # French Translation
                self.statusBar.showMessage("Aucun nombre trouvé dans le fichier.", 5000) # French Translation
                return

            # --- Missing Numbers --- 
            result_str = "🔢 Numéros manquants:\n" # French Translation
            result_str += "--------------------\n"
            if not numbers:
                result_str += "(Aucun nombre trouvé pour analyser les plages manquantes)\n" # French Translation
            else:
                all_present_numbers = set(range(1, max(numbers) + 1))
                missing_numbers_list = sorted(all_present_numbers - set(numbers))

                if not missing_numbers_list:
                    result_str += "Aucun numéro manquant (jusqu'au nombre maximum trouvé).\n" # French Translation
                else:
                    missing_ranges = []
                    for k, g in groupby(enumerate(missing_numbers_list), lambda x: x[0] - x[1]):
                        group = list(map(itemgetter(1), g))
                        missing_ranges.append((group[0], group[-1]))
                    
                    for start, end in missing_ranges:
                        if start == end:
                            result_str += f"{start}\n"
                        else:
                            result_str += " ".join(map(str, range(start, end + 1))) + "\n"
            result_str += "\n"

            # --- Occurrences --- 
            result_str += "📊 Occurrences des numéros (Plus d'une fois):\n" # French Translation
            result_str += "---------------------------------------\n"
            if not numbers:
                result_str += "(Aucun nombre trouvé pour compter les occurrences)\n" # French Translation
            else:
                occurrences = pd.Series(numbers).value_counts()
                occurrences_gt_1 = occurrences[occurrences > 1].sort_index()
                if not occurrences_gt_1.empty:
                    for num, count in occurrences_gt_1.items():
                        result_str += f"Numéro {num}: {count} fois\n" # French Translation
                else:
                    result_str += "Aucun numéro n'apparaît plus d'une fois.\n" # French Translation

            self.resultText.setText(result_str)
            self.statusBar.showMessage("Analyse terminée.", 5000) # French Translation
            self.copyButton.setEnabled(True)
            self.exportButton.setEnabled(True)

        except FileNotFoundError:
            self.resultText.setText(f"Erreur: Fichier non trouvé à {file_path}") # French Translation
            self.statusBar.showMessage("Erreur: Fichier non trouvé.", 5000) # French Translation
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)
        except pd.errors.EmptyDataError:
            self.resultText.setText("Erreur: Le fichier Excel sélectionné est vide.") # French Translation
            self.statusBar.showMessage("Erreur: Le fichier sélectionné est vide.", 5000) # French Translation
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)
        except ValueError as ve:
             self.resultText.setText(f"Erreur lors du traitement du fichier: {str(ve)}") # French Translation
             self.statusBar.showMessage(f"Erreur de valeur: {str(ve)}", 5000) # French Translation
             self.copyButton.setEnabled(False)
             self.exportButton.setEnabled(False)
        except Exception as e:
            self.resultText.setText(f"Une erreur inattendue s'est produite: {str(e)}\nVérifiez si le fichier Excel est valide et non corrompu.") # French Translation
            self.statusBar.showMessage(f"Erreur inattendue: {str(e)}", 5000) # French Translation
            self.copyButton.setEnabled(False)
            self.exportButton.setEnabled(False)

    def exportResults(self):
        if not self.resultText.toPlainText():
            self.statusBar.showMessage("Aucun résultat à exporter.", 3000) # French Translation
            return
        
        fileName, _ = QFileDialog.getSaveFileName(self, "Exporter les résultats en .txt", "resultats_analyse_excel.txt", "Fichiers Texte (*.txt);;Tous les fichiers (*)") # French Translation
        if fileName:
            try:
                with open(fileName, 'w', encoding='utf-8') as f:
                    f.write(self.resultText.toPlainText())
                self.statusBar.showMessage(f"Résultats exportés vers {fileName}", 3000) # French Translation
            except Exception as e:
                self.statusBar.showMessage(f"Erreur lors de l'exportation: {str(e)}", 5000) # French Translation

    def copyResults(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.resultText.toPlainText())
        self.statusBar.showMessage("Résultats copiés dans le presse-papiers!", 3000) # French Translation

    def clearResults(self):
        self.resultText.clear()
        self.selectedFileLabel.setText("Aucun fichier sélectionné.") # French Translation
        self.statusBar.showMessage("Prêt. Veuillez sélectionner un fichier Excel.") # French Translation
        self.copyButton.setEnabled(False)
        self.exportButton.setEnabled(False)

def main():
    app = QApplication(sys.argv)
    # app.setStyle('Fusion') 
    ex = ExcelAnalyzerGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
