import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import openpyxl
from itertools import groupby
from operator import itemgetter
from collections import Counter

__version__ = "2.0.2"

class ExcelAnalyzerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f'Analyseur de Nombres Excel v{__version__}')
        self.geometry('750x550')

        # Set window icon
        icon_path = os.path.join(os.path.dirname(sys.argv[0]), 'icon.png')
        if os.path.exists(icon_path):
            try:
                self.iconphoto(True, tk.PhotoImage(file=icon_path))
            except tk.TclError:
                print(f"Error loading icon: {icon_path}. Ensure it's a GIF or PNG compatible with Tkinter.")
        else:
            print(f"Icon not found at {icon_path}")

        self.create_widgets()
        self.update_status("Prêt. Veuillez sélectionner un fichier Excel.")

    def create_widgets(self):
        # --- File Selection --- 
        file_frame = ttk.Frame(self, padding="10")
        file_frame.grid(row=0, column=0, columnspan=3, sticky="ew")

        self.select_button = ttk.Button(file_frame, text='📂 Sélectionner un fichier Excel', command=self.select_file)
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.selected_file_label = ttk.Label(file_frame, text="Aucun fichier sélectionné.")
        self.selected_file_label.pack(side=tk.LEFT, padx=5)

        # --- Results Area --- 
        results_label = ttk.Label(self, text="📊 Résultats de l'analyse:", font=('Arial', 11, 'bold'))
        results_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=5)

        self.result_text = tk.Text(self, wrap=tk.WORD, font=('Courier New', 10))
        self.result_text.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=10, pady=5)
        self.result_text.config(state=tk.DISABLED) # Make it read-only

        # --- Action Buttons --- 
        buttons_frame = ttk.Frame(self, padding="10")
        buttons_frame.grid(row=3, column=0, columnspan=3, sticky="ew")

        self.export_button = ttk.Button(buttons_frame, text='💾 Exporter en .txt', command=self.export_results, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.copy_button = ttk.Button(buttons_frame, text='📋 Copier les résultats', command=self.copy_results, state=tk.DISABLED)
        self.copy_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = ttk.Button(buttons_frame, text='🧹 Effacer', command=self.clear_results)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # --- Status Bar --- 
        self.status_bar = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=4, column=0, columnspan=3, sticky="ew")

        # Configure grid weights for responsiveness
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)

    def update_status(self, message, duration=0):
        self.status_bar.config(text=message)
        if duration > 0:
            self.after(duration, lambda: self.status_bar.config(text=""))

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xls *.xlsx")]
        )
        if file_path:
            self.selected_file_label.config(text=f"Sélectionné: {os.path.basename(file_path)}")
            self.update_status(f"Traitement de {os.path.basename(file_path)}...")
            self.process_excel(file_path)
            self.copy_button.config(state=tk.NORMAL)
            self.export_button.config(state=tk.NORMAL)
        else:
            self.selected_file_label.config(text="Aucun fichier sélectionné.")
            self.update_status("Sélection de fichier annulée.")
            self.copy_button.config(state=tk.DISABLED)
            self.export_button.config(state=tk.DISABLED)

    def process_excel(self, file_path):
        numbers = []
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        try:
                            num = int(cell.value)
                            numbers.append(num)
                        except (ValueError, TypeError):
                            pass

            if not numbers:
                self.display_results("<p>Aucune donnée numérique trouvée dans le fichier Excel.</p>")
                self.update_status("Aucune donnée numérique trouvée.", 5000)
                self.copy_button.config(state=tk.DISABLED)
                self.export_button.config(state=tk.DISABLED)
                return

            numbers.sort()

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
            # result_html += "<br>" # Removed this line

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
                    for num in sorted(occurrences_gt_1.keys()):
                         count = occurrences_gt_1[num]
                         occurrences_list.append(f"Numéro {num}: {count} fois")
                    result_html += "<pre>" + "\n".join(occurrences_list) + "</pre>"

            self.display_results(result_html)
            self.update_status("Analyse terminée.", 5000)
            self.copy_button.config(state=tk.NORMAL)
            self.export_button.config(state=tk.NORMAL)

        except FileNotFoundError:
            self.display_results(f"<p style='color: red;'>Erreur: Fichier non trouvé à {file_path}</p>")
            self.update_status("Erreur: Fichier non trouvé.", 5000)
            self.copy_button.config(state=tk.DISABLED)
            self.export_button.config(state=tk.DISABLED)
        except openpyxl.utils.exceptions.InvalidFileException:
             self.display_results("<p style='color: red;'>Erreur: Le fichier sélectionné n'est pas un fichier Excel valide (.xlsx).</p>")
             self.update_status("Erreur: Fichier Excel invalide.", 5000)
             self.copy_button.config(state=tk.DISABLED)
             self.export_button.config(state=tk.DISABLED)
        except Exception as e:
            self.display_results(f"<p style='color: red;'>Une erreur inattendue s'est produite lors de la lecture du fichier: {str(e)}</p>")
            self.update_status(f"Erreur de lecture du fichier: {str(e)}", 5000)
            self.copy_button.config(state=tk.DISABLED)
            self.export_button.config(state=tk.DISABLED)

    def display_results(self, html_content):
        self.result_text.config(state=tk.NORMAL) # Enable editing to insert text
        self.result_text.delete(1.0, tk.END)

        # Define Tkinter tags for styling
        self.result_text.tag_configure('h3', font=('Arial', 12, 'bold'))
        self.result_text.tag_configure('h2', font=('Arial', 14, 'bold'))
        self.result_text.tag_configure('pre', font=('Courier New', 10))
        self.result_text.tag_configure('error', foreground='red')
        self.result_text.tag_configure('italic', font=('Arial', 10, 'italic'))

        # Simple HTML parsing and insertion
        import re

        # Replace <hr> with a line of dashes
        html_content = html_content.replace('<hr>', '----------------------------------------------------\n')

        # Process content within tags
        segments = re.split(r'(<h[23]>.*?</h[23]>|<pre>.*?</pre>|<p style=\'color: red;\'>.*?</p>|<p><i>.*?</i></p>)', html_content, flags=re.DOTALL)

        for segment in segments:
            if not segment.strip():
                continue

            if segment.startswith('<h3>') and segment.endswith('</h3>'):
                text = segment[4:-5].strip()
                self.result_text.insert(tk.END, text + '\n\n', 'h3') # Add an extra newline
            elif segment.startswith('<h2>') and segment.endswith('</h2>'):
                text = segment[4:-5].strip()
                self.result_text.insert(tk.END, '\n' + text + '\n\n', 'h2') # Add an extra newline
            elif segment.startswith('<pre>') and segment.endswith('</pre>'):
                text = segment[5:-6].strip()
                self.result_text.insert(tk.END, text + '\n\n', 'pre') # Add an extra newline
            elif segment.startswith('<p style=\'color: red;\'>') and segment.endswith('</p>'):
                text = re.sub(r'<p style=\'color: red;\'>|</p>', '', segment).strip()
                self.result_text.insert(tk.END, text + '\n\n', 'error') # Add an extra newline
            elif segment.startswith('<p><i>') and segment.endswith('</i></p>'):
                text = re.sub(r'<p><i>|</i></p>', '', segment).strip()
                self.result_text.insert(tk.END, text + '\n\n', 'italic') # Add an extra newline
            elif segment.startswith('<p>') and segment.endswith('</p>'):
                text = re.sub(r'<p>|</p>', '', segment).strip()
                self.result_text.insert(tk.END, text + '\n\n') # Add an extra newline
            else:
                self.result_text.insert(tk.END, segment + '\n\n') # Add an extra newline

        self.result_text.config(state=tk.DISABLED) # Make it read-only again

    def export_results(self):
        # This function will export the plain text content of the result_text widget
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Fichiers texte", "*.txt")],
            title="Exporter les résultats"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.result_text.get(1.0, tk.END))
                self.update_status(f"Résultats exportés vers {os.path.basename(file_path)}", 5000)
            except Exception as e:
                messagebox.showerror("Erreur d'exportation", f"Impossible d'exporter les résultats: {e}")
                self.update_status("Erreur lors de l'exportation.", 5000)

    def copy_results(self):
        # This function will copy the plain text content of the result_text widget to clipboard
        try:
            self.clipboard_clear()
            self.clipboard_append(self.result_text.get(1.0, tk.END))
            self.update_status("Résultats copiés dans le presse-papiers.", 5000)
        except Exception as e:
            messagebox.showerror("Erreur de copie", f"Impossible de copier les résultats: {e}")
            self.update_status("Erreur lors de la copie.", 5000)

    def clear_results(self):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state=tk.DISABLED)
        self.selected_file_label.config(text="Aucun fichier sélectionné.")
        self.update_status("Prêt. Veuillez sélectionner un fichier Excel.")
        self.copy_button.config(state=tk.DISABLED)
        self.export_button.config(state=tk.DISABLED)

if __name__ == '__main__':
    app = ExcelAnalyzerGUI()
    app.mainloop()