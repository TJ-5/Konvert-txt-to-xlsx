import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget,
    QPushButton, QTableView, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QKeySequence
from PyQt5.QtCore import QAbstractTableModel
from PyQt5.QtWidgets import QShortcut

class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._df.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._df.columns[section])
            else:
                return str(section)
        return None

    def removeRows(self, row_indices):
        self.beginResetModel()
        self._df = self._df.drop(index=row_indices).reset_index(drop=True)
        self.endResetModel()

    def get_dataframe(self):
        return self._df

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GEMA TXT → Excel Konverter mit Vorschau")
        self.setAcceptDrops(True)

        self.df = pd.DataFrame()
        self.model = None
        self.loaded_filepath = None

        layout = QVBoxLayout()
        self.table = QTableView()
        layout.addWidget(self.table)

        self.delete_button = QPushButton("Ausgewählte Zeilen löschen (Entf)")
        self.delete_button.clicked.connect(self.delete_selected_rows)
        layout.addWidget(self.delete_button)

        self.export_button = QPushButton("Exportieren als Excel (Enter)")
        self.export_button.clicked.connect(self.export_to_excel)
        layout.addWidget(self.export_button)

        # Neuer Button zum manuellen Zurücksetzen
        self.reset_button = QPushButton("Daten löschen / Neue Datei laden")
        self.reset_button.clicked.connect(self.reset_data)
        layout.addWidget(self.reset_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Initial state
        self.update_button_states()

        # Keyboard Shortcuts
        # Enter für Export
        self.export_shortcut = QShortcut(QKeySequence("Return"), self)
        self.export_shortcut.activated.connect(self.export_to_excel)
        
        # Numpad Enter für Export
        self.export_shortcut_numpad = QShortcut(QKeySequence("Enter"), self)
        self.export_shortcut_numpad.activated.connect(self.export_to_excel)
        
        # Delete und Entf für Zeilen löschen
        self.delete_shortcut = QShortcut(QKeySequence("Delete"), self)
        self.delete_shortcut.activated.connect(self.delete_selected_rows)
        
        # Backspace als Alternative
        self.delete_shortcut_backspace = QShortcut(QKeySequence("Backspace"), self)
        self.delete_shortcut_backspace.activated.connect(self.delete_selected_rows)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            self.load_file(file_path)

    def load_file(self, filepath):
        try:
            self.loaded_filepath = filepath

            data_start = self.find_data_start(filepath)
            if data_start is None:
                raise Exception("Konnte 'Assemble List' nicht finden.")

            df = pd.read_csv(filepath, sep="\t", engine="python", skiprows=data_start)

            df = df[df["MasDur"] != "MasDur"]
            df = df.dropna(subset=["Clip"])
            df = df.sort_values(by="Clip")

            self.df = df[["MasDur", "Clip"]].reset_index(drop=True)

            self.model = PandasModel(self.df)
            self.table.setModel(self.model)
            self.table.setSelectionBehavior(self.table.SelectRows)
            self.table.setSelectionMode(self.table.ExtendedSelection)

            # Update button states and window title
            self.update_button_states()
            filename = os.path.basename(filepath)
            self.setWindowTitle(f"GEMA TXT → Excel Konverter - {filename}")

        except Exception as e:
            QMessageBox.critical(self, "Fehler", str(e))

    def find_data_start(self, filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f):
                if "Assemble List" in line:
                    return i + 2
        return None

    def delete_selected_rows(self):
        # Prüfen ob Tabelle fokussiert ist und Zeilen ausgewählt sind
        if not self.table.hasFocus() and not self.has_selected_rows():
            return
            
        selected_indexes = self.table.selectionModel().selectedRows()
        row_indices = sorted([index.row() for index in selected_indexes], reverse=True)
        if self.model and row_indices:
            self.model.removeRows(row_indices)
    
    def has_selected_rows(self):
        """Prüft ob Zeilen ausgewählt sind"""
        if not self.table.selectionModel():
            return False
        return len(self.table.selectionModel().selectedRows()) > 0

    def export_to_excel(self):
        # Prüfen ob Daten vorhanden sind, bevor exportiert wird
        if not self.model or self.model.rowCount() == 0:
            QMessageBox.warning(self, "Hinweis", "Keine Daten zum Exportieren.")
            return
            
        if self.model:
            df = self.model.get_dataframe()

            if self.loaded_filepath:
                base_dir = os.path.dirname(self.loaded_filepath)
                base_name = os.path.splitext(os.path.basename(self.loaded_filepath))[0]
                export_path = os.path.join(base_dir, base_name + ".xlsx")

                try:
                    df.to_excel(export_path, index=False)
                    QMessageBox.information(self, "Gespeichert", f"Gespeichert unter:\n{export_path}")
                    
                    # Nach erfolgreichem Export: Daten zurücksetzen
                    self.reset_data()
                    
                except Exception as e:
                    QMessageBox.critical(self, "Fehler beim Speichern", str(e))
            else:
                QMessageBox.warning(self, "Fehler", "Kein Pfad zur Quelldatei bekannt.")

    def reset_data(self):
        """Setzt alle Daten zurück und bereitet das Programm für eine neue Datei vor"""
        self.df = pd.DataFrame()
        self.model = None
        self.loaded_filepath = None
        
        # Tabelle leeren
        empty_model = PandasModel(pd.DataFrame())
        self.table.setModel(empty_model)
        
        # Button-Zustände aktualisieren
        self.update_button_states()
        
        # Fenstertitel zurücksetzen
        self.setWindowTitle("GEMA TXT → Excel Konverter mit Vorschau")

    def update_button_states(self):
        """Aktualisiert den Zustand der Buttons basierend auf geladenen Daten"""
        has_data = self.model is not None and self.model.rowCount() > 0
        
        self.delete_button.setEnabled(has_data)
        self.export_button.setEnabled(has_data)
        self.reset_button.setEnabled(has_data)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec_())