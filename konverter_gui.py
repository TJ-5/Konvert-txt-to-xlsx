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
        self.resize(800, 600)  # Größeres Startfenster

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

        self.reset_button = QPushButton("Daten löschen / Neue Datei laden")
        self.reset_button.clicked.connect(self.reset_data)
        layout.addWidget(self.reset_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.update_button_states()

        # Keyboard Shortcuts
        self.export_shortcut = QShortcut(QKeySequence("Return"), self)
        self.export_shortcut.activated.connect(self.export_to_excel)
        
        self.export_shortcut_numpad = QShortcut(QKeySequence("Enter"), self)
        self.export_shortcut_numpad.activated.connect(self.export_to_excel)
        
        self.delete_shortcut = QShortcut(QKeySequence("Delete"), self)
        self.delete_shortcut.activated.connect(self.delete_selected_rows)
        
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
            
            # Automatische Spaltenbreite für Clip-Spalte
            self.table.resizeColumnToContents(1)
            if self.table.columnWidth(1) > 500:
                self.table.setColumnWidth(1, 500)  # Maximalbreite

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
        if not self.table.hasFocus() and not self.has_selected_rows():
            return
            
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            return
            
        # Position vor dem Löschen merken
        row_indices = sorted([index.row() for index in selected_indexes])
        first_selected = row_indices[0]
        
        # Zeilen löschen
        self.model.removeRows(sorted(row_indices, reverse=True))
        
        # Position nach dem Löschen wiederherstellen
        if self.model.rowCount() > 0:
            new_row = min(first_selected, self.model.rowCount() - 1)
            self.table.selectRow(new_row)
            self.table.scrollTo(self.model.index(new_row, 0))
        
        self.update_button_states()
    
    def has_selected_rows(self):
        if not self.table.selectionModel():
            return False
        return len(self.table.selectionModel().selectedRows()) > 0

    def export_to_excel(self):
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
                    self.reset_data()
                except Exception as e:
                    QMessageBox.critical(self, "Fehler beim Speichern", str(e))
            else:
                QMessageBox.warning(self, "Fehler", "Kein Pfad zur Quelldatei bekannt.")

    def reset_data(self):
        self.df = pd.DataFrame()
        self.model = None
        self.loaded_filepath = None
        
        empty_model = PandasModel(pd.DataFrame())
        self.table.setModel(empty_model)
        
        self.update_button_states()
        self.setWindowTitle("GEMA TXT → Excel Konverter mit Vorschau")

    def update_button_states(self):
        has_data = self.model is not None and self.model.rowCount() > 0
        
        self.delete_button.setEnabled(has_data)
        self.export_button.setEnabled(has_data)
        self.reset_button.setEnabled(has_data)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())