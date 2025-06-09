import sys
import os
import pandas as pd
import sqlite3
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTextEdit, QPushButton, QVBoxLayout,
    QWidget, QComboBox, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout,
    QMessageBox, QListWidget, QSplitter, QStyleFactory, QLineEdit, QInputDialog,
    QMenu, QListWidgetItem
)
from PyQt5.QtGui import QFont, QColor, QSyntaxHighlighter, QTextCharFormat
from PyQt5.QtCore import Qt, QRegExp

class SQLHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.keyword_format = QTextCharFormat()
        self.keyword_format.setForeground(QColor("blue"))
        self.keyword_format.setFontWeight(QFont.Bold)

        self.keywords = [
            "SELECT", "FROM", "WHERE", "AND", "OR", "INSERT", "INTO", "VALUES",
            "UPDATE", "SET", "DELETE", "CREATE", "TABLE", "DROP", "ALTER",
            "ADD", "JOIN", "ON", "AS", "DISTINCT", "GROUP", "BY", "ORDER", "LIMIT"
        ]

    def highlightBlock(self, text):
        for word in self.keywords:
            expression = QRegExp(f"\\b{word}\\b", Qt.CaseInsensitive)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, self.keyword_format)
                index = expression.indexIn(text, index + length)

class ExcelSQLAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel SQL Analyzer")
        self.setGeometry(100, 100, 1200, 800)

        self.df_dict = {}
        self.conn = sqlite3.connect(":memory:")
        self.history_file = "query_history.json"

        self.init_ui()
        self.load_history()

    def init_ui(self):
        layout = QVBoxLayout()

        theme_button = QPushButton("Cambiar Tema")
        theme_button.clicked.connect(self.toggle_theme)

        file_layout = QHBoxLayout()
        self.load_button = QPushButton("Cargar Excel")
        self.load_button.clicked.connect(self.load_excel)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self.preview_sheet)
        file_layout.addWidget(self.load_button)
        file_layout.addWidget(QLabel("Hoja:"))
        file_layout.addWidget(self.sheet_combo)
        file_layout.addWidget(theme_button)

        self.sql_text = QTextEdit()
        self.sql_text.setFont(QFont("Courier", 10))
        self.highlighter = SQLHighlighter(self.sql_text.document())

        self.query_name_edit = QLineEdit()
        self.query_name_edit.setPlaceholderText("Nombre para guardar la consulta")
        
        button_layout = QHBoxLayout()
        self.run_button = QPushButton("Ejecutar SQL")
        self.run_button.clicked.connect(self.run_query)
        
        self.save_query_button = QPushButton("Guardar Consulta")
        self.save_query_button.clicked.connect(self.save_named_query)
        button_layout.addWidget(self.run_button)
        button_layout.addWidget(self.save_query_button)

        self.table_widget = QTableWidget()

        self.history_list = QListWidget()
        self.history_list.itemDoubleClicked.connect(self.load_from_history)
        self.history_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.history_list.customContextMenuRequested.connect(self.show_history_context_menu)

        sql_area = QSplitter(Qt.Horizontal)
        sql_area.addWidget(self.sql_text)
        sql_area.addWidget(self.history_list)
        sql_area.setSizes([800, 200])

        layout.addLayout(file_layout)
        layout.addWidget(QLabel("Consulta SQL:"))
        layout.addWidget(self.sql_text)
        layout.addWidget(QLabel("Nombre de consulta:"))
        layout.addWidget(self.query_name_edit)
        layout.addLayout(button_layout)
        layout.addWidget(self.table_widget)
        layout.addWidget(QLabel("Historial de consultas:"))
        layout.addWidget(self.history_list)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.dark_mode = False
        self.toggle_theme()

    def toggle_theme(self):
        if self.dark_mode:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            self.setStyleSheet("")
            self.dark_mode = False
        else:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            dark_palette = self.palette()
            dark_palette.setColor(self.backgroundRole(), QColor(53, 53, 53))
            dark_palette.setColor(self.foregroundRole(), QColor(255, 255, 255))
            self.setStyleSheet("""
                QWidget {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QListWidget {
                    background-color: #3c3c3c;
                    color: #ffffff;
                }
                QPushButton {
                    background-color: #444444;
                    border: 1px solid #555555;
                    padding: 5px;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
            """)
            self.dark_mode = True

    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos Excel (*.xlsx *.xls)")
        if file_path:
            self.df_dict = pd.read_excel(file_path, sheet_name=None)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.df_dict.keys())
            self.preview_sheet()

    def preview_sheet(self):
        sheet_name = self.sheet_combo.currentText()
        if sheet_name:
            df = self.df_dict[sheet_name]
            df.to_sql(sheet_name, self.conn, if_exists='replace', index=False)
            self.show_dataframe(df)

    def run_query(self):
        query = self.sql_text.toPlainText().strip()
        if not query:
            return
        try:
            result = pd.read_sql_query(query, self.conn)
            self.show_dataframe(result)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al ejecutar consulta SQL:\n{str(e)}")

    def show_dataframe(self, df):
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)

        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns.astype(str).tolist())

        for row_idx, row in df.iterrows():
            self.table_widget.insertRow(row_idx)
            for col_idx, value in enumerate(row):
                self.table_widget.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))

    def save_named_query(self):
        query = self.sql_text.toPlainText().strip()
        if not query:
            QMessageBox.warning(self, "Advertencia", "No hay consulta para guardar")
            return
            
        name, ok = QInputDialog.getText(self, "Guardar consulta", 
                                       "Introduce un nombre para esta consulta:",
                                       QLineEdit.Normal, self.query_name_edit.text())
        if ok and name:
            self.query_name_edit.setText(name)
            self.save_to_history(query, name)

    def save_to_history(self, query, name=None):
        try:
            history = []
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            
            new_entry = {
                "query": query,
                "name": name if name else f"Consulta {len(history) + 1}",
                "timestamp": pd.Timestamp.now().isoformat()
            }
            
            history.append(new_entry)
            
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, indent=2, ensure_ascii=False)
            
            self.load_history()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar la consulta:\n{str(e)}")

    def load_from_history(self, item):
        data = item.data(Qt.UserRole)
        self.sql_text.setPlainText(data["query"])
        self.query_name_edit.setText(data.get("name", ""))

    def load_history(self):
        self.history_list.clear()
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
                
                for entry in reversed(history):
                    item_text = f"{entry.get('name', 'Sin nombre')}"
                    item = QListWidgetItem(item_text)
                    item.setData(Qt.UserRole, entry)
                    self.history_list.addItem(item)
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo cargar el historial:\n{str(e)}")

    def show_history_context_menu(self, position):
        item = self.history_list.itemAt(position)
        if not item:
            return
            
        menu = QMenu()
        rename_action = menu.addAction("Renombrar")
        delete_action = menu.addAction("Eliminar")
        
        action = menu.exec_(self.history_list.mapToGlobal(position))
        
        if action == rename_action:
            self.rename_history_item(item)
        elif action == delete_action:
            self.delete_history_item(item)

    def rename_history_item(self, item):
        data = item.data(Qt.UserRole)
        new_name, ok = QInputDialog.getText(self, "Renombrar consulta",
                                          "Nuevo nombre:",
                                          QLineEdit.Normal,
                                          data.get("name", ""))
        if ok and new_name:
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
                
                for entry in history:
                    if entry["query"] == data["query"] and entry.get("timestamp") == data.get("timestamp"):
                        entry["name"] = new_name
                        break
                
                with open(self.history_file, 'w', encoding='utf-8') as f:
                    json.dump(history, f, indent=2, ensure_ascii=False)
                
                self.load_history()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo renombrar:\n{str(e)}")

    def delete_history_item(self, item):
        data = item.data(Qt.UserRole)
        reply = QMessageBox.question(self, "Eliminar consulta",
                                   "¿Estás seguro de que quieres eliminar esta consulta?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
                
                history = [entry for entry in history if not (
                    entry["query"] == data["query"] and 
                    entry.get("timestamp") == data.get("timestamp")
                )]
                
                with open(self.history_file, 'w', encoding='utf-8') as f:
                    json.dump(history, f, indent=2, ensure_ascii=False)
                
                self.load_history()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo eliminar:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelSQLAnalyzer()
    window.show()
    sys.exit(app.exec_())