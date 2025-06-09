import sys
import os
import pandas as pd
import sqlite3
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTextEdit, QPushButton, QVBoxLayout,
    QWidget, QComboBox, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout,
    QMessageBox, QListWidget, QSplitter, QStyleFactory
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
        self.history_file = "query_history.txt"

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

        self.run_button = QPushButton("Ejecutar SQL")
        self.run_button.clicked.connect(self.run_query)

        self.table_widget = QTableWidget()

        self.history_list = QListWidget()
        self.history_list.itemClicked.connect(self.load_from_history)

        sql_area = QSplitter(Qt.Horizontal)
        sql_area.addWidget(self.sql_text)
        sql_area.addWidget(self.history_list)
        sql_area.setSizes([800, 200])

        layout.addLayout(file_layout)
        layout.addWidget(QLabel("Consulta SQL:"))
        layout.addWidget(sql_area)
        layout.addWidget(self.run_button)
        layout.addWidget(self.table_widget)

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
            self.save_to_history(query)
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

    def save_to_history(self, query):
        if query not in [self.history_list.item(i).text() for i in range(self.history_list.count())]:
            self.history_list.addItem(query)
            with open(self.history_file, 'a', encoding='utf-8') as f:
                f.write(query + '\n')

    def load_from_history(self, item):
        self.sql_text.setPlainText(item.text())

    def load_history(self):
        if os.path.exists(self.history_file):
            with open(self.history_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line:
                        self.history_list.addItem(line)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelSQLAnalyzer()
    window.show()
    sys.exit(app.exec_())
