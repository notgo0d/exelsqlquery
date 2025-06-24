import sys, os, json
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTextEdit, QPushButton, QVBoxLayout,
    QWidget, QComboBox, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout,
    QMessageBox, QTreeWidget, QTreeWidgetItem, QSplitter, QStyleFactory,
    QLineEdit, QInputDialog, QMenu, QToolButton, QAction, QCompleter,
    QTabWidget, QHeaderView
)
from PyQt5.QtGui import QFont, QColor, QSyntaxHighlighter, QTextCharFormat, QKeySequence
from PyQt5.QtCore import Qt, QRegExp

class SQLHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor("blue"))
        fmt.setFontWeight(QFont.Bold)
        self.keyword_format = fmt
        self.keywords = [
            "SELECT","FROM","WHERE","AND","OR","INSERT","INTO","VALUES",
            "UPDATE","SET","DELETE","CREATE","TABLE","DROP","ALTER",
            "ADD","JOIN","ON","AS","DISTINCT","GROUP","BY","ORDER","LIMIT"
        ]

    def highlightBlock(self, text):
        for kw in self.keywords:
            expr = QRegExp(f"\\b{kw}\\b", Qt.CaseInsensitive)
            idx = expr.indexIn(text)
            while idx >= 0:
                self.setFormat(idx, expr.matchedLength(), self.keyword_format)
                idx = expr.indexIn(text, idx + expr.matchedLength())

class AutoCompleteTextEdit(QTextEdit):
    def __init__(self, parent=None, word_list=None):
        super().__init__(parent)
        self.completer = QCompleter(word_list or [], self)
        self.completer.setWidget(self)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.activated.connect(self.insert_completion)

    def insert_completion(self, completion):
        tc = self.textCursor()
        pref = self.completer.completionPrefix()
        tc.movePosition(tc.Left, tc.KeepAnchor, len(pref))
        tc.insertText(completion)
        self.setTextCursor(tc)

    def textUnderCursor(self):
        tc = self.textCursor()
        tc.select(tc.WordUnderCursor)
        return tc.selectedText()

    def keyPressEvent(self, e):
        super().keyPressEvent(e)
        pref = self.textUnderCursor()
        if not pref:
            self.completer.popup().hide()
            return
        self.completer.setCompletionPrefix(pref)
        cr = self.cursorRect()
        cr.setWidth(self.completer.popup().sizeHintForColumn(0) +
                    self.completer.popup().verticalScrollBar().sizeHint().width())
        self.completer.complete(cr)

class ExcelSQLAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel SQL Analyzer Pro")
        self.setGeometry(100, 100, 1300, 850)
        self.df_dict = {}
        self.conn = sqlite3.connect(":memory:")
        self.history_file = "query_history.json"
        self.current_df = None
        self.dark_mode = True

        self.init_ui()
        self.load_history()
        self.toggle_theme()

    def init_ui(self):
        main_layout = QVBoxLayout()
        file_layout = QHBoxLayout()
        btn_load = QPushButton("Cargar Excel(s)")
        btn_load.clicked.connect(self.load_excel)
        self.sheetCombo = QComboBox()
        self.sheetCombo.currentIndexChanged.connect(self.preview_sheet)
        file_layout.addWidget(btn_load)
        file_layout.addWidget(QLabel("Hoja:"))
        file_layout.addWidget(self.sheetCombo)

        tpl_layout = QHBoxLayout()
        burger = QToolButton()
        burger.setText("☰ Plantillas SQL")
        burger.setPopupMode(QToolButton.InstantPopup)
        menu = QMenu()
        snippets = {
            "SELECT * FROM": "SELECT * FROM ",
            "CREATE TABLE": "CREATE TABLE  (\n);\n",
            "INSERT INTO": "INSERT INTO  () VALUES ();",
            "UPDATE SET": "UPDATE  SET  WHERE ;",
            "DELETE FROM": "DELETE FROM  WHERE ;"
        }
        for label, tmpl in snippets.items():
            act = QAction(label, self)
            act.triggered.connect(lambda _, t=tmpl: self.insert_template(t))
            menu.addAction(act)
        burger.setMenu(menu)
        tpl_layout.addWidget(burger)
        tpl_layout.addStretch()

        self.sqlText = AutoCompleteTextEdit(word_list=self.get_sql_suggestions())
        self.sqlText.setFont(QFont("Courier", 10))
        SQLHighlighter(self.sqlText.document())

        self.nameEdit = QLineEdit()
        self.nameEdit.setPlaceholderText("Nombre de consulta o carpeta/nombre")

        btn_run = QPushButton("Ejecutar SQL")
        btn_run.clicked.connect(self.run_query)
        btn_clear = QPushButton("Limpiar")
        btn_clear.clicked.connect(self.clear_query)
        btn_stats = QPushButton("Resumen Estadístico")
        btn_stats.clicked.connect(self.show_statistics)
        btn_plot = QPushButton("Visualizar Gráfico")
        btn_plot.clicked.connect(self.show_plot)
        btn_export_csv = QPushButton("Exportar CSV")
        btn_export_csv.clicked.connect(lambda: self.export_results("csv"))
        btn_export_excel = QPushButton("Exportar Excel")
        btn_export_excel.clicked.connect(lambda: self.export_results("xlsx"))

        btn_layout = QHBoxLayout()
        for btn in [btn_run, btn_clear, btn_stats, btn_plot, btn_export_csv, btn_export_excel]:
            btn_layout.addWidget(btn)

        self.filterEdit = QLineEdit()
        self.filterEdit.setPlaceholderText("Filtrar resultados...")
        self.filterEdit.textChanged.connect(self.filter_table)

        self.table = QTableWidget()
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.historyTree = QTreeWidget()
        self.historyTree.setHeaderLabel("Consultas / Carpetas")
        self.historyTree.itemDoubleClicked.connect(self.load_from_history)

        main_layout.addLayout(file_layout)
        main_layout.addLayout(tpl_layout)
        main_layout.addWidget(QLabel("Consulta SQL:"))
        main_layout.addWidget(self.sqlText)
        main_layout.addWidget(self.nameEdit)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(self.filterEdit)
        main_layout.addWidget(self.table)
        main_layout.addWidget(QLabel("Organización de consultas:"))
        main_layout.addWidget(self.historyTree)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def insert_template(self, template_text):
        tc = self.sqlText.textCursor()
        tc.insertText(template_text)
        self.sqlText.setTextCursor(tc)

    def toggle_theme(self):
        if self.dark_mode:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            self.setStyleSheet("""
                QWidget { background:#2b2b2b; color:#fff; }
                QLineEdit, QTextEdit, QComboBox, QTreeWidget, QTableWidget { background:#3c3c3c; color:#fff; }
                QPushButton { background:#444; border:1px solid #555; padding:5px; }
                QPushButton:hover { background:#555; }
                QHeaderView::section { background:#444; color:#fff; }
                QMenu { background:#3c3c3c; color:#fff; }
            """)
            self.dark_mode = True
        else:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            self.setStyleSheet("")
            self.dark_mode = False

    def load_excel(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Seleccionar archivo Excel", "", "Excel (*.xlsx *.xls)")
        if not paths:
            return
        for path in paths:
            try:
                dfs = pd.read_excel(path, sheet_name=None)
                for sheet, df in dfs.items():
                    table_name = f"{os.path.splitext(os.path.basename(path))[0]}_{sheet}"
                    df.to_sql(table_name, self.conn, if_exists='replace', index=False)
                    self.df_dict[table_name] = df
                    self.sheetCombo.addItem(table_name)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error cargando Excel:\n{e}")
        self.sqlText.completer.model().setStringList(self.get_sql_suggestions())

    def preview_sheet(self):
        sheet = self.sheetCombo.currentText()
        if sheet:
            df = self.df_dict.get(sheet)
            if df is not None:
                self.current_df = df
                self.show_df(df)

    def run_query(self):
        q = self.sqlText.toPlainText().strip()
        if not q:
            return
        try:
            self.current_df = pd.read_sql_query(q, self.conn)
            self.show_df(self.current_df)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al ejecutar consulta:\n{e}")

    def show_df(self, df):
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str))
        for r, row in enumerate(df.values):
            for c, val in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(val)))

    def filter_table(self, text):
        for row in range(self.table.rowCount()):
            match = False
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and text.lower() in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)

    def clear_query(self):
        self.sqlText.clear()
        self.table.clear()
        self.filterEdit.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

    def show_statistics(self):
        if self.current_df is not None:
            stats = self.current_df.describe(include='all').to_string()
            QMessageBox.information(self, "Resumen Estadístico", stats)

    def show_plot(self):
        if self.current_df is not None:
            numeric_cols = self.current_df.select_dtypes(include='number').columns.tolist()
            if not numeric_cols:
                QMessageBox.warning(self, "Atención", "No hay columnas numéricas para graficar")
                return
            self.current_df[numeric_cols].plot(kind='bar', figsize=(10, 5))
            plt.tight_layout()
            plt.show()

    def export_results(self, fmt):
        if self.current_df is None:
            QMessageBox.warning(self, "Aviso", "Sin resultados para exportar")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exportar",
            filter="CSV (*.csv)" if fmt == "csv" else "Excel (*.xlsx)")
        if not path:
            return
        try:
            if fmt == "csv":
                self.current_df.to_csv(path, index=False)
            else:
                self.current_df.to_excel(path, index=False)
            QMessageBox.information(self, "Éxito", f"Guardado en:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error de exportación", str(e))

    def get_sql_suggestions(self):
        keywords = [
            "SELECT", "FROM", "WHERE", "AND", "OR", "INSERT", "INTO", "VALUES",
            "UPDATE", "SET", "DELETE", "CREATE", "TABLE", "DROP", "ALTER",
            "ADD", "JOIN", "ON", "AS", "DISTINCT", "GROUP", "BY", "ORDER", "LIMIT"
        ]
        return keywords + list(self.df_dict.keys())

    def load_history(self):
        self.historyTree.clear()
        if not os.path.exists(self.history_file):
            return
        with open(self.history_file, encoding='utf-8') as f:
            hist = json.load(f)
        folders = {}
        for e in hist:
            fld = e.get("folder", "")
            if fld not in folders:
                folders[fld] = QTreeWidgetItem(self.historyTree, [fld])
            itm = QTreeWidgetItem(folders[fld], [e.get("label")])
            itm.setData(0, Qt.UserRole, e)
        self.historyTree.expandAll()

    def load_from_history(self, item):
        data = item.data(0, Qt.UserRole)
        if data:
            self.sqlText.setPlainText(data["query"])
            self.nameEdit.setText(data["full_name"])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelSQLAnalyzer()
    window.show()
    sys.exit(app.exec_())