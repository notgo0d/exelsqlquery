import sys, os, json
import pandas as pd
import sqlite3
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTextEdit, QPushButton, QVBoxLayout,
    QWidget, QComboBox, QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout,
    QMessageBox, QTreeWidget, QTreeWidgetItem, QSplitter, QStyleFactory,
    QLineEdit, QInputDialog, QMenu, QToolButton, QAction, QCompleter
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
        self.keywords = ["SELECT","FROM","WHERE","AND","OR","INSERT","INTO","VALUES",
                         "UPDATE","SET","DELETE","CREATE","TABLE","DROP","ALTER",
                         "ADD","JOIN","ON","AS","DISTINCT","GROUP","BY","ORDER","LIMIT"]
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
        self.setWindowTitle("Excel SQL Analyzer")
        self.setGeometry(100,100,1200,800)
        self.df_dict, self.conn = {}, sqlite3.connect(":memory:")
        self.history_file = "query_history.json"
        self.current_df = None
        self.init_ui()
        self.load_history()

    def init_ui(self):
        main = QVBoxLayout()
        fileLayout = QHBoxLayout()
        btnLoad = QPushButton("Cargar Excel")
        btnLoad.clicked.connect(self.load_excel)
        self.sheetCombo = QComboBox()
        self.sheetCombo.currentIndexChanged.connect(self.preview_sheet)
        fileLayout.addWidget(btnLoad); fileLayout.addWidget(QLabel("Hoja:")); fileLayout.addWidget(self.sheetCombo)

        tplLayout = QHBoxLayout()
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
            act.setShortcut(label[0:4])  # shortcut: first 4 letters
            self.addAction(act)
            menu.addAction(act)
        burger.setMenu(menu)
        tplLayout.addWidget(burger); tplLayout.addStretch()

        self.sqlText = AutoCompleteTextEdit(word_list=self.get_sql_suggestions())
        self.sqlText.setFont(QFont("Courier", 10))
        SQLHighlighter(self.sqlText.document())

        self.nameEdit = QLineEdit()
        self.nameEdit.setPlaceholderText("Nombre de consulta o carpeta/nombre")
        btnRun = QPushButton("Ejecutar SQL")
        btnRun.clicked.connect(self.run_query)
        btnExportCSV = QPushButton("Exportar CSV")
        btnExportCSV.clicked.connect(lambda: self.export_results("csv"))
        btnExportExcel = QPushButton("Exportar Excel")
        btnExportExcel.clicked.connect(lambda: self.export_results("xlsx"))
        btnSave = QPushButton("Guardar Consulta")
        btnSave.clicked.connect(self.save_named_query)
        # Keyboard shortcuts
        btnRun.setShortcut(QKeySequence("Ctrl+R"))
        btnExportCSV.setShortcut(QKeySequence("Ctrl+Shift+C"))
        btnExportExcel.setShortcut(QKeySequence("Ctrl+Shift+X"))
        btnSave.setShortcut(QKeySequence("Ctrl+S"))
        btnLayout = QHBoxLayout()
        for b in [btnRun, btnSave, btnExportCSV, btnExportExcel]:
            btnLayout.addWidget(b)

        self.table = QTableWidget()
        self.historyTree = QTreeWidget()
        self.historyTree.setHeaderLabel("Consultas / Carpetas")
        self.historyTree.itemDoubleClicked.connect(self.load_from_history)
        self.historyTree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.historyTree.customContextMenuRequested.connect(self.history_context_menu)

        main.addLayout(fileLayout)
        main.addLayout(tplLayout)
        main.addWidget(QLabel("Consulta SQL:"))
        main.addWidget(self.sqlText)
        main.addWidget(QLabel("Nombre de consulta / carpeta/nombre"))
        main.addWidget(self.nameEdit)
        main.addLayout(btnLayout)
        main.addWidget(self.table)
        main.addWidget(QLabel("Organización de consultas:"))
        main.addWidget(self.historyTree)

        container = QWidget()
        container.setLayout(main)
        self.setCentralWidget(container)
        self.dark_mode = False
        self.toggle_theme()

    def insert_template(self, t):
        tc = self.sqlText.textCursor()
        tc.insertText(t)
        self.sqlText.setTextCursor(tc)

    def toggle_theme(self):
        if self.dark_mode:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            self.setStyleSheet("")
            self.dark_mode = False
        else:
            QApplication.setStyle(QStyleFactory.create("Fusion"))
            self.setStyleSheet("""
                QWidget{background:#2b2b2b;color:#fff}
                QLineEdit,QTextEdit,QComboBox,QTreeWidget{background:#3c3c3c;color:#fff}
                QPushButton{background:#444;border:1px solid #555;padding:5px}
                QPushButton:hover{background:#555}
            """)
            self.dark_mode = True

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            self.df_dict = pd.read_excel(path, sheet_name=None)
            self.sheetCombo.clear()
            self.sheetCombo.addItems(self.df_dict.keys())
            self.preview_sheet()
            self.sqlText.completer.model().setStringList(self.get_sql_suggestions())

    def preview_sheet(self):
        sheet = self.sheetCombo.currentText()
        if sheet:
            df = self.df_dict[sheet]
            df.to_sql(sheet, self.conn, if_exists='replace', index=False)
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
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.astype(str))
        for r, row in enumerate(df.values):
            for c, val in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(val)))

    def export_results(self, fmt):
        if self.current_df is None:
            QMessageBox.warning(self, "Aviso", "Sin resultados para exportar")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exportar", filter="CSV (*.csv)" if fmt=="csv" else "Excel (*.xlsx)")
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

    def save_named_query(self):
        q = self.sqlText.toPlainText().strip()
        if not q:
            QMessageBox.warning(self, "Aviso", "Consulta vacía")
            return
        name, ok = QInputDialog.getText(self, "Guardar", "Carpeta/NombreConsulta", QLineEdit.Normal, self.nameEdit.text())
        if ok and name:
            self.nameEdit.setText(name)
            self.save_history(name, q)

    def save_history(self, name, query):
        hist = []
        if os.path.exists(self.history_file):
            with open(self.history_file, encoding='utf-8') as f:
                hist = json.load(f)
        entry = {"query": query, "full_name": name, "timestamp": pd.Timestamp.now().isoformat()}
        name_parts = name.split("/", 1)
        entry["folder"] = name_parts[0]
        entry["label"] = name_parts[1] if len(name_parts) > 1 else name_parts[0]
        hist.append(entry)
        with open(self.history_file, "w", encoding='utf-8') as f:
            json.dump(hist, f, indent=2, ensure_ascii=False)
        self.load_history()

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

    def history_context_menu(self, pos):
        it = self.historyTree.itemAt(pos)
        if not it:
            return
        data = it.data(0, Qt.UserRole)
        menu = QMenu()
        if data:
            menu.addAction("Eliminar", lambda: self.delete_history(data))
            menu.addAction("Mover a carpeta...", lambda: self.move_history(data))
        else:
            menu.addAction("Crear carpeta", lambda: self.create_folder(it.text(0)))
        menu.exec_(self.historyTree.mapToGlobal(pos))

    def delete_history(self, data):
        with open(self.history_file, encoding='utf-8') as f:
            hist = json.load(f)
        hist = [e for e in hist if e["timestamp"] != data["timestamp"]]
        with open(self.history_file, "w", encoding='utf-8') as f:
            json.dump(hist, f, indent=2, ensure_ascii=False)
        self.load_history()

    def move_history(self, data):
        folder, ok = QInputDialog.getText(self, "Mover a carpeta", "Nueva carpeta:", QLineEdit.Normal, data.get("folder", ""))
        if ok:
            with open(self.history_file, encoding='utf-8') as f:
                hist = json.load(f)
            for e in hist:
                if e["timestamp"] == data["timestamp"]:
                    e["folder"] = folder
                    break
            with open(self.history_file, "w", encoding='utf-8') as f:
                json.dump(hist, f, indent=2, ensure_ascii=False)
            self.load_history()

    def create_folder(self, name):
        folder, ok = QInputDialog.getText(self, "Crear carpeta", "Nombre carpeta:")
        if ok:
            self.save_history(f"{folder}/", "")

    def get_sql_suggestions(self):
        kws = ["SELECT","FROM","WHERE","AND","OR","INSERT","INTO","VALUES",
               "UPDATE","SET","DELETE","CREATE","TABLE","DROP","ALTER","ADD",
               "JOIN","ON","AS","DISTINCT","GROUP","BY","ORDER","LIMIT"]
        tables = list(self.df_dict.keys())
        cols = [c for df in self.df_dict.values() for c in df.columns.astype(str)]
        return sorted(set(kws + tables + cols))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = ExcelSQLAnalyzer()
    win.show()
    sys.exit(app.exec_())
