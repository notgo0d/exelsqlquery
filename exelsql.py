#!/usr/bin/env python3
import sys
import sqlite3
import pandas as pd
import json
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, 
                            QLabel, QPushButton, QFileDialog, QTextEdit, 
                            QTableView, QComboBox, QHBoxLayout, QMessageBox,
                            QInputDialog, QLineEdit, QDialog, QFormLayout,
                            QDialogButtonBox, QListWidget)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QStandardItemModel, QStandardItem

class QueryManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Saved Queries")
        self.setModal(True)
        self.resize(500, 400)
        self.layout = QVBoxLayout()
        
        # Query list
        self.query_list = QListWidget()
        self.query_list.itemDoubleClicked.connect(self.edit_query)
        self.layout.addWidget(self.query_list)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        self.btn_delete = QPushButton("Delete")
        self.btn_delete.clicked.connect(self.delete_query)
        button_layout.addWidget(self.btn_delete)
        
        self.btn_edit = QPushButton("Edit")
        self.btn_edit.clicked.connect(self.edit_query)
        button_layout.addWidget(self.btn_edit)
        
        self.btn_close = QPushButton("Close")
        self.btn_close.clicked.connect(self.accept)
        button_layout.addWidget(self.btn_close)
        
        self.layout.addLayout(button_layout)
        self.setLayout(self.layout)
        
    def load_queries(self, templates):
        """Load all custom queries into the list"""
        self.query_list.clear()
        if "Custom" in templates:
            for name, query in templates["Custom"]:
                self.query_list.addItem(name)
                
    def delete_query(self):
        """Delete selected query"""
        current_item = self.query_list.currentItem()
        if current_item:
            reply = QMessageBox.question(
                self, 
                'Delete Query', 
                f"Delete '{current_item.text()}'?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.parent().delete_custom_query(current_item.text())
                self.query_list.takeItem(self.query_list.row(current_item))
                
    def edit_query(self):
        """Edit selected query"""
        current_item = self.query_list.currentItem()
        if current_item:
            self.parent().edit_custom_query(current_item.text())
            self.accept()

class ExcelSQLApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ExcelSQLAnalyzer", "AppConfig")
        self.init_ui()
        self.load_query_templates()
        self.db_conn = None
        self.current_results = None

    def init_ui(self):
        self.setWindowTitle("Excel SQL Analyzer")
        self.setGeometry(100, 100, 1000, 800)
        
        # Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout()
        
        # File Section
        file_layout = QHBoxLayout()
        self.btn_open = QPushButton("📂 Open Excel File")
        self.btn_open.clicked.connect(self.open_file)
        file_layout.addWidget(self.btn_open)
        
        self.lbl_file = QLabel("No file loaded")
        file_layout.addWidget(self.lbl_file)
        self.layout.addLayout(file_layout)
        
        # Table Selection
        self.table_combo = QComboBox()
        self.table_combo.currentTextChanged.connect(self.table_changed)
        self.layout.addWidget(QLabel("Select Sheet:"))
        self.layout.addWidget(self.table_combo)
        
        # Query Templates
        query_template_layout = QHBoxLayout()
        self.layout.addWidget(QLabel("Query Templates:"))
        self.query_combo = QComboBox()
        self.query_combo.currentTextChanged.connect(self.apply_template)
        query_template_layout.addWidget(self.query_combo)
        
        self.btn_manage = QPushButton("Manage Queries")
        self.btn_manage.clicked.connect(self.manage_queries)
        query_template_layout.addWidget(self.btn_manage)
        self.layout.addLayout(query_template_layout)
        
        # Format-Specific Queries
        self.format_query_combo = QComboBox()
        self.format_query_combo.currentTextChanged.connect(self.apply_format_query)
        self.layout.addWidget(QLabel("Format-Specific Queries:"))
        self.layout.addWidget(self.format_query_combo)
        
        # SQL Editor
        self.sql_editor = QTextEdit()
        self.sql_editor.setPlaceholderText("Write your SQL query here...\nExample: SELECT * FROM Sheet1 LIMIT 10")
        self.sql_editor.setStyleSheet("font-family: monospace;")
        self.layout.addWidget(self.sql_editor)
        
        # Button Row
        button_layout = QHBoxLayout()
        self.btn_execute = QPushButton("⚡ Execute Query")
        self.btn_execute.clicked.connect(self.execute_query)
        button_layout.addWidget(self.btn_execute)
        
        self.btn_save = QPushButton("💾 Save Query")
        self.btn_save.clicked.connect(self.save_query)
        button_layout.addWidget(self.btn_save)
        
        self.btn_export = QPushButton("📤 Export Results")
        self.btn_export.clicked.connect(self.export_results)
        button_layout.addWidget(self.btn_export)
        self.layout.addLayout(button_layout)
        
        # Results Table
        self.results_table = QTableView()
        self.layout.addWidget(self.results_table)
        
        central_widget.setLayout(self.layout)
        
        # Load window state
        if self.settings.value("windowGeometry"):
            self.restoreGeometry(self.settings.value("windowGeometry"))
        if self.settings.value("windowState"):
            self.restoreState(self.settings.value("windowState"))

    def load_query_templates(self):
        """Load query templates from JSON file or defaults"""
        default_templates = {
            "Basic": [
                ("Select All", "SELECT * FROM ${table} LIMIT 100"),
                ("Column Summary", "SELECT ${columns} FROM ${table}"),
                ("Count Rows", "SELECT COUNT(*) as total_rows FROM ${table}")
            ],
            "Analysis": [
                ("Find Duplicates", "SELECT ${columns}, COUNT(*) as duplicates FROM ${table} GROUP BY ${columns} HAVING duplicates > 1"),
                ("Top N Records", "SELECT * FROM ${table} ORDER BY ${column} DESC LIMIT ${n}")
            ],
            "Custom": []
        }
        
        self.templates = default_templates
        
        # Try loading custom templates
        if Path("query_templates.json").exists():
            try:
                with open("query_templates.json") as f:
                    custom_templates = json.load(f)
                    # Ensure we have all default categories
                    for category in default_templates:
                        if category not in custom_templates:
                            custom_templates[category] = default_templates[category]
                    self.templates = custom_templates
            except Exception as e:
                QMessageBox.warning(self, "Warning", f"Couldn't load custom templates:\n{str(e)}")

    def save_templates_to_file(self):
        """Save all templates to JSON file"""
        try:
            with open("query_templates.json", "w") as f:
                json.dump(self.templates, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Couldn't save templates:\n{str(e)}")

    def manage_queries(self):
        """Open query management dialog"""
        dialog = QueryManagerDialog(self)
        dialog.load_queries(self.templates)
        dialog.exec_()

    def delete_custom_query(self, query_name):
        """Delete a custom query by name"""
        for i, (name, _) in enumerate(self.templates["Custom"]):
            if name == query_name:
                self.templates["Custom"].pop(i)
                self.save_templates_to_file()
                self.update_query_templates()
                self.statusBar().showMessage(f"Deleted query: {query_name}", 3000)
                break

    def edit_custom_query(self, query_name):
        """Edit an existing custom query"""
        for i, (name, query) in enumerate(self.templates["Custom"]):
            if name == query_name:
                new_query, ok = QInputDialog.getMultiLineText(
                    self, 
                    "Edit Query", 
                    f"Edit query '{query_name}':",
                    query
                )
                
                if ok and new_query.strip():
                    self.templates["Custom"][i] = (name, new_query.strip())
                    self.save_templates_to_file()
                    self.update_query_templates()
                    self.sql_editor.setPlainText(new_query.strip())
                    self.statusBar().showMessage(f"Updated query: {query_name}", 3000)
                return

    def open_file(self):
        """Open and load Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Open Excel File", 
            self.settings.value("last_dir", ""),
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            try:
                self.settings.setValue("last_dir", str(Path(file_path).parent))
                self.lbl_file.setText(Path(file_path).name)
                
                xls = pd.ExcelFile(file_path)
                self.db_conn = sqlite3.connect(':memory:')
                
                self.table_combo.clear()
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name)
                    # Clean column names for SQL
                    df.columns = [str(col).replace(' ', '_').replace('-', '_')
                                 .replace('(', '').replace(')', '')
                                 .replace('/', '_').replace('\\', '_')
                                 for col in df.columns]
                    df.to_sql(sheet_name, self.db_conn, index=False)
                    self.table_combo.addItem(sheet_name)
                
                self.update_query_templates()
                self.statusBar().showMessage(f"Loaded: {Path(file_path).name}", 3000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")

    def table_changed(self):
        """When user selects a different table/sheet"""
        current_table = self.table_combo.currentText()
        if current_table and self.db_conn:
            self.update_format_specific_queries(current_table)
            self.statusBar().showMessage(f"Selected table: {current_table}", 2000)

    def update_query_templates(self):
        """Refresh the template dropdowns"""
        self.query_combo.clear()
        for category in self.templates:
            for query_name, _ in self.templates[category]:
                self.query_combo.addItem(f"{category}: {query_name}")

    def update_format_specific_queries(self, table_name):
        """Update queries based on detected data format"""
        if not self.db_conn:
            return
            
        # Get sample data to detect format
        df = pd.read_sql(f"SELECT * FROM {table_name} LIMIT 1", self.db_conn)
        format_type = self.detect_format(df)
        
        self.format_query_combo.clear()
        queries = self.get_format_queries(format_type, table_name)
        for name, query in queries:
            self.format_query_combo.addItem(name, query)

    def detect_format(self, df):
        """Auto-detect data format from columns"""
        cols = [str(c).lower() for c in df.columns]
        
        if any(x in cols for x in ['date', 'time', 'year', 'month']):
            return "timeseries"
        elif any(x in cols for x in ['price', 'amount', 'total', 'cost', 'revenue']):
            return "financial"
        elif any(x in cols for x in ['product', 'sku', 'inventory', 'stock', 'quantity']):
            return "inventory"
        elif any(x in cols for x in ['name', 'address', 'email', 'phone']):
            return "contact"
        return "generic"

    def get_format_queries(self, format_type, table_name):
        """Get queries for specific formats"""
        # First get column info
        columns = pd.read_sql(f"PRAGMA table_info({table_name})", self.db_conn)['name']
        numeric_cols = []
        date_cols = []
        
        # Try to detect column types
        sample = pd.read_sql(f"SELECT * FROM {table_name} LIMIT 1", self.db_conn)
        for col in columns:
            if pd.api.types.is_numeric_dtype(sample[col]):
                numeric_cols.append(col)
            elif pd.api.types.is_datetime64_any_dtype(sample[col]):
                date_cols.append(col)
        
        queries = {
            "timeseries": [
                ("Daily Summary", f"SELECT {date_cols[0] if date_cols else columns[0]}, " + 
                 f"SUM({numeric_cols[0] if numeric_cols else columns[1]}) FROM {table_name} GROUP BY " +
                 f"{date_cols[0] if date_cols else columns[0]}"),
                ("Monthly Trend", f"SELECT strftime('%Y-%m', {date_cols[0] if date_cols else columns[0]}) as month, " +
                 f"COUNT(*) FROM {table_name} GROUP BY month")
            ],
            "financial": [
                ("Transaction Summary", f"SELECT {columns[0]}, SUM({numeric_cols[0] if numeric_cols else columns[1]}) " +
                 f"FROM {table_name} GROUP BY {columns[0]}"),
                ("Large Transactions", f"SELECT * FROM {table_name} WHERE " +
                 f"{numeric_cols[0] if numeric_cols else columns[1]} > (SELECT AVG({numeric_cols[0] if numeric_cols else columns[1]}) FROM {table_name})")
            ],
            "inventory": [
                ("Low Stock", f"SELECT {columns[0]}, {numeric_cols[0] if numeric_cols else columns[1]} " +
                 f"FROM {table_name} WHERE {numeric_cols[0] if numeric_cols else columns[1]} < " +
                 f"(SELECT AVG({numeric_cols[0] if numeric_cols else columns[1]}) FROM {table_name})"),
                ("Popular Items", f"SELECT {columns[0]}, SUM({numeric_cols[0] if numeric_cols else columns[1]}) as total " +
                 f"FROM {table_name} GROUP BY {columns[0]} ORDER BY total DESC")
            ],
            "contact": [
                ("Contact Count by Category", f"SELECT {columns[1]}, COUNT(*) FROM {table_name} GROUP BY {columns[1]}"),
                ("Missing Information", f"SELECT * FROM {table_name} WHERE " + 
                 " OR ".join(f"{col} IS NULL" for col in columns[:3]))
            ],
            "generic": [
                ("Basic Analysis", f"SELECT * FROM {table_name} LIMIT 100"),
                ("Column Stats", "SELECT " + ", ".join(
                    f"COUNT({col}) as {col}_count, " + 
                    (f"AVG({col}) as {col}_avg" if col in numeric_cols else f"MIN({col}) as {col}_min")
                    for col in columns) + f" FROM {table_name}")
            ]
        }
        return queries.get(format_type, queries["generic"])

    def apply_template(self):
        """Apply selected query template"""
        if not self.db_conn or self.table_combo.currentText() == "":
            return
            
        current_table = self.table_combo.currentText()
        template_text = self.query_combo.currentText()
        
        # Find the matching template
        for category in self.templates:
            for name, template in self.templates[category]:
                if f"{category}: {name}" == template_text:
                    # Replace placeholders
                    query = template.replace("${table}", current_table)
                    
                    # Get column list for ${columns} placeholder
                    cols = pd.read_sql(f"PRAGMA table_info({current_table})", self.db_conn)['name']
                    query = query.replace("${columns}", ", ".join(cols))
                    
                    self.sql_editor.setPlainText(query)
                    return

    def apply_format_query(self):
        """Apply selected format-specific query"""
        if self.format_query_combo.currentIndex() >= 0:
            query = self.format_query_combo.currentData()
            if query:
                self.sql_editor.setPlainText(query)

    def execute_query(self):
        """Run the SQL query and display results"""
        if not self.db_conn:
            QMessageBox.warning(self, "Warning", "No database connection! Load a file first.")
            return
            
        query = self.sql_editor.toPlainText()
        if not query.strip():
            QMessageBox.warning(self, "Warning", "Please enter a SQL query")
            return
            
        try:
            df = pd.read_sql(query, self.db_conn)
            
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(df.columns)
            
            for _, row in df.iterrows():
                items = [QStandardItem(str(x)) for x in row]
                model.appendRow(items)
                
            self.results_table.setModel(model)
            self.results_table.resizeColumnsToContents()
            self.current_results = df
            self.statusBar().showMessage(f"Query executed. Returned {len(df)} rows", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Query Error", f"Failed to execute query:\n{str(e)}")

    def save_query(self):
        """Save current query to templates"""
        query_text = self.sql_editor.toPlainText().strip()
        if not query_text:
            QMessageBox.warning(self, "Warning", "No query to save!")
            return
            
        # Check if this is already a saved query
        for name, saved_query in self.templates["Custom"]:
            if saved_query == query_text:
                QMessageBox.information(self, "Info", f"This query is already saved as '{name}'")
                return
        
        query_name, ok = QInputDialog.getText(
            self, 
            "Save Query", 
            "Enter a name for this query:",
            QLineEdit.Normal,
            ""
        )
        
        if ok and query_name:
            # Check if name already exists
            for name, _ in self.templates["Custom"]:
                if name == query_name:
                    QMessageBox.warning(self, "Warning", f"A query named '{query_name}' already exists!")
                    return
            
            # Add to Custom category
            self.templates["Custom"].append((query_name, query_text))
            
            # Save to JSON file
            try:
                self.save_templates_to_file()
                self.update_query_templates()
                QMessageBox.information(self, "Saved", f"Query '{query_name}' saved successfully!")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Couldn't save query:\n{str(e)}")

    def export_results(self):
        """Export query results to file"""
        if not hasattr(self, 'current_results') or self.current_results is None:
            QMessageBox.warning(self, "Warning", "No results to export! Run a query first.")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Results",
            self.settings.value("last_export_dir", ""),
            "CSV Files (*.csv);;Excel Files (*.xlsx);;JSON Files (*.json)"
        )
        
        if file_path:
            try:
                self.settings.setValue("last_export_dir", str(Path(file_path).parent))
                
                if file_path.endswith('.csv'):
                    self.current_results.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    self.current_results.to_excel(file_path, index=False)
                elif file_path.endswith('.json'):
                    self.current_results.to_json(file_path, orient='records', indent=2)
                else:
                    self.current_results.to_csv(file_path, index=False)
                
                QMessageBox.information(self, "Success", f"Results exported to:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

    def closeEvent(self, event):
        """Save window state on close"""
        self.settings.setValue("windowGeometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        if self.db_conn:
            self.db_conn.close()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern style
    
    # Set application info (for QSettings)
    QApplication.setApplicationName("ExcelSQLAnalyzer")
    QApplication.setOrganizationName("YourOrg")
    
    window = ExcelSQLApp()
    window.show()
    sys.exit(app.exec_())