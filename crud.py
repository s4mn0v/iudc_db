import sys
import psycopg2
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTableWidget, QTableWidgetItem, 
                             QFileDialog, QMessageBox, QTextEdit)
from PyQt6.QtCore import Qt
import pandas as pd
import os

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Student CRUD")
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "CEDULA", "APELLIDO 1", "APELLIDO 2", "NOMBRE 1", "NOMBRE 2",
            "TELEFONO", "CORREO", "estado_u", "jornada", "SheetName", "FileName"
        ])

        layout.addWidget(self.table)

        upload_button = QPushButton("Upload XLSX")
        upload_button.clicked.connect(self.upload_xlsx)
        layout.addWidget(upload_button)

        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.refresh_data)
        layout.addWidget(refresh_button)

        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self.delete_selected)
        layout.addWidget(delete_button)

        self.debug_text = QTextEdit()
        self.debug_text.setReadOnly(True)
        layout.addWidget(self.debug_text)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        self.conn = psycopg2.connect(
            dbname="railway",
            user="postgres",
            password="GQtBwDUEGmPKMbtDyofizQsNEWuCksXs",
            host="junction.proxy.rlwy.net",
            port="36061"
        )
        self.schema = "relacional"
        self.refresh_data()

    def upload_xlsx(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open XLSX File", "", "XLSX Files (*.xlsx)")
        if file_name:
            try:
                df = pd.read_excel(file_name)
                cursor = self.conn.cursor()

                cursor.execute(f"SET search_path TO {self.schema}")

                sheet_name = pd.ExcelFile(file_name).sheet_names[0]
                file_basename = os.path.basename(file_name)

                self.debug_text.append(f"Uploading file: {file_basename}")
                self.debug_text.append(f"Sheet name: {sheet_name}")

                for _, row in df.iterrows():
                    query = f"""
                    INSERT INTO {self.schema}.estudiantes 
                    (CEDULA, APELLIDO1, APELLIDO2, NOMBRE1, NOMBRE2,
                    TELEFONO, CORREO, estado_u, jornada, SheetName, FileName)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (CEDULA) DO UPDATE SET
                        APELLIDO1 = EXCLUDED.APELLIDO1,
                        APELLIDO2 = EXCLUDED.APELLIDO2,
                        NOMBRE1 = EXCLUDED.NOMBRE1,
                        NOMBRE2 = EXCLUDED.NOMBRE2,
                        TELEFONO = EXCLUDED.TELEFONO,
                        CORREO = EXCLUDED.CORREO,
                        estado_u = EXCLUDED.estado_u,
                        jornada = EXCLUDED.jornada,
                        FileName = EXCLUDED.FileName,
                        SheetName = COALESCE({self.schema}.estudiantes.SheetName, EXCLUDED.SheetName)
                    """
                    params = (
                        row['CEDULA'], row['APELLIDO 1'], row['APELLIDO 2'],
                        row['NOMBRE 1'], row['NOMBRE 2'], row['TELEFONO'],
                        row['CORREO'], row['estado_u'], row['jornada'],
                        sheet_name, file_basename
                    )
                    
                    self.debug_text.append(f"Executing query for CEDULA: {row['CEDULA']}")
                    cursor.execute(query, params)

                self.conn.commit()
                QMessageBox.information(self, "Success", "Data uploaded successfully!")
                self.refresh_data()
            except Exception as e:
                self.debug_text.append(f"Error: {str(e)}")
                QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def refresh_data(self):
        cursor = self.conn.cursor()
        cursor.execute(f"SET search_path TO {self.schema}")
        cursor.execute(f"SELECT * FROM {self.schema}.estudiantes")
        data = cursor.fetchall()

        self.table.setRowCount(len(data))
        for row, record in enumerate(data):
            for col, value in enumerate(record):
                self.table.setItem(row, col, QTableWidgetItem(str(value)))

    def delete_selected(self):
        selected_rows = set(index.row() for index in self.table.selectedIndexes())
        if not selected_rows:
            return

        reply = QMessageBox.question(self, "Confirm Deletion", 
                                     "Are you sure you want to delete the selected records?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            cursor = self.conn.cursor()
            cursor.execute(f"SET search_path TO {self.schema}")
            for row in sorted(selected_rows, reverse=True):
                cedula = self.table.item(row, 0).text()
                cursor.execute(f"DELETE FROM {self.schema}.estudiantes WHERE CEDULA = %s", (cedula,))
                self.table.removeRow(row)
            self.conn.commit()
            QMessageBox.information(self, "Success", "Selected records deleted successfully!")

    def closeEvent(self, event):
        self.conn.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())