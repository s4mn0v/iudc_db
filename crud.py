from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTableWidget, QTableWidgetItem, 
                             QFileDialog, QMessageBox, QTextEdit, QLineEdit, QLabel, QFormLayout)
import psycopg2
import pandas as pd
import os
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Student CRUD")
        self.setGeometry(100, 100, 800, 600)

        self.layout = QVBoxLayout()

        # Formulario de conexión a la base de datos
        self.db_form_layout = QFormLayout()

        self.host_input = QLineEdit()
        self.db_form_layout.addRow("Host:", self.host_input)

        self.port_input = QLineEdit()
        self.db_form_layout.addRow("Port:", self.port_input)

        self.dbname_input = QLineEdit()
        self.db_form_layout.addRow("Database Name:", self.dbname_input)

        self.user_input = QLineEdit()
        self.db_form_layout.addRow("User:", self.user_input)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.db_form_layout.addRow("Password:", self.password_input)

        self.connect_button = QPushButton("Connect to DB")
        self.connect_button.clicked.connect(self.db_connect)
        self.db_form_layout.addWidget(self.connect_button)

        self.layout.addLayout(self.db_form_layout)

        # Tabla y botones del programa principal
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "CEDULA", "APELLIDO 1", "APELLIDO 2", "NOMBRE 1", "NOMBRE 2",
            "TELEFONO", "CORREO", "estado_u", "jornada", "SheetName", "FileName"
        ])

        self.layout.addWidget(self.table)

        self.upload_button = QPushButton("Upload XLSX")
        self.upload_button.clicked.connect(self.upload_xlsx)
        self.layout.addWidget(self.upload_button)

        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.clicked.connect(self.refresh_data)
        self.layout.addWidget(self.refresh_button)

        self.delete_button = QPushButton("Delete Selected")
        self.delete_button.clicked.connect(self.delete_selected)
        self.layout.addWidget(self.delete_button)

        self.debug_text = QTextEdit()
        self.debug_text.setReadOnly(True)
        self.layout.addWidget(self.debug_text)

        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)

        self.conn = None
        self.schema = "relacional"

    def db_connect(self):
        try:
            self.conn = psycopg2.connect(
                dbname=self.dbname_input.text(),
                user=self.user_input.text(),
                password=self.password_input.text(),
                host=self.host_input.text(),
                port=self.port_input.text()
            )
            self.debug_text.append(f"Connected to the database {self.dbname_input.text()} at {self.host_input.text()}:{self.port_input.text()} successfully.")
            
            # Ocultar el formulario después de la conexión exitosa
            for i in reversed(range(self.db_form_layout.count())): 
                widget = self.db_form_layout.itemAt(i).widget()
                if widget:
                    widget.setVisible(False)

            self.refresh_data()
        except Exception as e:
            self.debug_text.append(f"Failed to connect to the database: {str(e)}")
            QMessageBox.critical(self, "Connection Error", f"Failed to connect: {str(e)}")

    def upload_xlsx(self):
        if not self.conn:
            QMessageBox.critical(self, "Error", "Not connected to the database!")
            return

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
                    if 'SheetName' in df.columns:
                        row_sheet_name = row['SheetName']
                    else:
                        row_sheet_name = sheet_name

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
                        row_sheet_name, file_basename
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
        if not self.conn:
            return

        cursor = self.conn.cursor()
        cursor.execute(f"SET search_path TO {self.schema}")
        cursor.execute(f"SELECT * FROM {self.schema}.estudiantes")
        data = cursor.fetchall()

        self.table.setRowCount(len(data))
        for row, record in enumerate(data):
            for col, value in enumerate(record):
                self.table.setItem(row, col, QTableWidgetItem(str(value)))

    def delete_selected(self):
        if not self.conn:
            QMessageBox.critical(self, "Error", "Not connected to the database!")
            return

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
        if self.conn:
            self.conn.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
