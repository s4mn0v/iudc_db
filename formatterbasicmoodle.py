import pandas as pd
import os
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMainWindow, QApplication, QVBoxLayout, QPushButton, QLabel, QWidget
import re

class ExcelTransformer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transformador de Excel")
        self.resize(500, 400)

        self.path = ''
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.log_text = QLabel("Logs:")
        layout.addWidget(self.log_text)

        self.select_folder_button = QPushButton('Seleccionar Carpeta')
        self.select_folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.select_folder_button)

        self.select_file_button = QPushButton('Seleccionar Archivo')
        self.select_file_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_file_button)

        # Dejar solo la opción "Estudiantes Moodle"
        self.format_combo = QtWidgets.QComboBox()
        self.format_combo.addItems(["Estudiantes Moodle"])
        layout.addWidget(self.format_combo)

        self.process_button = QPushButton('Procesar Archivos')
        self.process_button.clicked.connect(self.combine_excel_sheets)
        layout.addWidget(self.process_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        if folder:
            self.path = folder
            self.log_text.setText(f"Carpeta seleccionada: {folder}")

    def select_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo", "", "Archivos Excel (*.xlsx)")
        if file:
            self.path = file
            self.log_text.setText(f"Archivo seleccionado: {file}")

    def combine_excel_sheets(self):
        if not self.path:
            self.log_text.setText("Por favor, selecciona una carpeta o archivo primero.")
            return

        all_data = []
        log = []

        if os.path.isdir(self.path):
            files = [os.path.join(self.path, f) for f in os.listdir(self.path) if f.endswith(".xlsx")]
        else:
            files = [self.path]

        for file_path in files:
            try:
                log.append(f"Procesando archivo: {file_path}")
                xls = pd.ExcelFile(file_path)

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    # Solo procesar usuarios de Moodle
                    df = self.process_estudiantes_moodle(df)

                    df = df.dropna(how="all")

                    if not df.empty:
                        df["sheetname"] = sheet_name
                        df["filename"] = os.path.basename(file_path)
                        all_data.append(df)

                log.append(f"Archivo procesado con éxito: {file_path}")
            except Exception as e:
                log.append(f"Error al procesar el archivo {file_path}: {e}")

        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            combined_data = self.finalize_combined_data(combined_data)
            output_file = os.path.join(os.path.dirname(self.path), "combined_data.xlsx")
            combined_data.to_excel(output_file, index=False)
            log.append(f"Archivo combinado creado: {output_file}")
        else:
            log.append("No se encontraron datos para combinar.")

        self.log_text.setText("\n".join(log))

    def process_estudiantes_moodle(self, df):
        # Separar firstname y lastname en nombre1, nombre2 y apellido1, apellido2
        if 'firstname' in df.columns:
            df[['nombre1', 'nombre2']] = df['firstname'].str.extract(r'(\S+)\s*(.*)', expand=True)
        if 'lastname' in df.columns:
            df[['apellido1', 'apellido2']] = df['lastname'].str.extract(r'(\S+)\s*(.*)', expand=True)
        
        df = df.rename(columns={
            'idnumber': 'cedula',
            'profile_field_Proaca': 'estado_u',
            'email': 'correo'
        })

        df = df[['cedula', 'apellido1', 'apellido2', 'nombre1', 'nombre2', 'correo', 'estado_u']]
        return df

    def finalize_combined_data(self, combined_data):
        column_order = [
            "cedula",
            "apellido1",
            "apellido2",
            "nombre1",
            "nombre2",
            "telefono",
            "correo",
            "estado_u",
            "jornada",
            "sheetname",
            "filename",
        ]
        combined_data = combined_data[
            [col for col in column_order if col in combined_data.columns]
        ]
        return combined_data

if __name__ == "__main__":
    app = QApplication([])
    window = ExcelTransformer()
    window.show()
    app.exec()
