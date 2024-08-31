import os
import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog

class ExcelProcessor(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.path = None

    def initUI(self):
        layout = QtWidgets.QVBoxLayout()

        self.file_type_combo = QtWidgets.QComboBox()
        self.file_type_combo.addItems(["Estudiantes Activos", "Estudiantes Moodle"])
        layout.addWidget(self.file_type_combo)

        button_layout = QtWidgets.QHBoxLayout()
        self.select_file_btn = QtWidgets.QPushButton("Seleccionar archivo")
        self.select_file_btn.clicked.connect(self.open_file_dialog)
        button_layout.addWidget(self.select_file_btn)

        self.select_folder_btn = QtWidgets.QPushButton("Seleccionar carpeta")
        self.select_folder_btn.clicked.connect(self.open_folder_dialog)
        button_layout.addWidget(self.select_folder_btn)

        layout.addLayout(button_layout)

        self.process_btn = QtWidgets.QPushButton("Procesar")
        self.process_btn.clicked.connect(self.process_files)
        layout.addWidget(self.process_btn)

        self.log_text = QtWidgets.QTextEdit()
        layout.addWidget(self.log_text)

        self.setLayout(layout)
        self.setWindowTitle('Procesador de Archivos Excel')
        self.show()

    def open_file_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos Excel (*.xlsx)", options=options)
        if file_name:
            self.path = file_name
            self.log_text.append(f"Archivo seleccionado: {self.path}")

    def open_folder_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        folder_name = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta", options=options)
        if folder_name:
            self.path = folder_name
            self.log_text.append(f"Carpeta seleccionada: {self.path}")

    def process_files(self):
        if not self.path:
            self.show_error("Por favor, selecciona un archivo o carpeta primero.")
            return

        file_type = self.file_type_combo.currentText()
        try:
            data, log = self.combine_excel_sheets(self.path, file_type)
            self.save_data(data)
            self.show_log(log)
        except Exception as e:
            self.show_error(f"Error al procesar: {e}")

    def combine_excel_sheets(self, path, file_type):
        all_data = []
        log = []

        if os.path.isfile(path):
            files = [path]
        elif os.path.isdir(path):
            files = [os.path.join(path, filename) for filename in os.listdir(path) if filename.endswith('.xlsx')]
        else:
            raise ValueError("La ruta proporcionada no es válida")

        for file_path in files:
            try:
                log.append(f"Procesando archivo: {os.path.basename(file_path)}")
                xls = pd.ExcelFile(file_path)

                for sheet_name in xls.sheet_names:
                    if file_type == "Estudiantes Activos":
                        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=6, usecols="B:H")
                        df = self.normalize_columns(df)
                    else:  # Estudiantes Moodle
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        df = self.process_estudiantes_moodle(df)

                    df = df.dropna(how="all")

                    if not df.empty:
                        df["SheetName"] = sheet_name
                        df["FileName"] = os.path.basename(file_path)
                        all_data.append(df)

                log.append(f"Archivo procesado con éxito: {os.path.basename(file_path)}")
            except Exception as e:
                log.append(f"Error al procesar el archivo {os.path.basename(file_path)}: {e}")

        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            combined_data = self.finalize_combined_data(combined_data, file_type)
            return combined_data, log
        else:
            log.append("No se encontraron datos para combinar.")
            return pd.DataFrame(), log

        
    def normalize_columns(self, df):
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        column_mapping = {
            "APELLIDO 1": "apellido1",
            "APELLIDO 2": "apellido2",
            "NOMBRE 1": "nombre1",
            "NOMBRE 2": "nombre2",
            "# CELULAR": "TELEFONO",
            "CELULAR": "TELEFONO",
            "CORREO ELECTRONICO": "CORREO",
            "CORREO ELECTRÓNICO": "CORREO",
        }

        for old_col, new_col in column_mapping.items():
            if old_col in df.columns:
                df[new_col] = df.get(new_col, df[old_col])
                if old_col != new_col:
                    df = df.drop(columns=[old_col])

        df = df.loc[:, ~df.columns.duplicated()]
        df = df.dropna(how="all")

        return df

    def process_estudiantes_moodle(self, df):
        if 'firstname' in df.columns:
            df[['nombre1', 'nombre2']] = df['firstname'].str.extract(r'(\S+)\s*(.*)', expand=True)
        if 'lastname' in df.columns:
            df[['apellido1', 'apellido2']] = df['lastname'].str.extract(r'(\S+)\s*(.*)', expand=True)
        
        df = df.rename(columns={
            'idnumber': 'CEDULA',
            'profile_field_Proaca': 'estado_u',
            'email': 'CORREO'
        })

        return df[['CEDULA', 'apellido1', 'apellido2', 'nombre1', 'nombre2', 'CORREO', 'estado_u']]

    def finalize_combined_data(self, combined_data, file_type):
        if 'CEDULA' in combined_data.columns:
            combined_data['CEDULA'] = pd.to_numeric(combined_data['CEDULA'], errors='coerce')
            combined_data.dropna(subset=['CEDULA'], inplace=True)
            combined_data['CEDULA'] = combined_data['CEDULA'].astype(int).astype(str)

        if 'TELEFONO' in combined_data.columns:
            combined_data['TELEFONO'] = pd.to_numeric(combined_data['TELEFONO'], errors='coerce')
            combined_data.dropna(subset=['TELEFONO'], inplace=True)
            combined_data['TELEFONO'] = combined_data['TELEFONO'].astype(int).astype(str)

        if file_type == "Estudiantes Activos":
            combined_data['jornada'] = combined_data['FileName'].apply(
                lambda x: 'FS' if 'FS' in x else ('DIU' if 'DIU' in x else ('NOC' if 'NOC' in x or 'ESPECIALIZA' in x else '')))
            combined_data['estado_u'] = combined_data['FileName'].apply(
                lambda x: 'Diplomado' if 'DIPLOMADO' in x else ('Tecnico' if 'TECNICO' in x else ('Profesional' if 'PROF' in x or 'DERECHO' in x else ('Especialización' if 'ESPECIALIZA' in x else ''))))

        column_order = [
            "CEDULA",
            "apellido1",
            "apellido2",
            "nombre1",
            "nombre2",
            "TELEFONO",
            "CORREO",
            "estado_u",
            "jornada",
            "SheetName",
            "FileName",
        ]
        
        result = combined_data[[col for col in column_order if col in combined_data.columns]]
        
        result.columns = result.columns.str.lower()
         
        return result

    def save_data(self, data):
        if not data.empty:
            output_file, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Guardar archivo', '', 'Archivos Excel (*.xlsx)')
            if output_file:
                data.to_excel(output_file, index=False)
                self.log_text.append(f"Archivo combinado guardado exitosamente en: {output_file}")
        else:
            self.log_text.append('No se creó el archivo combinado debido a la falta de datos.')

    def show_error(self, message):
        QtWidgets.QMessageBox.critical(self, 'Error', message)

    def show_log(self, log):
        log_message = '\n'.join(log)
        self.log_text.append(log_message)

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    ex = ExcelProcessor()
    app.exec_()