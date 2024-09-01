import sys
import pandas as pd
import os
from PyQt6 import QtWidgets

class ExcelTransformer(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.folder_path = ""
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout()

        self.select_button = QtWidgets.QPushButton('Seleccionar archivo o carpeta')
        self.select_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.select_button)

        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        self.setLayout(layout)
        self.setWindowTitle('Transformador de Archivos XLSX')
        self.setGeometry(100, 100, 600, 400)

    def open_file_dialog(self):
        file_dialog = QtWidgets.QFileDialog(self, "Seleccionar archivo o carpeta")
        file_dialog.setFileMode(QtWidgets.QFileDialog.FileMode.AnyFile)
        file_dialog.setNameFilter('Archivos Excel (*.xlsx)')
        file_dialog.setViewMode(QtWidgets.QFileDialog.ViewMode.List)

        if file_dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            file_paths = file_dialog.selectedFiles()
            if file_paths:
                path = file_paths[0]
                if os.path.isdir(path):
                    self.folder_path = path
                    self.combine_excel_sheets()
                elif path.endswith('.xlsx'):
                    self.folder_path = os.path.dirname(path)
                    self.process_file(path)
                else:
                    self.show_error("Solo se admiten archivos .xlsx")

    def process_file(self, file_path):
        try:
            data, log = self.combine_excel_sheets(file_path)
            self.save_data(data)
            self.show_log(log)
        except Exception as e:
            self.show_error(f"Error al procesar el archivo {file_path}: {e}")

    def combine_excel_sheets(self, path=None):
        if not path:
            if not hasattr(self, 'folder_path'):
                self.show_error("Por favor, selecciona una carpeta primero.")
                return
            path = self.folder_path
        
        all_data = []
        log = []

        if os.path.isfile(path):
            files = [path]
        elif os.path.isdir(path):
            files = [os.path.join(path, filename) for filename in os.listdir(path) if filename.endswith('.xlsx')]
        else:
            self.show_error("La ruta proporcionada no es válida")
            return

        for file_path in files:
            try:
                log.append(f"Procesando archivo: {os.path.basename(file_path)}")
                xls = pd.ExcelFile(file_path)

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=6, usecols="B:H")
                    df = self.normalize_columns(df)

                    # Eliminar filas completamente vacías
                    df = df.dropna(how="all")

                    # Solo agregar si hay datos después de la limpieza
                    if not df.empty:
                        df["SheetName"] = sheet_name
                        df["FileName"] = os.path.basename(file_path)
                        all_data.append(df)

                log.append(f"Archivo procesado con éxito: {os.path.basename(file_path)}")
            except Exception as e:
                log.append(f"Error al procesar el archivo {os.path.basename(file_path)}: {e}")

        if all_data:
            combined_data = pd.concat(all_data, ignore_index=True)
            combined_data = self.finalize_combined_data(combined_data)
            return combined_data, log
        else:
            log.append("No se encontraron datos para combinar.")
            return pd.DataFrame(), log

    def normalize_columns(self, df):
        # Convertir los nombres de las columnas a string por seguridad
        df.columns = df.columns.astype(str)
        
        # Eliminar columnas vacías o "Unnamed"
        df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]

        # Mapa de columnas esperado al nuevo formato
        column_mapping = {
            "DOCUMENTO": "documento",
            "TIPO DE DOCUMENTO": "tipo de documento",
            "APELLIDO 1": "apellido1",
            "APELLIDO 2": "apellido2",
            "NOMBRE 1": "nombre1",
            "NOMBRE 2": "nombre2",
            "EXPEDICIÓN": "expedicion",
            "ACTA DE GRADO NO.": "acta de grado no.",
            "LIBRO DE GRADO NO.": "libro de grado no.",
            "FOLIO NO.": "folio no.",
            "TÍTULO": "titulo",
            "DÍA DE GRADUACIÓN": "dia de graduacion",
            "SNIES PROGRAMA": "SNIES programa",
            "DIRECCIÓN": "direccion",
            "TELÉFONO": "telefono",
            "CORREO ELECTRÓNICO": "correo electronico"
        }

        # Buscar y normalizar columnas
        for old_col, new_col in column_mapping.items():
            if old_col in df.columns:
                df[new_col] = df.get(new_col, df[old_col])
                if old_col != new_col:
                    df = df.drop(columns=[old_col])

        # Eliminar duplicación de columnas
        df = df.loc[:, ~df.columns.duplicated()]

        # Eliminar filas donde todas las columnas son NaN
        df = df.dropna(how="all")

        return df

    def finalize_combined_data(self, combined_data):
        # Orden de columnas esperado
        column_order = [
            "documento",
            "tipo de documento",
            "apellido1",
            "apellido2",
            "nombre1",
            "nombre2",
            "expedicion",
            "acta de grado no.",
            "libro de grado no.",
            "folio no.",
            "titulo",
            "dia de graduacion",
            "SNIES programa",
            "direccion",
            "telefono",
            "correo electronico",
            "SheetName",
            "FileName"
        ]
        combined_data = combined_data[[col for col in column_order if col in combined_data.columns]]

        return combined_data

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

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ExcelTransformer()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
