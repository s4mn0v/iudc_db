import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QLabel,
    QTextEdit,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QHBoxLayout,
    QRadioButton,
    QButtonGroup,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QCursor

class ExcelCombinerApp(QMainWindow):
    def get_stylesheet(self):
        return """
        QMainWindow {
            background-color: #1E1E1E;
        }
        QMainWindow::title {
            background-color: #2D2D2D;
            color: white;
            padding-left: 4px;
        }
        QWidget {
            color: white;
        }
        QPushButton {
            background-color: #3C3C3C;
            border: none;
            color: white;
            padding: 8px 16px;
            text-align: center;
            text-decoration: none;
            font-size: 14px;
            margin: 4px 2px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #4C4C4C;
        }
        QPushButton:pressed {
            background-color: #2D2D2D;
        }
        QLabel {
            color: #CCCCCC;
        }
        QTextEdit, QTableWidget {
            background-color: #2D2D2D;
            color: #CCCCCC;
            border: 1px solid #3C3C3C;
            border-radius: 4px;
        }
        QTabWidget::pane {
            border: 1px solid #3C3C3C;
            background-color: #2D2D2D;
        }
        QTabBar::tab {
            background-color: #3C3C3C;
            color: white;
            padding: 8px 16px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #4C4C4C;
        }
        QScrollBar:vertical {
            border: none;
            background: #2D2D2D;
            width: 10px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            background: #3C3C3C;
            min-height: 20px;
            border-radius: 5px;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """
        
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('./convert.ico'))
        self.setStyleSheet(self.get_stylesheet())
        self.setWindowTitle("Data Manager - Consultorio Tecnologico IUDC")
        self.setFixedSize(800, 600)

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        self.create_transformation_tab()
        self.create_admin_tab()
        self.create_1fn()

        self.folder_path = ""
        self.file_path = ""
        self.combined_data = None

    def create_transformation_tab(self):
        transformation_widget = QWidget()
        layout = QVBoxLayout()

        # Opciones de selección
        self.selection_group = QButtonGroup()
        self.folder_radio = QRadioButton("Seleccionar Carpeta")
        self.file_radio = QRadioButton("Seleccionar Archivo")
        self.selection_group.addButton(self.folder_radio)
        self.selection_group.addButton(self.file_radio)
        
        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.folder_radio)
        radio_layout.addWidget(self.file_radio)
        layout.addLayout(radio_layout)

        # Botón para seleccionar carpeta o archivo
        self.select_button = QPushButton("Seleccionar", self)
        self.select_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.select_button.clicked.connect(self.select_folder_or_file)
        layout.addWidget(self.select_button)

        # Etiqueta para mostrar la selección
        self.selection_label = QLabel("Selección: Ninguna")
        layout.addWidget(self.selection_label)

        # Botón para iniciar la transformación
        self.transform_button = QPushButton("Transformar", self)
        self.transform_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.transform_button.clicked.connect(self.transform_data)
        layout.addWidget(self.transform_button)

        # Texto donde mostrar el log del proceso
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        transformation_widget.setLayout(layout)
        self.tab_widget.addTab(transformation_widget, "Transformación")

    def select_folder_or_file(self):
        if self.folder_radio.isChecked():
            folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
            if folder:
                self.folder_path = folder
                self.file_path = ""
                self.selection_label.setText(f"Carpeta seleccionada: {folder}")
        elif self.file_radio.isChecked():
            file, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo", "", "Excel Files (*.xlsx)")
            if file:
                self.file_path = file
                self.folder_path = ""
                self.selection_label.setText(f"Archivo seleccionado: {file}")
        else:
            self.selection_label.setText("Por favor, seleccione una opción primero.")

    def transform_data(self):
        if not self.folder_path and not self.file_path:
            self.log_text.append("Por favor, selecciona una carpeta o un archivo primero.")
            return

        all_data = []
        log = []

        if self.folder_path:
            # Procesar todos los archivos en la carpeta
            for filename in os.listdir(self.folder_path):
                if filename.endswith(".xlsx"):
                    file_path = os.path.join(self.folder_path, filename)
                    self.process_file(file_path, all_data, log)
        elif self.file_path:
            # Procesar un solo archivo
            self.process_file(self.file_path, all_data, log)

        # Concatenar todos los DataFrames
        if all_data:
            self.combined_data = pd.concat(all_data, ignore_index=True)
            self.combined_data = self.finalize_combined_data(self.combined_data)

            # Guardar el resultado en un archivo Excel
            output_file = os.path.join(os.path.dirname(self.file_path) if self.file_path else self.folder_path, "combined_data.xlsx")
            self.combined_data.to_excel(output_file, index=False)
            log.append(f"Archivo combinado creado: {output_file}")
        else:
            log.append("No se encontraron datos para combinar.")

        # Mostrar el log en la interfaz
        for entry in log:
            self.log_text.append(entry)

    def process_file(self, file_path, all_data, log):
        try:
            log.append(f"Procesando archivo: {os.path.basename(file_path)}")
            xls = pd.ExcelFile(file_path)

            # Procesar cada hoja del archivo
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
            
    def create_admin_tab(self):
        admin_widget = QWidget()
        layout = QVBoxLayout()

        # Botón para cargar el archivo combinado
        load_button = QPushButton("Cargar Archivo Combinado", self)
        load_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        load_button.clicked.connect(self.load_combined_file)
        layout.addWidget(load_button)

        # Tabla para mostrar los datos
        self.data_table = QTableWidget()
        layout.addWidget(self.data_table)

        admin_widget.setLayout(layout)
        self.tab_widget.addTab(admin_widget, "Visualización")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"Carpeta seleccionada: {folder}")
        else:
            self.folder_label.setText("Carpeta seleccionada: Ninguna")

    def combine_excel_sheets(self):
        if not self.folder_path:
            self.log_text.append("Por favor, selecciona una carpeta primero.")
            return

        all_data = []
        log = []

        # Iterar por los archivos en la carpeta
        for filename in os.listdir(self.folder_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(self.folder_path, filename)
                try:
                    log.append(f"Procesando archivo: {filename}")
                    xls = pd.ExcelFile(file_path)

                    # Procesar cada hoja del archivo
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

                    log.append(f"Archivo procesado con éxito: {filename}")
                except Exception as e:
                    log.append(f"Error al procesar el archivo {filename}: {e}")

        # Concatenar todos los DataFrames
        if all_data:
            self.combined_data = pd.concat(all_data, ignore_index=True)
            self.combined_data = self.finalize_combined_data(self.combined_data)

            # Guardar el resultado en un archivo Excel
            output_file = os.path.join(self.folder_path, "combined_data.xlsx")
            self.combined_data.to_excel(output_file, index=False)
            log.append(f"Archivo combinado creado: {output_file}")
        else:
            log.append("No se encontraron datos para combinar.")

        # Mostrar el log en la interfaz
        for entry in log:
            self.log_text.append(entry)

    def normalize_columns(self, df):
        # (El código de normalize_columns permanece igual)
        # Eliminar columnas vacías o "Unnamed"
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        # Normalización de columnas de teléfono y correo
        column_mapping = {
            "# CELULAR": "TELEFONO",
            "CELULAR": "TELEFONO",
            "CORREO ELECTRONICO": "CORREO",
            "CORREO ELECTRÓNICO": "CORREO",
            "NOMBRE 2 ": "NOMBRE 2",
        }

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
        # (El código de finalize_combined_data permanece igual)
        # Limpiar filas vacías
        if 'CEDULA' in combined_data.columns:
            # Asegurarnos de que la columna CEDULA sea numérica
            combined_data['CEDULA'] = pd.to_numeric(combined_data['CEDULA'], errors='coerce')
            # Eliminar filas donde la cédula sea NaN
            combined_data.dropna(subset=['CEDULA'], inplace=True)
            # Convertir la cédula a entero sin separadores de miles
            combined_data['CEDULA'] = combined_data['CEDULA'].astype(int).astype(str)

        if 'TELEFONO' in combined_data.columns:
            # Asegurarnos de que la columna TELEFONO sea numérica
            combined_data['TELEFONO'] = pd.to_numeric(combined_data['TELEFONO'], errors='coerce')
            # Eliminar filas donde el teléfono sea NaN
            combined_data.dropna(subset=['TELEFONO'], inplace=True)
            # Convertir el teléfono a entero sin separadores de miles
            combined_data['TELEFONO'] = combined_data['TELEFONO'].astype(int).astype(str)

        # Añadir columnas 'jornada' y 'estado_u'
        combined_data['jornada'] = combined_data['FileName'].apply(
            lambda x: 'FS' if 'FS' in x else ('DIU' if 'DIU' in x else ('NOC' if 'NOC' in x or 'ESPECIALIZA' in x else '')))
        combined_data['estado_u'] = combined_data['FileName'].apply(
            lambda x: 'Diplomado' if 'DIPLOMADO' in x else ('Tecnico' if 'TECNICO' in x else ('Profesional' if 'PROF' in x or 'DERECHO' in x else ('Especialización' if 'ESPECIALIZA' in x else ''))))

        # Reordenar las columnas según el orden que necesitas
        column_order = [
            "CEDULA",
            "APELLIDO 1",
            "APELLIDO 2",
            "NOMBRE 1",
            "NOMBRE 2",
            "TELEFONO",
            "CORREO",
            "estado_u",
            "jornada",
            "SheetName",
            "FileName",
        ]
        combined_data = combined_data[
            [col for col in column_order if col in combined_data.columns]
        ]

        return combined_data

    def load_combined_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo Combinado", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                self.combined_data = pd.read_excel(file_path)
                self.display_data_in_table()
            except Exception as e:
                self.log_text.append(f"Error al cargar el archivo: {e}")

    def display_data_in_table(self):
        if self.combined_data is not None:
            self.data_table.setRowCount(self.combined_data.shape[0])
            self.data_table.setColumnCount(self.combined_data.shape[1])
            self.data_table.setHorizontalHeaderLabels(self.combined_data.columns)

            for row in range(self.combined_data.shape[0]):
                for col in range(self.combined_data.shape[1]):
                    item = QTableWidgetItem(str(self.combined_data.iloc[row, col]))
                    self.data_table.setItem(row, col, item)

            self.data_table.resizeColumnsToContents()
    def create_1fn(self):
        fn_widget = QWidget()
        layout = QVBoxLayout()

        # Botón para cargar el archivo combinado
        load_button = QPushButton("Cargar Archivo Combinado", self)
        load_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        load_button.clicked.connect(self.load_combined_file_1fn)
        layout.addWidget(load_button)

        # Botón para aplicar la normalización 1FN
        normalize_button = QPushButton("Aplicar Normalización 1FN", self)
        normalize_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        normalize_button.clicked.connect(self.apply_1fn_normalization)
        layout.addWidget(normalize_button)

        # Tabla para mostrar los datos
        self.data_table_1fn = QTableWidget()
        layout.addWidget(self.data_table_1fn)

        fn_widget.setLayout(layout)
        self.tab_widget.addTab(fn_widget, "Forma Normal 1FN")

    def load_combined_file_1fn(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo Combinado", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                self.combined_data_1fn = pd.read_excel(file_path)
                self.display_data_in_table_1fn()
            except Exception as e:
                self.log_text.append(f"Error al cargar el archivo: {e}")

    def display_data_in_table_1fn(self):
        if self.combined_data_1fn is not None:
            self.data_table_1fn.setRowCount(self.combined_data_1fn.shape[0])
            self.data_table_1fn.setColumnCount(self.combined_data_1fn.shape[1])
            self.data_table_1fn.setHorizontalHeaderLabels(self.combined_data_1fn.columns)

            for row in range(self.combined_data_1fn.shape[0]):
                for col in range(self.combined_data_1fn.shape[1]):
                    item = QTableWidgetItem(str(self.combined_data_1fn.iloc[row, col]))
                    self.data_table_1fn.setItem(row, col, item)

            self.data_table_1fn.resizeColumnsToContents()

    def apply_1fn_normalization(self):
        if self.combined_data_1fn is not None:
            # Aplicar la normalización 1FN
            normalized_data = self.normalize_1fn(self.combined_data_1fn)
            
            # Actualizar la tabla con los datos normalizados
            self.combined_data_1fn = normalized_data
            self.display_data_in_table_1fn()
            
            # Guardar los datos normalizados
            self.save_normalized_data()

    def normalize_1fn(self, df):
        # 1. Eliminar columnas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]

        # 2. Asegurar que cada celda contenga un valor atómico
        for column in df.columns:
            if df[column].dtype == 'object':
                df[column] = df[column].apply(lambda x: str(x).split(',')[0] if isinstance(x, str) else x)

        # 3. Eliminar filas duplicadas
        df = df.drop_duplicates()

        # 4. Asegurar que cada columna tenga un nombre único
        df.columns = [f"{col}_{i}" if df.columns.tolist().count(col) > 1 else col for i, col in enumerate(df.columns)]

        return df

    def save_normalized_data(self):
        if self.combined_data_1fn is not None:
            file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Datos Normalizados", "", "Excel Files (*.xlsx)")
            if file_path:
                try:
                    self.combined_data_1fn.to_excel(file_path, index=False)
                    self.log_text.append(f"Datos normalizados guardados en: {file_path}")
                except Exception as e:
                    self.log_text.append(f"Error al guardar los datos normalizados: {e}")
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelCombinerApp()
    window.show()
    sys.exit(app.exec())