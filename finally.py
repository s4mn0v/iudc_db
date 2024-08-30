import sys
import os
import pandas as pd
import psycopg2
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog,
    QLabel, QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem, QHBoxLayout,
    QRadioButton, QButtonGroup, QLineEdit, QFormLayout, QMessageBox, QSizePolicy, QDialogButtonBox, QComboBox, QDialog, QGridLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QCursor
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QPolygon
from PyQt6.QtCore import QPoint

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
        # Create a QPixmap for the file icon
        pixmap = QPixmap(64, 64)  # Size of the icon
        pixmap.fill(QColor('transparent'))  # Fill with transparent color

        # Create a QPainter to draw the file icon
        painter = QPainter(pixmap)

        # Set a pen for the outline (optional)
        pen = QPen(QColor('black'), 2)
        painter.setPen(pen)

        # Draw the body of the file (a rectangle)
        painter.setBrush(QColor('white'))  # Set the brush color to white for the file
        painter.drawRect(10, 10, 44, 54)  # Draw the main body of the file

        # Create the folded corner using QPolygon
        folded_corner = QPolygon([
            QPoint(42, 10),  # Top-right corner
            QPoint(54, 10),  # Right side
            QPoint(42, 22)   # Bottom-left of fold
        ])

        # Draw the folded corner (triangle)
        painter.setBrush(QColor('lightgray'))
        painter.drawPolygon(folded_corner)

        # End the painter to apply the drawing
        painter.end()

        # Set the QPixmap as an icon
        self.setWindowIcon(QIcon(pixmap))
        
        self.setStyleSheet(self.get_stylesheet())
        self.setWindowTitle("Data Manager - Consultorio Tecnologico IUDC")
        self.resize(800, 600)

        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)
        
        self.create_transformation_tab()
        self.create_admin_tab()
        self.create_1fn()
        self.crud()

        self.folder_path = ""
        self.file_path = ""
        self.combined_data = None

        self.schema = None
        
        self.conn = None

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
    
    def close_connection(self):
        if self.conn:
            self.conn.close()
            self.conn = None
            self.debug_text.append("Database connection closed.")
            QMessageBox.information(self, "Connection Closed", "Database connection has been closed.")
            
            # Show the connection form again
            for i in range(self.db_form_layout.count()): 
                widget = self.db_form_layout.itemAt(i).widget()
                if widget:
                    widget.setVisible(True)
            
            # Clear the table
            self.table.setRowCount(0)
        else:
            self.debug_text.append("No active database connection to close.")
                
    def crud(self):
        crud_widget = QWidget()
        main_layout = QHBoxLayout()

        # Formulario para los datos de conexión a la base de datos
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

        self.connect_button = QPushButton("Connect")
        self.connect_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.connect_button.clicked.connect(self.connect_to_database)
        self.db_form_layout.addWidget(self.connect_button)

        # Layout for CRUD operation buttons in columns
        button_layout = QVBoxLayout()

        self.close_button = QPushButton("Close Connection")
        self.close_button.clicked.connect(self.close_connection)
        self.close_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        button_layout.addWidget(self.close_button)

        self.connection_status = QLabel("Not Connected")
        button_layout.addWidget(self.connection_status)

        # Nuevo layout para selección de esquema y tabla
        self.select_layout = QFormLayout()

        self.schema_label = QLabel("Select Schema:")
        self.schema_combo = QComboBox()
        self.schema_combo.currentIndexChanged.connect(self.schema_selected)  # Conectar al cambio de selección
        self.select_layout.addRow(self.schema_label, self.schema_combo)

        self.table_label = QLabel("Select Table:")
        self.table_combo = QComboBox()
        self.select_layout.addRow(self.table_label, self.table_combo)

        button_layout.addLayout(self.select_layout)

        # Layout para los botones CRUD en columnas
        self.crud_buttons_layout = QHBoxLayout()

        self.load_button = QPushButton("Load Data")
        self.load_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.load_button.clicked.connect(self.load_data)
        self.crud_buttons_layout.addWidget(self.load_button)

        self.insert_button = QPushButton("Insert Data")
        self.insert_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.insert_button.clicked.connect(self.insert_data)
        self.crud_buttons_layout.addWidget(self.insert_button)

        self.update_button = QPushButton("Update Data")
        self.update_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.update_button.clicked.connect(self.update_data)
        self.crud_buttons_layout.addWidget(self.update_button)

        self.delete_button = QPushButton("Delete Data")
        self.delete_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.delete_button.clicked.connect(self.delete_data)
        self.crud_buttons_layout.addWidget(self.delete_button)

        # Añadir el layout de botones CRUD a button_layout
        button_layout.addLayout(self.crud_buttons_layout)

        # Table for showing data
        self.table = QTableWidget()
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        button_layout.addWidget(self.table)

        # Add db_form_layout and button_layout to the main layout
        main_layout.addLayout(self.db_form_layout)
        main_layout.addLayout(button_layout)

        crud_widget.setLayout(main_layout)
        self.tab_widget.addTab(crud_widget, "CRUD Operations")

    def connect_to_database(self):
        host = self.host_input.text().strip()
        port = self.port_input.text().strip()
        dbname = self.dbname_input.text().strip()
        user = self.user_input.text().strip()
        password = self.password_input.text().strip()

        try:
            self.conn = psycopg2.connect(
                host=host,
                port=port,
                dbname=dbname,
                user=user,
                password=password
            )
            self.fetch_schemas()
            
            # Ocultar los widgets del formulario de conexión
            for i in range(self.db_form_layout.count()):
                item = self.db_form_layout.itemAt(i)
                if item.widget():
                    item.widget().setVisible(False)
            
            self.connection_status.setText("Connected")
            
        except Exception as e:
            self.connection_status.setText(f"Connection Failed: {e}")

    def close_connection(self):
        if self.conn:
            self.conn.close()
            self.conn = None
            
            # Hacer visibles los widgets del formulario de conexión
            for i in range(self.db_form_layout.count()):
                item = self.db_form_layout.itemAt(i)
                if item.widget():
                    item.widget().setVisible(True)
            
            self.connection_status.setText("Not Connected")
            self.schema_combo.clear()
            self.table_combo.clear()
            self.table.clear()
    
    def fetch_schemas(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT schema_name FROM information_schema.schemata")
            schemas = cursor.fetchall()
            self.schema_combo.clear()
            for schema in schemas:
                self.schema_combo.addItem(schema[0])
            cursor.close()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to fetch schemas: {e}")

    def schema_selected(self):
        self.schema = self.schema_combo.currentText()
        self.fetch_tables()

    def fetch_tables(self):
        try:
            if not self.schema:
                return
            
            cursor = self.conn.cursor()
            query = """
            SELECT table_name 
            FROM information_schema.tables 
            WHERE table_schema = %s
            """
            cursor.execute(query, (self.schema,))
            tables = cursor.fetchall()
            self.table_combo.clear()
            for table in tables:
                self.table_combo.addItem(table[0])
            cursor.close()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to fetch tables: {e}")

    def load_data(self):
        if not self.conn:
            QMessageBox.warning(self, "Error", "Not connected to the database.")
            return

        self.schema = self.schema_combo.currentText()
        self.table_name = self.table_combo.currentText()

        if not self.table_name:
            QMessageBox.warning(self, "Error", "No table selected.")
            return

        try:
            query = f"SELECT * FROM {self.schema}.{self.table_name}"
            df = pd.read_sql_query(query, self.conn)
            self.populate_table(df)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {e}")

    def populate_table(self, df):
        self.table.setRowCount(df.shape[0])
        self.table.setColumnCount(df.shape[1])
        self.table.setHorizontalHeaderLabels(df.columns)

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                self.table.setItem(i, j, item)

    def insert_data(self):
        if not self.conn:
            QMessageBox.warning(self, "Error", "Not connected to the database.")
            return

        schema = self.schema_combo.currentText()
        table = self.table_combo.currentText()

        if not table:
            QMessageBox.warning(self, "Error", "No table selected.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Insert Data")
        form_layout = QFormLayout(dialog)

        column_names = self.get_column_names(schema, table)
        inputs = {}
        for column in column_names:
            input_field = QLineEdit()
            form_layout.addRow(column, input_field)
            inputs[column] = input_field

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, dialog)
        form_layout.addWidget(button_box)

        button_box.accepted.connect(lambda: self.perform_insert(schema, table, inputs))
        button_box.rejected.connect(dialog.reject)

        dialog.exec()

    def get_column_names(self, schema, table):
        try:
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT column_name FROM information_schema.columns WHERE table_schema = %s AND table_name = %s", (schema, table))
            columns = cursor.fetchall()
            cursor.close()
            return [col[0] for col in columns]
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to fetch columns: {e}")
            return []

    def perform_insert(self, schema, table, inputs):
        columns = list(inputs.keys())
        values = [input_field.text() for input_field in inputs.values()]

        columns_str = ", ".join(columns)
        values_str = ", ".join(["%s"] * len(values))
        query = f"INSERT INTO {schema}.{table} ({columns_str}) VALUES ({values_str})"

        try:
            cursor = self.conn.cursor()
            cursor.execute(query, values)
            self.conn.commit()
            cursor.close()
            QMessageBox.information(self, "Success", "Data inserted successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to insert data: {e}")

    def update_data(self):
        if not self.conn:
            QMessageBox.warning(self, "Error", "Not connected to the database.")
            return

        schema = self.schema_combo.currentText()
        table = self.table_combo.currentText()

        if not table:
            QMessageBox.warning(self, "Error", "No table selected.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Update Data")
        form_layout = QFormLayout(dialog)

        column_names = self.get_column_names(schema, table)
        inputs = {}
        for column in column_names:
            input_field = QLineEdit()
            form_layout.addRow(column, input_field)
            inputs[column] = input_field

        primary_key = QLineEdit()
        form_layout.addRow("Primary Key Value", primary_key)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, dialog)
        form_layout.addWidget(button_box)

        button_box.accepted.connect(lambda: self.perform_update(schema, table, inputs, primary_key.text()))
        button_box.rejected.connect(dialog.reject)

        dialog.exec()

    def perform_update(self, schema, table, inputs, primary_key_value):
        set_clause = ", ".join([f"{col} = %s" for col in inputs.keys()])
        query = f"UPDATE {schema}.{table} SET {set_clause} WHERE id = %s"  # Assuming 'id' is the primary key

        values = [input_field.text() for input_field in inputs.values()] + [primary_key_value]

        try:
            cursor = self.conn.cursor()
            cursor.execute(query, values)
            self.conn.commit()
            cursor.close()
            QMessageBox.information(self, "Success", "Data updated successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to update data: {e}")

    def delete_data(self):
        if not self.conn:
            QMessageBox.warning(self, "Error", "Not connected to the database.")
            return

        schema = self.schema_combo.currentText()
        table = self.table_combo.currentText()

        if not table:
            QMessageBox.warning(self, "Error", "No table selected.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Delete Data")
        form_layout = QFormLayout(dialog)

        primary_key = QLineEdit()
        form_layout.addRow("Primary Key Value", primary_key)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, dialog)
        form_layout.addWidget(button_box)

        button_box.accepted.connect(lambda: self.perform_delete(schema, table, primary_key.text()))
        button_box.rejected.connect(dialog.reject)

        dialog.exec()

    def perform_delete(self, schema, table, primary_key_value):
        query = f"DELETE FROM {schema}.{table} WHERE id = %s"  # Assuming 'id' is the primary key

        try:
            cursor = self.conn.cursor()
            cursor.execute(query, (primary_key_value,))
            self.conn.commit()
            cursor.close()
            QMessageBox.information(self, "Success", "Data deleted successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to delete data: {e}")
            
        
    def db_connect(self):
        try:
            self.schema = self.schema_input.text().strip() or "public"

            dbname = self.dbname_input.text().strip()
            user = self.user_input.text().strip()
            password = self.password_input.text().strip()
            host = self.host_input.text().strip()
            port = self.port_input.text().strip()

            self.conn = psycopg2.connect(
                dbname=dbname,
                user=user,
                password=password,
                host=host,
                port=port,
            )

            cursor = self.conn.cursor()
            cursor.execute(f"SET search_path TO {self.schema};")
            self.conn.commit()

            self.debug_text.append(f"Connected to the database {dbname} at {host}:{port} successfully.")

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
        event.accept()    
        
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