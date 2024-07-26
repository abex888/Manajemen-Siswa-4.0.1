import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout, QWidget
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QShortcut, QKeySequence

import pandas as pd
import sqlite3


class KelolaData(QMainWindow):
    def __init__(self):
        super().__init__()
        self.JendelaUtama()

    def JendelaUtama(self):
        self.setWindowTitle("Manajemen Database")
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.SubWindow)
        
        # Set window icon
        self.setWindowIcon(QIcon('smp2.ico')) 
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        self.setFixedSize(400, 210)

        # Create layout
        layout = QVBoxLayout()
        layout_button = QHBoxLayout()

        # Create UI elements
        self.excel_file_label = QLabel("Masukan file excel sumber database:")
        self.excel_file_entry = QLineEdit(self)
        self.browse_button = QPushButton("Cari File Excel")
        self.convert_button = QPushButton("Buat Database")
        #self.export_button = QPushButton("Buat Format Excel")
        self.result_label = QLabel("")
        
        # Set fixed size for buttons
        
        self.browse_button.setFixedSize(100, 30)  # Width: 200, Height: 30
        self.convert_button.setFixedSize(100, 30)  # Width: 200, Height: 30
        #self.export_button.setFixedSize(150, 30)  # Width: 200, Height: 30

        # Connect buttons to functions
        self.browse_button.clicked.connect(self.browse_file)
        self.convert_button.clicked.connect(self.ubah_data)
        #self.export_button.clicked.connect(self.buat_file_excel)
        
        #Key Binding agar ketika ESC ditekan, window tutup
        self.ESC = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.ESC.activated.connect(self.close)

        # Add elements to layout
        layout.addWidget(self.excel_file_label)
        layout.addWidget(self.excel_file_entry)
        layout.addWidget(self.result_label)
        layout_button.addWidget(self.browse_button)
        layout_button.addWidget(self.convert_button)
        #layout_button.addWidget(self.export_button)
        layout_button.addWidget(self.result_label)
        layout.addLayout(layout_button)  # Add button layout to main layout

        # Set layout
        central_widget.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Buka File Excel", "", "File Excel (*.xlsx)")
        if file_path:
            self.excel_file_entry.setText(file_path)

    import os
    def ubah_data(self):
        excel_file = self.excel_file_entry.text()
        if excel_file:
            # Get the directory of the Excel file
            excel_dir = os.path.dirname(excel_file)
            
            # Generate the database name
            db_name = os.path.splitext(os.path.basename(excel_file))[0] + '.db'
            
            # Create the path for the database file in the same directory as the Excel file
            db_path = os.path.join(excel_dir, db_name)
            
            # Convert Excel to database
            self.excel_ke_db(excel_file, db_path)
            
            # Update the result label
            self.result_label.setText(f"Database '{db_name}' telah dibuat di folder file Excel!")
        else:
            self.result_label.setText("Masukan file excel.")


    def excel_ke_db(self, excel_file, db_name):
        """
        Creates a database from an Excel file with unknown content and informs the user about the structure.
        Optionally overwrites existing database and deletes existing data.

        Args:
            excel_file (str): Path to the Excel file.
            db_name (str): Name for the database to be created.
        """
        # Read the first sheet from the Excel file
        try:
            df = pd.read_excel(excel_file, sheet_name=0)
        except FileNotFoundError:
            print(f"Kesalahan: File '{excel_file}' Tidak ditemukan.")
            return

        # Get column names from the first row (assuming headers)
        column_names = df.columns.tolist()
        column_types = df.dtypes  # Get data types for informative message

        # Connect to a new database
        conn = sqlite3.connect(db_name)

        # Delete existing data (if database already exists)
        cursor = conn.cursor()
        cursor.execute("DROP TABLE IF EXISTS data_siswa")

        # String formatting for column names with quotes
        column_names_quoted = [f'"{col}"' for col in column_names]
        create_table_sql = f"""CREATE TABLE IF NOT EXISTS data_siswa ({', '.join(column_names_quoted)})"""

        # Execute the statement
        conn.execute(create_table_sql)

        # Insert data from each row of the Excel sheet
        for index, row in df.iterrows():
            data_masukan = tuple(row.values)
            insert_sql = f"""INSERT INTO data_siswa VALUES ({', '.join(['?' for _ in data_masukan])})"""
            conn.execute(insert_sql, data_masukan)

        # Commit changes and close connection
        conn.commit()
        conn.close()

        # Create informative message about table structure
        table_info = f"Database '{db_name}' dibuat dari file '{excel_file}'.\n"
        table_info += f"Strukturnya:\n"
        for col, dtype in zip(column_names, column_types):
            table_info += f"\t- {col} ({dtype})\n"

        self.tampilkan_info(table_info)

    def buat_file_excel(self):
        # Define the columns
        columns = [
            "No", "Kelas", "No Absen", "NIPD", "NISN", "Nama Peserta Didik", 
            "L/P", "Wali Kelas", "Kelas Titipan", 
            "Tempat, Tgl Lahir", "Nama Orang Tua", "Kampung", "RT", "RW", "Desa", "Kecamatan"
        ]
        
        # Create an empty DataFrame with these columns
        df = pd.DataFrame(columns=columns)
        
        # Hardcoded filename
        filename = "data_siswa.xlsx"
        
        # Write the DataFrame to an Excel file
        df.to_excel(filename, index=False)
        pesan = QMessageBox()
        pesan.setIcon(QMessageBox.Icon.Information)
        pesan.setText(f"File '{filename}' telah dibuat.")
        pesan.setWindowTitle("Informasi")
        #print(f"File '{filename}' telah dibuat.")
        pesan.exec()

    def tampilkan_info(self, table_info):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText("Informasi Database")
        msg.setInformativeText(table_info)
        msg.setWindowTitle("Informasi Database")
        msg.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    converter = KelolaData()
    converter.show()
    sys.exit(app.exec())
