import sys
import sqlite3
from PyQt6.QtWidgets import (QMainWindow, QApplication, QWidget, QVBoxLayout, QTableView, QPushButton, QLineEdit, QLabel,
                             QFormLayout, QHBoxLayout, QDialog, QVBoxLayout, QGroupBox, QMenu, QMenuBar, QMessageBox,)
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QAction, QIcon, QPixmap, QKeySequence, QShortcut 
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex
import pandas as pd
import openpyxl
from library_data import KelolaData
from tentang import Tentang


deskripsi = '''MANAJEMEN SISWA SMP NEGERI 2 BOJONGPICUNG
Penulis		: Aries Aprilian (abex888@gmail.com)
Versi		: 4.0.1
Lisensi		: GNU GPL versi 3

=============================================
PENJELASAN:
Program ini membuat dan memanipulasi Database Siswa. Pada dasarnya bisa digunakan di sekolah mana saja tapi versi awal sampai dengan versi saat ini dikembangkan untuk digunakan digunakan di SMP Negeri 2 Bojongpicung

Program ini akan membuat database dari file format Microsoft Excel. Pada dasarnya modul database di file ini dapat membuat database dengan header apapun karena modulnya dibuat agar dapat membuat database terlepas dari kolom yang ada di file Excelnya. 

Namun demikian, modul utama program ini dibuat untuk database yang biasa digunakan di SMPN 2 Bojongpicung, yaitu database yang biasa ada di Data Siswa pada File Hanca SMPN 2 Bojongpicung. 

Adapun header sheet data siswa di file Hanca terdiri dari:
No, Kelas, "No Absen", NIPD, NISN, "Nama Peserta Didik", "L/P", "Wali Kelas", "Kelas Titipan", "Tempat, Tgl Lahir", "Nama Orang Tua", Kampung, RT, RW, Desa, Kecamatan, sehingga jika sekolah lain akan menggunakan Program ini, file Excel yang dibuat harus terdiri dari kolom tersebut. 

Untuk memudahkan, pada menu File -> Buat Format Excel) pengguna bisa membuat file format Excel kosong dengan kolom yang dipakai pada program ini.

'''


class DataSiswaModel(QAbstractTableModel):
    """
    A model representing student data for use in a Qt view.

    Attributes:
        _data (list of list): A 2D list containing the data to be displayed in the table.
        _headers (list of str): A list of column headers for the table.

    Methods:
        rowCount(parent=QModelIndex()):
            Returns the number of rows in the model.
        
        columnCount(parent=QModelIndex()):
            Returns the number of columns in the model.
        
        data(index, role=Qt.ItemDataRole.DisplayRole):
            Returns the data for the given role and index.
        
        headerData(section, orientation, role=Qt.ItemDataRole.DisplayRole):
            Returns the header data for the given section and orientation.
    """
    def __init__(self, data, headers, parent=None):
        """
        Initializes the model with the given data and headers.

        Args:
            data (list of list): A 2D list containing the data to be displayed.
            headers (list of str): A list of column headers.
            parent (QObject, optional): The parent object. Defaults to None.
        """
        
        super().__init__(parent)
        self._data = data
        self._headers = headers

    def rowCount(self, parent=QModelIndex()):
        """
        Returns the number of columns in the model.

        Args:
            parent (QModelIndex, optional): The parent index. Defaults to QModelIndex().

        Returns:
            int: The number of columns in the model.
        """
        
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        """
        Returns the data for the given role and index.

        Args:
            index (QModelIndex): The index of the data.
            role (Qt.ItemDataRole, optional): The role for which the data is required. 
                                               Defaults to Qt.ItemDataRole.DisplayRole.

        Returns:
            Any: The data at the given index for the given role.
        """
        
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            row = index.row()
            col = index.column()
            return self._data[row][col]
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        """
        Returns the header data for the given section and orientation.

        Args:
            section (int): The section number.
            orientation (Qt.Orientation): The orientation (horizontal or vertical).
            role (Qt.ItemDataRole, optional): The role for which the data is required.
                                               Defaults to Qt.ItemDataRole.DisplayRole.

        Returns:
            Any: The header data for the given section and orientation.
        """
        
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self._headers[section]
        return None


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.loadData()
        self.showMaximized()
        
            
    def initUI(self):
        self.setWindowTitle('Manajemen Siswa SMPN 2 Bojongpicung')
        self.setWindowIcon(QIcon('smp2.ico')) 
        self.setGeometry(100, 100, 800, 600)

        # Create central widget and layout
        central_widget = QWidget()
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Title Label
        titleLabel = QLabel('DATABASE SISWA SMPN 2 BOJONGPICUNG')
        titleLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        titleLabel.setStyleSheet("font-size: 18pt; font-weight: bold;")

        # Membuat Menu Bar and Sub Menu
        menubar = self.menuBar()
        file_menu = menubar.addMenu('&File')
        bantuan_menu = menubar.addMenu('&Bantuan')
   

        # Create Menu Actions
        kelola_action = QAction('&Buat Database    F2', self)
        export_action = QAction('Buat &Format Excel   F3', self)
        tulis_action = QAction('Export Database ke Excel    F5', self)
        keluar_action = QAction('&Keluar    F4', self)
        tentang_action = QAction('&Tentang  F1', self)

        # Connect menu action to method
        kelola_action.triggered.connect(self.kelola)
        export_action.triggered.connect(self.konfirmasi_buat_file_excel)
        tulis_action.triggered.connect(self.TulisKeExcel)
        keluar_action.triggered.connect(self.close)
        tentang_action.triggered.connect(self.tentang_program)
        
        # Tambahkan aksi ke menu
        file_menu.addAction(kelola_action)
        file_menu.addAction(export_action)
        file_menu.addAction(tulis_action)
        file_menu.addAction(keluar_action)
        bantuan_menu.addAction(tentang_action)
        

        # Search functionality
        searchLayout = QHBoxLayout()
        self.searchField = QLineEdit()
        self.searchButton = QPushButton('Cari')
        self.searchButton.clicked.connect(self.searchRecord)

        # Connect the returnPressed signal of searchField to searchButton's click
        self.searchField.returnPressed.connect(self.searchButton.click)

        searchLayout.addWidget(QLabel('Pencarian:'))
        searchLayout.addWidget(self.searchField)
        searchLayout.addWidget(self.searchButton)

        # Table View
        self.tableView = QTableView()
        self.model = QStandardItemModel()
        self.tableView.setModel(self.model)
        self.tableView.doubleClicked.connect(self.updateRecord)
        #self.tableView.doubleClicked.connect(self.onTableDoubleClicked)

        # Buttons for CRUD operations
        self.btnLoad = QPushButton('Tampilkan Semua')
        self.btnLoad.clicked.connect(self.loadData)

        self.btnAdd = QPushButton('Tambah Data')
        self.btnAdd.clicked.connect(self.addRecord)

        self.btnUpdate = QPushButton('Ubah Data')
        self.btnUpdate.clicked.connect(self.updateRecord)

        self.btnDelete = QPushButton('Hapus Data')
        self.btnDelete.clicked.connect(self.deleteRecord)

        buttonLayout = QHBoxLayout()
        buttonLayout.addWidget(self.btnLoad)
        buttonLayout.addWidget(self.btnAdd)
        buttonLayout.addWidget(self.btnUpdate)
        buttonLayout.addWidget(self.btnDelete)
        
        #binding shortcut
        
        self.F1 = QShortcut(QKeySequence(Qt.Key.Key_F1), self)
        self.F1.activated.connect(self.tentang_program)
        
        self.F2 = QShortcut(QKeySequence(Qt.Key.Key_F2), self)
        self.F2.activated.connect(self.kelola)
        
        self.F3 = QShortcut(QKeySequence(Qt.Key.Key_F3), self)
        self.F3.activated.connect(self.konfirmasi_buat_file_excel)
        
        self.F4 = QShortcut(QKeySequence(Qt.Key.Key_F4), self)
        self.F4.activated.connect(self.close)
        
        self.F5 = QShortcut(QKeySequence(Qt.Key.Key_F5), self)
        self.F5.activated.connect(self.TulisKeExcel)
        
        
        layout.addWidget(titleLabel)
        layout.addLayout(searchLayout)
        layout.addWidget(self.tableView)
        layout.addLayout(buttonLayout)
        
    def tentang_program(self):
        self.JendelaTentang = Tentang()
        self.JendelaTentang.show()
        

    def kelola(self):
        self.kd_window = KelolaData()  # sebelum dieksekusi disimpan dulu dalam variabel agar tidak dibersihkan sistem (garabage collection)
        self.kd_window.show()

    def loadData(self):
        conn = sqlite3.connect('data_siswa.db')
        query = 'SELECT Kelas, NIPD, NISN, "Nama Peserta Didik", "L/P" FROM data_siswa'
        result = conn.execute(query).fetchall()
        conn.close()

        headers = ["Kelas", "NIPD", "NISN", "NAMA", "JK"]
        self.model.clear()
        self.model.setHorizontalHeaderLabels(headers)

        for row in result:
            items = [QStandardItem(str(field)) for field in row]
            self.model.appendRow(items)

        # Adjust column widths
        self.tableView.setColumnWidth(0, 35)
        self.tableView.setColumnWidth(1, 75)
        self.tableView.setColumnWidth(2, 75)
        self.tableView.setColumnWidth(3, 990)
        self.tableView.setColumnWidth(4, 15)
        #self.tableView.resizeColumnsToContents()

    def searchRecord(self):
        search_term = self.searchField.text()
        conn = sqlite3.connect('data_siswa.db')

        # Modify the query to search by NAMA, Kelas, and NISN
        query = """SELECT Kelas, NIPD, NISN, "Nama Peserta Didik", "L/P" FROM data_siswa 
                   WHERE "Nama Peserta Didik" LIKE ? OR "Kelas Titipan" LIKE ? OR NISN LIKE ? OR Kelas LIKE ?"""
        result = conn.execute(query, (f'%{search_term}%', f'%{search_term}%', f'%{search_term}%', f'%{search_term}%')).fetchall()
        conn.close()

        headers = ["Kelas", "NIPD", "NISN", "NAMA", "JK"]
        self.model.clear()
        self.model.setHorizontalHeaderLabels(headers)

        for row in result:
            items = [QStandardItem(str(field)) for field in row]
            self.model.appendRow(items)

        # Adjust column widths
        self.tableView.setColumnWidth(0, 35)
        self.tableView.setColumnWidth(1, 75)
        self.tableView.setColumnWidth(2, 75)
        self.tableView.setColumnWidth(3, 990)
        self.tableView.setColumnWidth(4, 15)
        #self.tableView.resizeColumnsToContents()

    def addRecord(self):
        try:
            dialog = EditDialog(self, 'Tambah Data')
            if dialog.exec():
                record = dialog.getData()
                # Check record data for debugging
                print(f"Record to add: {record}")

                conn = sqlite3.connect('data_siswa.db')
                query = """INSERT INTO data_siswa (No, Kelas, "No Absen", NIPD, NISN, "Nama Peserta Didik", "L/P", "Wali Kelas", "Kelas Titipan", "Tempat, Tgl Lahir", "Nama Orang Tua", Kampung, RT, RW, Desa, Kecamatan) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

                conn.execute(query, record)
                conn.commit()
                conn.close()
                self.loadData()
        except Exception as e:
            print(f"Error adding record: {e}")

    def updateRecord(self):
        try:
            index = self.tableView.currentIndex()
            if not index.isValid():
                return

            row = index.row()
            nisn = self.model.item(row, 2).text()  # Assuming 'NISN' is in the 5th column (index 4)

            # Debugging print
            print(f"Selected NISN: {nisn}")

            # Query the database to get the full record details
            conn = sqlite3.connect('data_siswa.db')
            query = "SELECT * FROM data_siswa"
            records = conn.execute(query).fetchall()
            conn.close()

            # Find the current record index based on the selected NISN
            current_index = next((i for i, record in enumerate(records) if record[4] == nisn), -1)

            if current_index != -1:
                dialog = EditDialog(self, 'Ubah Data', records, current_index)
                if dialog.exec():
                    newRecord = dialog.getData()
                    print(f"New Record Data: {newRecord}")

                    # Ensure the list of newRecord has the correct number of elements
                    if len(newRecord) != 16:
                        raise ValueError("Data to update does not match the number of columns.")

                    conn = sqlite3.connect('data_siswa.db')
                    query = """UPDATE data_siswa 
                               SET No=?, Kelas=?, "No Absen"=?, NIPD=?, NISN=?, "Nama Peserta Didik"=?, "L/P"=?, "Wali Kelas"=?, "Kelas Titipan"=?, "Tempat, Tgl Lahir"=?, "Nama Orang Tua"=?, Kampung=?, RT=?, RW=?, Desa=?, Kecamatan=? 
                               WHERE NISN=?"""
                    
                    # Debugging print
                    print(f"Executing query: {query}")
                    print(f"With parameters: {newRecord + [newRecord[4]]}")  # Ensure correct parameter
                    
                    cursor = conn.cursor()
                    cursor.execute(query, newRecord + [newRecord[4]])  # Update record where NISN matches
                    conn.commit()
                    updated_rows = cursor.rowcount  # Get the number of rows affected
                    conn.close()
                    
                    # Debugging print
                    print(f"Rows updated: {updated_rows}")

                    if updated_rows == 0:
                        print("No rows were updated. Please check if the NISN exists in the database.")

                    self.loadData()
        except Exception as e:
            print(f"Error updating record: {e}")


    def deleteRecord(self):
        index = self.tableView.currentIndex()
        if not index.isValid():
            return

        row = index.row()
        nisn = self.model.item(row, 2).text()  # Assuming NISN
        conn = sqlite3.connect('data_siswa.db')
        query = "DELETE FROM data_siswa WHERE NISN=?"
        conn.execute(query, (nisn,))
        conn.commit()
        conn.close()
        self.loadData()
        
    def konfirmasi_buat_file_excel(self):
    
        dialog_pertanyaan = QMessageBox(self)
        dialog_pertanyaan.setWindowTitle("Anda Yakin")
        dialog_pertanyaan.setText("Apakah Anda Ingin Membuat File Excel? File data_siswa.xlsx akan ada di folder program")
        
        # Buat Tombol
        tombol_yes = dialog_pertanyaan.addButton("Yakin", QMessageBox.ButtonRole.YesRole)
        tombol_no = dialog_pertanyaan.addButton("Tidak", QMessageBox.ButtonRole.NoRole)

        # Set tombol default
        dialog_pertanyaan.setDefaultButton(tombol_yes)
        
        #Jalankan Dialog
        dialog_pertanyaan.exec()
    
        if dialog_pertanyaan.clickedButton() == tombol_yes:
            self.buat_file_excel()
            
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
        
    def onTableDoubleClicked(self, index):
        row = index.row()
        nisn = self.model.item(row, 2).text()  # Assuming 'NISN' is in the 5th column (index 4)

        # Debugging print
        print(f"Selected NISN: {nisn}")

        # Query the database to get the full record details
        conn = sqlite3.connect('data_siswa.db')
        query = "SELECT * FROM data_siswa"
        records = conn.execute(query).fetchall()
        conn.close()

        # Debugging print
        print(f"Records from DB: {records}")

        # Find the current record index based on the selected NISN
        current_index = next((i for i, record in enumerate(records) if record[4] == nisn), -1)

        if current_index != -1:
            dialog = EditDialog(self, 'Ubah Data', records, current_index)
            if dialog.exec():
                newRecord = dialog.getData()
                print(f"New Record Data: {newRecord}")

                # Ensure the list of newRecord has the correct number of elements
                if len(newRecord) != 16:
                    raise ValueError("Data to update does not match the number of columns.")

                conn = sqlite3.connect('data_siswa.db')
                query = """UPDATE data_siswa 
                           SET No=?, Kelas=?, "No Absen"=?, NIPD=?, NISN=?, "Nama Peserta Didik"=?, "L/P"=?, "Wali Kelas"=?, "Kelas Titipan"=?, "Tempat, Tgl Lahir"=?, "Nama Orang Tua"=?, Kampung=?, RT=?, RW=?, Desa=?, Kecamatan=?  
                           WHERE NISN=?"""
                
                # Debugging print
                print(f"Executing query: {query}")
                print(f"With parameters: {newRecord + [newRecord[4]]}")  # Ensure correct parameter
                
                cursor = conn.cursor()
                cursor.execute(query, newRecord + [newRecord[4]])  # Update record where NISN matches
                conn.commit()
                updated_rows = cursor.rowcount  # Get the number of rows affected
                conn.close()
                
                # Debugging print
                print(f"Rows updated: {updated_rows}")

                if updated_rows == 0:
                    print("No rows were updated. Please check if the NISN exists in the database.")

                self.loadData()

    def TulisKeExcel(self):
        
        nama_file = "data_siswa_isi.xlsx"
        dialog_pertanyaan = QMessageBox(self)
        dialog_pertanyaan.setWindowTitle("Anda Yakin")
        dialog_pertanyaan.setText("Apakah Anda Ingin Membuat File Excel dari database ini? File data_siswa_isi.xlsx akan ada di folder program")
        
        # Buat Tombol
        tombol_yes = dialog_pertanyaan.addButton("Yakin", QMessageBox.ButtonRole.YesRole)
        tombol_no = dialog_pertanyaan.addButton("Tidak", QMessageBox.ButtonRole.NoRole)

        # Set tombol default
        dialog_pertanyaan.setDefaultButton(tombol_yes)
        
        #Jalankan Dialog
        dialog_pertanyaan.exec()
    
        if dialog_pertanyaan.clickedButton() == tombol_yes:
            self.ExportKeExcel(nama_file)
        
    def ExportKeExcel(self, nama_file):
        # Koneksi ke SQLite database
        conn = sqlite3.connect('data_siswa.db')
        
        # Ambil data dari database masukan ke DataFrame pandas
        df = pd.read_sql_query("SELECT * FROM data_siswa", conn)
        
        # Tulis DataFrame ke File Excel
        df.to_excel(nama_file, index=False)
        
        # Close the database connection
        conn.close()



class EditDialog(QDialog):
    def __init__(self, parent=None, title='', records=None, current_index=0):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.records = records
        self.current_index = current_index
        self.initUI()
        self.updateUI()

    def initUI(self):
        layout = QFormLayout()

        # Define the headers that match the database columns
        self.headers = ["No", "Kelas", "No Absen", "NIPD", "NISN", "Nama Peserta Didik", "L/P", "Wali Kelas", "Kelas Titipan", "Tempat, Tgl Lahir", "Nama Orang Tua", "Kampung", "RT", "RW", "Desa", "Kecamatan"]

        # Create input fields for each header
        self.inputs = {}
        for header in self.headers:
            label = QLabel(header)
            inputField = QLineEdit()
            self.inputs[header] = inputField
            layout.addRow(label, inputField)

        # Navigation buttons
        btnPrevious = QPushButton('Previous')
        btnPrevious.setFixedSize(125, 40)  
        btnPrevious.clicked.connect(self.showPreviousRecord)
        btnNext = QPushButton('Next')
        btnNext.clicked.connect(self.showNextRecord)
        btnNext.setFixedSize(125, 40)
        navigationLayout = QHBoxLayout()
        navigationLayout.addWidget(btnPrevious)
        navigationLayout.addWidget(btnNext)
        layout.addRow(navigationLayout)

        # Save button
        btnSave = QPushButton('Simpan')
        btnSave.clicked.connect(self.accept)
        btnSave.setFixedSize(125, 40)
        #layout.addWidget(btnSave)
        
        # Tombol Cancel
        btnCancel = QPushButton('Batal')
        btnCancel.clicked.connect(self.close)
        btnCancel.setFixedSize(125, 40)
        #layout.addWidget(btnCancel)
        
        #Save Cancel Layout
        posisiSaveCancel = QHBoxLayout()
        posisiSaveCancel.addWidget(btnSave)
        posisiSaveCancel.addWidget(btnCancel)
        layout.addRow(posisiSaveCancel)
        
        
        
        self.setLayout(layout)

    def updateUI(self):
        if self.records and 0 <= self.current_index < len(self.records):
            record = self.records[self.current_index]
            for header, value in zip(self.headers, record):
                self.inputs[header].setText(str(value))  # Ensure value is a string

    def showPreviousRecord(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.updateUI()

    def showNextRecord(self):
        if self.current_index < len(self.records) - 1:
            self.current_index += 1
            self.updateUI()

    def getData(self):
        # Retrieve the data from input fields
        return [self.inputs[header].text() for header in self.headers]



        
    





if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    print("PROGRAM SELESAI DIJALANKAN, TERIMA KASIH")
    sys.exit(app.exec())
