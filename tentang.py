import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout, QWidget, QScrollArea
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QShortcut, QKeySequence

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

class Tentang(QMainWindow):
    def __init__(self):
        super().__init__()
        self.JendelaUtama()

    def JendelaUtama(self):
        self.setWindowTitle("Tentang Program")
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.SubWindow)
        
        # Set window icon
        self.setWindowIcon(QIcon('smp2.ico')) 
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        self.setFixedSize(500, 500)

        # Create layout
        layout = QVBoxLayout()
       
        # Create UI elements
        self.label_tentang = QLabel(deskripsi)
        self.label_tentang.setWordWrap(True)
        
        self.label_email = QLabel('Email: <a href="mailto:abex888@gmail">abex888@gmail.com</a>')
        self.label_github = QLabel('Github: <a href="github.com/abex888">github.com/abex888</a>')
        self.label_instagram = QLabel('Instagram: <a href="https://www.instagram.com/abex888?igsh=eWk0cTljYTVyN2Fi">abex888</a>')
        
        self.baris_gulung = QScrollArea()
        self.baris_gulung.setWidgetResizable(True)
        self.baris_gulung.setWidget(self.label_tentang)
              
        self.Tombol_OK = QPushButton("Ok")
                
        # Set fixed size for buttons
        self.Tombol_OK.setFixedSize(100, 30)
        
        # Connect buttons to functions
        self.Tombol_OK.clicked.connect(self.close)
        
        #Key Binding agar ketika ESC ditekan, window tutup
        self.ESC = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.ESC.activated.connect(self.close)

        # Add elements to layout
        #layout.addWidget(self.label_tentang)
        layout.addWidget(self.baris_gulung)
        layout.addWidget(self.label_email)
        layout.addWidget(self.label_github)
        layout.addWidget(self.label_instagram)
        layout.addWidget(self.Tombol_OK)

        # Set layout
        central_widget.setLayout(layout)
        

   

if __name__ == "__main__":
    app = QApplication(sys.argv)
    converter = Tentang()
    converter.show()
    sys.exit(app.exec())
