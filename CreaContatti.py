import sys
import csv
import vobject
import pandas as pd
import datetime
import os
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QVBoxLayout, QMessageBox, QFrame
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

class VCardGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Crea Contatti')
        self.setGeometry(300, 300, 500, 400)
        self.setAcceptDrops(True)  # Abilita gli eventi di drag and drop per la finestra

        layout = QVBoxLayout()

        self.label1 = QLabel('Seleziona un file XLSX o CSV da convertire in vCard')
        self.label1.setWordWrap(True)
        self.label1.setFont(QFont('Arial', 12))
        layout.addWidget(self.label1)

        self.label2 = QLabel('I file Excel accettati hanno solo 2 colonne senza riga d\'intestazione, contenenti i nomi e i numeri di telefono.')
        self.label2.setWordWrap(True)  # Abilita il word wrap nella QLabel
        self.label2.setFont(QFont('Arial', 12))
        layout.addWidget(self.label2)

        self.label3 = QLabel('Trascina e rilascia il file qui:')
        self.label3.setFont(QFont('Arial', 12))
        layout.addWidget(self.label3)

        self.dropFrame = QFrame()
        self.dropFrame.setFrameShape(QFrame.Box)
        self.dropFrame.setLineWidth(2)
        self.dropFrame.setStyleSheet("QFrame { border: 2px dashed #aaa; }")
        self.dropFrame.setMinimumHeight(150)
        layout.addWidget(self.dropFrame)

        self.button = QPushButton('Seleziona file')
        self.button.clicked.connect(self.selectFile)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def selectFile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        fileName, _ = QFileDialog.getOpenFileName(self, "Seleziona file", "", "XLSX Files (*.xlsx);;CSV Files (*.csv);;All Files (*)", options=options)
        if fileName:
            if fileName.endswith('.csv'):
                self.processCSV(fileName)
            elif fileName.endswith('.xlsx'):
                self.processXLSX(fileName)

    def processCSV(self, fileName):
        # Genera un nuovo nome di file con data e ora corrente
        current_datetime = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        vcfFileName = f'Contatti_{current_datetime}.vcf'
        
        # Ottieni il percorso del desktop dell'utente
        desktop_folder = os.path.expanduser("~/Desktop")
        vcf_file_path = os.path.join(desktop_folder, vcfFileName)

        # Cancella il contenuto del file vCard unico
        open(vcf_file_path, 'w').close()

        delimiter = ','
        try:
            with open(fileName, 'r') as csvfile:
                dialect = csv.Sniffer().sniff(csvfile.read(1024))
                delimiter = dialect.delimiter
        except:
            pass

        with open(fileName, 'r') as csvfile:
            reader = csv.reader(csvfile, delimiter=delimiter)
            for row in reader:
                self.createVCard(row[0], row[1], vcf_file_path)

        QMessageBox.information(self, 'Completato', 'Conversione vCard completata!')

    def processXLSX(self, fileName):
        # Genera un nuovo nome di file con data e ora corrente
        current_datetime = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        vcfFileName = f'Contatti_{current_datetime}.vcf'
        
        # Ottieni il percorso del desktop dell'utente
        desktop_folder = os.path.expanduser("~/Desktop")
        vcf_file_path = os.path.join(desktop_folder, vcfFileName)

        # Cancella il contenuto del file vCard unico
        open(vcf_file_path, 'w').close()

        df = pd.read_excel(fileName, header=None, names=['Nome', 'Numero'])
        for _, row in df.iterrows():
            self.createVCard(row['Nome'], row['Numero'], vcf_file_path)

        QMessageBox.information(self, 'Completato', 'Conversione vCard completata!')

    def createVCard(self, fullName, phoneNumber, vcf_file_path):
        try:
            vcard = vobject.vCard()
            vcard.add('fn').value = fullName
            tel = vcard.add('tel')
            tel.type_param = 'CELL'
            tel.value = str(phoneNumber)  # Converte phoneNumber in stringa

            with open(vcf_file_path, 'a') as vcf:
                vcf.write(vcard.serialize())
        except AttributeError:
            print(f"Errore durante la creazione della vCard per {fullName}: {phoneNumber}")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.endswith('.csv'):
                    self.processCSV(file_path)
                elif file_path.endswith('.xlsx'):
                    self.processXLSX(file_path)
            event.acceptProposedAction()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    vcardGenerator = VCardGenerator()
    vcardGenerator.show()
    sys.exit(app.exec_())
