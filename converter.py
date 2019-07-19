from PyQt5 import QtCore, QtWidgets             # gerekli kütüphaneler ve modülleri import  ettik
from PyQt5.QtWidgets import QFileDialog , QInputDialog, QLineEdit,QMessageBox  # gerekli kütüphaneler ve modülleri import  ettik
import os,sys # gerekli kütüphaneler ve modülleri import  ettik
import pandas as pd # gerekli kütüphaneler ve modülleri import  ettik
import xlrd # gerekli kütüphaneler ve modülleri import  ettik

class Ui_Form(object):  #ana sınıfımız
    def csv_to(self): # csvden excele dönüştürmemeizde yardımcı olan method
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(None, "Lütfen CSV uzantılı dosyayı seçiniz", "","All Files (*.csv)", options=options)
        if fileName:
            text, okPressed = QInputDialog.getText(None, "Dosya ismini yazınız", "Oluşturulacak dosya ismi ne olsun?",QLineEdit.Normal, "")
            if okPressed and text != '':
                path_excel = self.Folder_path + "\{}.xlsx".format(text)  # oluşturulacak excel dosyasının tam konumunu belirtiyoruz.
                try:
                    try:
                        os.mkdir(self.Folder_path)  # os.mkdir modülü ile masaüstünde cevirilenler adlı klasör oluşturduk
                    except:
                        pass
                    data_xls = pd.read_csv(fileName, index_col=None) # burada pandas ile ilk önce csv yi okuduk
                    data_xls.to_excel(path_excel, encoding='utf-8') # okunanan csvyi excele dönüştürdük.  Dil utf-8
                except:                                                                                                                                #### hata mesajları #####
                    QMessageBox.critical(None, 'İç Kabuk Hatası',
                                         "Lütfen CSV uzantılı bir dosya seçin.  Dosyada sorun olmadığından emin olunuz.",
                                         QMessageBox.Ok)
            else:
                QMessageBox.critical(None, 'Orta Kabuk Hatası',
                                     "Lütfen oluşturulacak dosyanın ismi ne olacağını giriniz.",
                                     QMessageBox.Ok)
        else:
            QMessageBox.critical(None, 'Dış Kabuk Hatası',
                                 "Lütfen CSV uzantılı bir dosya seçin.",
                                 QMessageBox.Ok)



    def excel_to(self): # excelden  csvye dönüştürmemeizde yardımcı olan method
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(None, "Lütfen  XLSX (excel)uzantılı dosyayı seçiniz", "","All Files (*.xlsx)", options=options)        #bu üç satır sizden dosya seçme penceresini açıyor

        if fileName:
            text, okPressed = QInputDialog.getText(None, "Dosya ismini yazınız", "Oluşturulacak dosya ismi ne olsun?", QLineEdit.Normal, "")  #oluşturalacak dosyanın ismini alıyor
            if okPressed and text != '':
                path_excel = self.Folder_path + "\{}.csv".format(text)  # oluşturulacak excel dosyasının tam konumunu belirtiyoruz.
                try:
                    try:
                        os.mkdir(self.Folder_path)  # os.mkdir modülü ile masaüstünde cevirilenler adlı klasör oluşturduk
                    except:
                        pass
                    data_xls = pd.read_excel(fileName, index_col=None)  # burada pandas ile ilk önce xlsx yi okuduk
                    data_xls.to_csv(path_excel, encoding='utf-8') # okunanan xlsxyi csvye dönüştürdük.  Dil utf-8
                except:                                                                                                                                                                                 #### hata mesajları #####
                    QMessageBox.critical(None, 'İç Kabuk Hatası',
                                         "Lütfen xlsx(excel) uzantılı bir dosya seçin.  Dosyada sorun olmadığından emin olunuz.",
                                         QMessageBox.Ok)
            else:
                QMessageBox.critical(None, 'Orta Kabuk Hatası',
                                     "Lütfen oluşturulacak dosyanın ismi ne olacağını giriniz.",
                                     QMessageBox.Ok)
        else:
            QMessageBox.critical(None, 'Dış Kabuk Hatası',
                                 "Lütfen xlsx(excel) uzantılı bir dosya seçin.",
                                 QMessageBox.Ok)

                                                                                          #### penceremiz ###
    def setupUi(self, Form):
        self.path = os.path.expanduser('~')
        self.Folder_path = self.path + "\Desktop\Cevirilenler"
        Form.setObjectName("Form")
        Form.resize(381, 289)
        self.csv_to_excel = QtWidgets.QPushButton(Form)
        self.csv_to_excel.setGeometry(QtCore.QRect(25, 30, 131, 241))
        self.csv_to_excel.clicked.connect(self.csv_to)
        self.csv_to_excel.setObjectName("css_to_excel")

        self.excel_to_csv = QtWidgets.QPushButton(Form)
        self.excel_to_csv.setGeometry(QtCore.QRect(230, 30, 131, 241))
        self.excel_to_csv.clicked.connect(self.excel_to)
        self.excel_to_csv.setObjectName("excel_to_csv")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)
                                                                                     ### başlık ve buton üzerindeki yazılar  ###
    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Çevirici"))
        self.csv_to_excel.setText(_translate("Form", "Csv\'den Excel\'e"))
        self.excel_to_csv.setText(_translate("Form", "Excel\'den Csv\'ye"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())