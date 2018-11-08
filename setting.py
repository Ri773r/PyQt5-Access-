# -*- coding: utf-8 -*-

"""
Module implementing Dialog.
"""

from PyQt5.QtCore import pyqtSlot, pyqtSignal
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox
from Ui_setting import Ui_Dialog


class Dialog(QDialog, Ui_Dialog):
    sinOut = pyqtSignal(str)
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        self.setupUi(self)
        self.fileName = None
    
    @pyqtSlot()
    def on_pushButton_clicked(self):
        self.fileName, _ = QFileDialog.getOpenFileName(self, '选择文件', '.','Access DataBase(*.mdb *.accdb)')
        self.lineEdit.setText(self.fileName.replace('/', '\\'))
    
    @pyqtSlot()
    def on_pushButton_2_clicked(self):
        if self.lineEdit.text():
            self.sinOut.emit(self.fileName)
            self.close()
        else:
            QMessageBox.warning(self, '警告', '链接为空', QMessageBox.Ok)
