import ast
import os
import re
import sys

from PyQt5 import uic
from PyQt5 import sip
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QColor
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from xlsxwriter.exceptions import FileCreateError


class ErrorAddReport(QDialog):
    # Окно ошибок с изменяемым тектстом
    def __init__(self, data, parent=None, flag=Qt.Dialog):
        super().__init__(parent, flag)
        uic.loadUi('ui/error_dialog_report.ui', self)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.text_error = data
        self.label_dscr_of_error.clear()
        self.label_dscr_of_error.setText(self.text_error)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.ok_btn.clicked.connect(self.ok_btn_press)

    def focusOutEvent(self, event):
        self.activateWindow()
        self.raise_()
        self.show()

    def create_log(self, log_error):
        with open('logs\logs_of_errors.txt', 'w') as f:
            f.write(log_error)

    def ok_btn_press(self):
        self.close()


class MessageDialogWindow(QtWidgets.QMessageBox):
    # Сообщение - подтверждение
    def __init__(self, title, text):
        super().__init__()
        self.title = title
        self.text_message = text
        self.msg = QtWidgets.QMessageBox(self)
        self.msg.setFocus()
        self.msg.setStyleSheet("font: 75 12pt bold \"Times New Romadn\";")

    def confirm_message(self):
        self.msg.setWindowIcon(QIcon("images/dop/attantion.png"))
        self.msg.setWindowTitle(self.title)
        self.msg.setIcon(QtWidgets.QMessageBox.Question)
        self.msg.setIconPixmap(QPixmap("images/dop/attantion.png"))
        self.msg.setText(self.text_message)
        buttonAceptar = self.msg.addButton("Да", QtWidgets.QMessageBox.YesRole)
        buttonCancelar = self.msg.addButton("Отменить", QtWidgets.QMessageBox.RejectRole)
        self.msg.setDefaultButton(buttonAceptar)
        self.msg.exec_()
        if self.msg.clickedButton() == buttonAceptar:
            return 1
        else:
            return 0

    def two_roles_confirm_message(self):
        self.msg.setWindowIcon(QIcon("images/dop/attantion.png"))
        self.msg.setWindowTitle(self.title)
        self.msg.setIcon(QtWidgets.QMessageBox.Question)
        self.msg.setIconPixmap(QPixmap("images/dop/attantion.png"))
        self.msg.setText(self.text_message)
        buttonAceptar = self.msg.addButton("Добавить", QtWidgets.QMessageBox.YesRole)
        button_accept_2 = self.msg.addButton("Перезаписать", QtWidgets.QMessageBox.AcceptRole)
        buttonCancelar = self.msg.addButton("Отменить", QtWidgets.QMessageBox.RejectRole)
        self.msg.setDefaultButton(buttonAceptar)
        self.msg.exec_()
        if self.msg.clickedButton() == buttonAceptar:
            return 1
        elif self.msg.clickedButton() == button_accept_2:
            return 2
        else:
            return 0

    def success_msg(self):
        self.msg.setWindowIcon(QIcon("images/dop/success.png"))
        self.msg.setWindowTitle(self.title)
        self.msg.setIcon(QtWidgets.QMessageBox.Question)
        self.msg.setIconPixmap(QPixmap("images/dop/success.png"))
        self.msg.setText(self.text_message)
        buttonAceptar = self.msg.addButton("Ок", QtWidgets.QMessageBox.YesRole)
        self.msg.setDefaultButton(buttonAceptar)
        self.msg.exec_()
        if self.msg.clickedButton() == buttonAceptar:
            return 1