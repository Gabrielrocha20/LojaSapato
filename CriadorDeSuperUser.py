import os.path
import sys
from os import getenv

from PyQt5.QtWidgets import QApplication, QMainWindow

from CriarSU import *
from crud import CrudLoja


class Interface(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.btnCriar.clicked.connect(self.cadastrar_SU)

    def cadastrar_SU(self):
        user = self.inputUser.text()
        password = self.inputPass.text()
        adm = self.checkAdm.isChecked()
        if (len(user) < 6) or (len(password) < 6):
            self.labelCheck.setStyleSheet('color: Red;')
            self.labelCheck.setText('Os campos precisa ter mais de 6 digitos')
            return
        cadastrar = CrudLoja(login=user, senha=password, adm=adm)
        cadastrar.cadastrar_conta()
        if cadastrar.cliente:
            self.labelCheck.setStyleSheet('color: green;')
            self.labelCheck.setText('Criado')
            return
        self.labelCheck.setStyleSheet('color: Red;')
        self.labelCheck.setText('Usuario ja existe')
        


if __name__ == '__main__':
    app = QApplication(sys.argv)
    estacionamento = Interface()
    estacionamento.show()
    app.exec()