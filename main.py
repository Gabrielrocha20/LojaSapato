import os.path
import sys
from os import getenv

import pandas as pd
import PyPDF2
import win32api
import win32print
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (QApplication, QDialog, QMainWindow, QRadioButton,
                             QTreeWidget, QTreeWidgetItem)
from reportlab.lib.pagesizes import A7
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

import errordialog as ed
from crud import CrudLoja
from dialogo import *
from interface import *


class ErrorDialogBox(QDialog, ed.Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

class DialogBox(QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.check = False

        self.btnImprimirLogar.clicked.connect(self.loginImprimir)

    def loginImprimir(self):
        try:
            login = self.inputImprimirLogin.text()
            senha = self.inputImprimirSenha.text()

            logar = CrudLoja(login=login, senha=senha, funcionario=True)
            logar.read_funcionario()
            resultado = logar.resultados
            
            if len(resultado) == 1:
                self.login = login
                self.senha = senha
                self.check = True
                self.close()
                self.inputImprimirSenha.setText('')
                return
            self.labelCheck.setStyleSheet('color: rgb(255, 0, 0)')
            self.labelCheck.setText('Login Invalido')
        except TypeError as e:
            pass

class Interface(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.username = getenv("USERNAME")

        # Paginas Botoes
        # self.CadastrarClienteBtnPag.clicked.connect(lambda: self.PaginaCentral.setCurrentWidget(self.CadastroClientePage))
        self.btnPageCadastro.clicked.connect(lambda: self.PaginaCentral.setCurrentWidget(self.pageCadastro))
        self.btnPagVoltar.clicked.connect(lambda: self.PaginaCentral.setCurrentWidget(self.pageHome))

        self.CadastrarClienteBtnPag.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.CadastroClientePage))
        self.CadastrarProdutoBtnPag.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.CadastroProdutoPage))
        self.ClientesProdutosBtnPag.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.ClientesProdutosPage))
        self.ClienteProdutoBtnPag.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.ClienteProdutoPage))
        self.FinalizarProdutoBtnPag.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.FinalizarPage))

        

        # Botões
        self.btnEntrar.clicked.connect(self.logar_usuario)
        self.btnFuncionarioCadastrar.clicked.connect(self.cadastrar_usuario)
        self.btnSair.clicked.connect(self.logout)


        self.btnCadastrarCliente.clicked.connect(self.cadastrar_cliente)
        self.btnCadastrarImprimir.clicked.connect(self.cadastrar_produto)
        self.btnCheckLoginImprimir.clicked.connect(self.checarloginParaImprimir)

        self.btnClientes.clicked.connect(self.mostrar_todos_clientes)
        self.btnProdutos.clicked.connect(self.mostrar_todos_produtos)
        self.btnFiltro.clicked.connect(self.abrirfiltros)
        self.btnFiltrar.clicked.connect(self.filtrar)

        self.btnCliente.clicked.connect(self.mostrar_cliente)
        self.btnProduto.clicked.connect(self.mostrar_produto)

        self.btnPesquisarFinalizar.clicked.connect(self.pesquisa_finalizar)
        self.btnFinalizar.clicked.connect(self.finalizar_produto)

        # user
        self.login = None
        self.senha = None
        
        self.dialogo = DialogBox()
        self.errorDialogo = ErrorDialogBox()
    
    def abrirfiltros(self):
        try:
            width = self.frame_25.width()

            if width == 0:
                newWidth = 200
            else:
                newWidth = 0
            
            self.animation = QtCore.QPropertyAnimation(self.frame_25, b'minimumWidth')
            self.animation.setDuration(1000)
            self.animation.setStartValue(width)
            self.animation.setEndValue(newWidth)
            self.animation.setEasingCurve(QtCore.QEasingCurve.InOutQuart)
            self.animation.start()
        except TypeError as e:
            pass
    
    def logar_usuario(self):
        try:
            login = self.inputLogin.text()
            senha = self.inputSenha.text()

            logar = CrudLoja(login=login, senha=senha, funcionario=False)
            logar.read_funcionario()
            resultado = logar.resultados
            
            if len(resultado) == 1:
                self.login = login
                self.senha = senha
                self.labelNomeFuncionario.setText(self.login)
                self.inputLogin.setText('')
                self.inputSenha.setText('')
                if resultado[0][3] == 'True':
                    self.btnPageCadastro.setEnabled(True)
                    self.btnPageCadastro.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
                return self.PaginaCentral.setCurrentWidget(self.pageHome)
            self.labelLoginMensageBox.setText('Usuario ou Senha invalidos')
        except TypeError as e:
            pass

    def cadastrar_usuario(self):
        try:
            login = self.inputFuncionarioLogin.text()
            senha = self.inputFuncionarioSenha.text()
            telefone = self.inputFuncionarioTelefone.text()
            error = 0
            if len(login) < 6:
                self.labelCadastroMensageBoxLogin.setText('Usuario precisa ter mais de 6 letras')
                error = 1
            if len(senha) < 6:
                self.labelCadastroMensageBoxSenha.setText('Senha precisa ter mais de 6 letras')
                error = 1
            if len(telefone) != 11:
                self.labelCadastroMensageBoxTelefone.setText('Telefone precisa ter exatamente 11 digitos')
                error = 1
            if not telefone.isdigit():
                self.labelCadastroMensageBoxTelefone_2.setText('Telefone precisa ser digito')
                error = 1
            if error == 1:

                return
            cadastrar = CrudLoja(login=login, senha=senha, telefone=telefone)
            cadastrar.cadastrar_funcionario()
            self.labelNomeFuncionario.setText(self.login)
            return self.PaginaCentral.setCurrentWidget(self.pageHome)
        except TypeError as e:
            pass
    
    def logout(self):
        try:
            sair = CrudLoja(login=self.login, senha=self.senha)
            sair.update_saida_funcionario()
            self.login = None
            self.senha = None
            self.btnPageCadastro.setEnabled(False)
            self.btnFuncionarioCadastrar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            return self.PaginaCentral.setCurrentWidget(self.pageLogin)
        except TypeError as e:
            pass
        
    def checarloginParaImprimir(self):
        self.dialogo.show()

    def cadastrar_cliente(self):
        try:
            nome_cliente = self.InputNomeCliente.text()
            telefone_cliente = self.InputTelefoneCliente.text()
            
            cadastrar = CrudLoja(nome=nome_cliente, telefone=telefone_cliente, clientes='1')
            cadastrar.create_cliente()
            cadastrar.mostrar_clientes()

            resultado = cadastrar.resultados[-1]

            font=QtGui.QFont() 
            font.setPointSize(12)        

            child = QTreeWidgetItem(self.colunCadastrarCliente, [f'{resultado[2]}', f'{resultado[0]}', f'{resultado[1]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)
        except TypeError as e:
            pass
    
    def cadastrar_produto(self):
        try:
            self.labelCadastrarProduto.clear()
            check = self.dialogo.check
            if not check:
                self.errorDialogo.show()
                self.errorDialogo.labelErrorDialogo.setText('Precisa Logar para Cadastrar produto')
                return
            self.dialogo.check = False

            funcionario = self.dialogo.inputImprimirLogin.text()
            servico_cliente = self.InputServico.text()
            produto_cliente = self.InputProduto.text()
            telefone_cliente = self.InputTelefoneProduto.text()
            preco_cliente = self.InputPreco.text()
            sinal = self.InputSinal.text()
            prazo = self.InputPrazo.text()

            radios = [self.radioBranco, self.radioPreto, self.radioAzul, self.radioVermelho, self.radioMarron,
            self.radioRosa, self.radioVerde, self.radioAmarelo, self.radioBege, self.radioLaranja, self.radioRoxo, self.radioBicolor,]

            radios_parpe = [self.radioPar, self.radioPe]

            if sinal == '':
                sinal = 0
            if not sinal.isdigit():
                return
            sinal = float(sinal)

            cor_cliente = []
            for radio in radios:
                if radio.isChecked():
                    cor_cliente.append(radio.text())
            par_pe_cliente = ''
            for radio_parpe in radios_parpe:
                if radio_parpe.isChecked():
                    par_pe_cliente = radio_parpe.text()

            preco_lista = preco_cliente.split('/')
            preco = 0
            for p in preco_lista:
                preco += float(p.replace(',','.'))
            preco_sinal = preco - sinal
            cor_cliente = ','.join(cor_cliente)
            cadastrar = CrudLoja(produto=produto_cliente, 
            servico=servico_cliente, cor=cor_cliente, 
            telefone=telefone_cliente, par_pe= par_pe_cliente, 
            preco=preco, prazo=prazo, sinal=sinal, 
            funcionario=funcionario, produtos='1')
            cadastrar.create_produto()
            cadastrar.mostrar_produtos()

            if len(cadastrar.resultados) == 0:
                child = QTreeWidgetItem(self.labelCadastrarProduto, ['Cliente', 'nao existe', 'ou Telefone', 'incorreto'])
                return

            resultado = cadastrar.resultados[-1]
            servicos = resultado[3].split('/')
            
            font=QtGui.QFont() 
            font.setPointSize(12)        
            child = QTreeWidgetItem(self.labelCadastrarProduto, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)

            diretorio = fr'C:\Users\{self.username}\Desktop'
            if os.path.exists(diretorio):
                pass
            else:
                diretorio = fr'C:\Users\{self.username}\OneDrive\Área de Trabalho'
            notaLoja = diretorio + '\\notaLoja.pdf'
            notaCliente = diretorio + '\\notaCliente.pdf'
            
            cnv = canvas.Canvas(notaCliente)
            cnv.setFontSize(15)
            cnv.scale(0.9, 0.65)
            
            cnv.drawString(0, 1275, f'OS:       {resultado[0]}       ')
            cnv.setFontSize(10)
            cnv.drawString(0, 1260, f'RÁPIDO DOS CALÇADOS            Cliente')
            cnv.drawString(0, 1245, f'Praça D.Pedro, 31-LOJA 5/Centro')
            cnv.drawString(0, 1230, f'Petrópolis')
            cnv.drawString(0, 1215, f'Telef.: (00) 0000-0000')
            cnv.drawString(0, 1200, f'Horário: de Seg. a Sex. das 09:00 as 18:00')
            cnv.drawString(0, 1185, f'____________________________________________')
            cnv.drawString(0, 1170, f'{resultado[10]} {resultado[12]}')
            cnv.drawString(0, 1155, f'____________________________________________')
            cnv.setFontSize(12)
            cnv.drawString(0, 1140, f'O.S: {resultado[0]}')
            cnv.setFontSize(10)
            cnv.drawString(0, 1125, f'{resultado[6]}   Prazo: {resultado[5]}')
            cnv.drawString(0, 1110, f'____________________________________________')
            cnv.drawString(0, 1095, f'Impresso por: {funcionario}')
            cnv.drawString(0, 1080, f'____________________________________________')
            cnv.drawString(0, 1065, f'It Objeto    Cor: {resultado[2]}')
            cnv.drawString(0, 1050, f'Serviço: {resultado[3]}')
            cnv.drawString(0, 1035, f'____________________________________________')
            cnv.drawString(0, 1020, f'{resultado[4]} {resultado[1]} {resultado[2]}')
            contador = 0
            y = 1005
            for servico in servicos:
                if len(preco_lista) > 1:
                    cnv.drawString(0, y, f'                    {servico}           R${preco_lista[contador]}')
                else:
                    cnv.drawString(0, y, f'                    {servico}           R${preco_cliente}')
                y -= 15
                contador += 1
            cnv.drawString(0, y - 15, f' ___________________________________________')
            cnv.drawString(0, y - 30, f'                     SubTotal: R$ {preco}')
            if sinal == 0:
                cnv.drawString(0, y - 45, f'                     Sinal:    R$ 00.00')
                cnv.drawString(0, y - 60, f'                     Total:    R$ {preco}')
            else:
                cnv.drawString(0, y - 45, f'                     Sinal:    R$ {sinal}')
                cnv.drawString(0, y - 60, f'                     Total:    R$ {preco_sinal}')
            cnv.drawString(0, y - 75, f' ___________________________________________')
            cnv.drawString(0, y - 90, f' Prezado cliente! Preserve este Documento')
            cnv.drawString(0, y - 105, f' pois será através dele, que nossos')
            cnv.drawString(0, y - 120, f' funcionarios os indentificarão os objetos')
            cnv.drawString(0, y - 135, f' aqui deixados.')
            cnv.drawString(0, y - 150, f' ___________________________________________')
            cnv.drawString(0, y - 165, f'     NÃO NOS RESPONSABILIZAMOS POR')
            cnv.drawString(0, y - 180, f'     OBJETOS DEIXADOS POR MAIS')
            cnv.drawString(0, y - 195, f'         DE 30 DIAS.')
            cnv.drawString(0, y - 210, f' DATA: __/__/____     Ass.:__________________')
            cnv.setPageSize((300, 840))
            cnv.save()

            cnvloja = canvas.Canvas(notaLoja)
            cnvloja.setFontSize(15)
            cnvloja.scale(0.9, 0.65)
            cnvloja.drawString(0, 1275, f'OS:       {resultado[0]}       ')
            cnvloja.setFontSize(10)
            cnvloja.drawString(0, 1260, f'RÁPIDO DOS CALÇADOS            Balcão')
            cnvloja.drawString(0, 1245, f'Praça D.Pedro, 31-LOJA 5/Centro')
            cnvloja.drawString(0, 1230, f'Petrópolis')
            cnvloja.drawString(0, 1215, f'Telef.: (00) 0000-0000')
            cnvloja.drawString(0, 1200, f'Horário: de Seg. a Sex. das 09:00 as 18:00')
            cnvloja.drawString(0, 1185, f'____________________________________________')
            cnvloja.drawString(0, 1170, f'{resultado[10]} {resultado[12]}')
            cnvloja.drawString(0, 1155, f'____________________________________________')
            cnvloja.setFontSize(12)
            cnvloja.drawString(0, 1140, f'O.S: {resultado[0]}')
            cnvloja.setFontSize(10)
            cnvloja.drawString(0, 1125, f'{resultado[6]}   Prazo: {resultado[5]}')
            cnvloja.drawString(0, 1110, f'____________________________________________')
            cnvloja.drawString(0, 1095, f'Impresso por: {funcionario}')
            cnvloja.drawString(0, 1080, f'____________________________________________')
            cnvloja.drawString(0, 1065, f'It Objeto    Cor: {resultado[2]}')
            cnvloja.drawString(0, 1050, f'Serviço: {resultado[3]}')
            cnvloja.drawString(0, 1035, f'____________________________________________')
            cnvloja.drawString(0, 1020, f'{resultado[4]} {resultado[1]} {resultado[2]}')
            contador = 0
            y = 1005
            for servico in servicos:
                if len(preco_lista) > 1:
                    cnvloja.drawString(0, y, f'                    {servico}           R${preco_lista[contador]}')
                else:
                    cnvloja.drawString(0, y, f'                    {servico}           R${preco_cliente}')
                y -= 15
                contador += 1
            cnvloja.drawString(0, y - 15, f' ___________________________________________')
            cnvloja.drawString(0, y - 30, f'                     SubTotal: R$ {preco}')
            if sinal == 0:
                cnvloja.drawString(0, y - 45, f'                     Sinal:    R$ 00.00')
                cnvloja.drawString(0, y - 60, f'                     Total:    R$ {preco}')
            else:
                cnvloja.drawString(0, y - 45, f'                     Sinal:    R$ {sinal}')
                cnvloja.drawString(0, y - 60, f'                     Total:    R$ {preco_sinal}')
            cnvloja.drawString(0, y - 75, f' ___________________________________________')
            cnvloja.drawString(0, y - 90, f' Prezado cliente! Preserve este Documento')
            cnvloja.drawString(0, y - 105, f' pois será através dele, que nossos')
            cnvloja.drawString(0, y - 120, f' funcionarios os indentificarão os objetos')
            cnvloja.drawString(0, y - 135, f' aqui deixados.')
            cnvloja.drawString(0, y - 150, f' ___________________________________________')
            cnvloja.drawString(0, y - 165, f'     NÃO NOS RESPONSABILIZAMOS POR')
            cnvloja.drawString(0, y - 180, f'     OBJETOS DEIXADOS POR MAIS')
            cnvloja.drawString(0, y - 195, f'         DE 30 DIAS.')
            cnvloja.drawString(0, y - 210, f' DATA: __/__/____     Ass.:__________________')
            cnvloja.setPageSize((300, 840))
            cnvloja.save()
            # impressora = win32print.EnumPrinters(2)[1]
            # win32print.SetDefaultPrinter(impressora[2])
            # win32api.ShellExecute(0, "print", notaCliente, None, diretorio[:-16], 0)

            # impressora = win32print.EnumPrinters(2)[1]
            # win32print.SetDefaultPrinter(impressora[2])
            # win32api.ShellExecute(0, "print", notaLoja, None, notaLoja[:-13], 0)
        except TypeError as e:
            pass


    def mostrar_todos_produtos(self):
        try:
            self.labelClientesProdutos.clear()
            self.labelClientesProdutos.headerItem().setText(0, "O.S")
            self.labelClientesProdutos.headerItem().setText(1, "Produto")
            self.labelClientesProdutos.headerItem().setText(2, "Cor")
            self.labelClientesProdutos.headerItem().setText(3, "Serviço")
            self.labelClientesProdutos.headerItem().setText(4, "Par/Pé")
            self.labelClientesProdutos.headerItem().setText(5, "Data Prazo")
            self.labelClientesProdutos.headerItem().setText(6, "Hora Entrada")
            self.labelClientesProdutos.headerItem().setText(7, "Hora Saida")
            self.labelClientesProdutos.headerItem().setText(8, "Preço")
            self.labelClientesProdutos.headerItem().setText(9, "Sinal")
            self.labelClientesProdutos.headerItem().setText(10, "Nome")
            self.labelClientesProdutos.headerItem().setText(11, "Funcionario")
            produtos = CrudLoja(produtos='todos')
            produtos.mostrar_produtos()
            resultados = produtos.resultados

            font=QtGui.QFont() 
            font.setPointSize(12) 
            for resultado in resultados:
                child = QTreeWidgetItem(self.labelClientesProdutos, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
                for i, _ in enumerate(resultado):
                    child.setFont(i , font)
            self.create_file_excel(lista_dados=resultados, tabela='Produtos')
        except TypeError as e:
            pass
        
    def filtrar(self):
        try:
            self.labelClientesProdutos.clear()
            self.labelClientesProdutos.headerItem().setText(0, "O.S")
            self.labelClientesProdutos.headerItem().setText(1, "Produto")
            self.labelClientesProdutos.headerItem().setText(2, "Cor")
            self.labelClientesProdutos.headerItem().setText(3, "Serviço")
            self.labelClientesProdutos.headerItem().setText(4, "Par/Pé")
            self.labelClientesProdutos.headerItem().setText(5, "Data Prazo")
            self.labelClientesProdutos.headerItem().setText(6, "Hora Entrada")
            self.labelClientesProdutos.headerItem().setText(7, "Hora Saida")
            self.labelClientesProdutos.headerItem().setText(8, "Preço")
            self.labelClientesProdutos.headerItem().setText(9, "Sinal")
            self.labelClientesProdutos.headerItem().setText(10, "Nome")
            self.labelClientesProdutos.headerItem().setText(11, "Funcionario")
            filtro = self.inputFiltro.text()
            produtos = CrudLoja(filtro=filtro, produtos='filtro')
            produtos.mostrar_produtos()
            resultados = produtos.resultados 

            font=QtGui.QFont() 
            font.setPointSize(12) 
            for resultado in resultados:
                child = QTreeWidgetItem(self.labelClientesProdutos, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
                for i, _ in enumerate(resultado):
                    child.setFont(i , font)
        except TypeError as e:
            pass
        
    def mostrar_todos_clientes(self):
        try:
            self.labelClientesProdutos.clear()
            self.labelClientesProdutos.setColumnCount(3)
            self.labelClientesProdutos.headerItem().setText(0, "ID")
            self.labelClientesProdutos.headerItem().setText(1, "Nome")
            self.labelClientesProdutos.headerItem().setText(2, "Telefone")
            
            clientes = CrudLoja(clientes='todos')
            clientes.mostrar_clientes()
            resultados = clientes.resultados

            font=QtGui.QFont() 
            font.setPointSize(12) 
            for resultado in resultados:
                child = QTreeWidgetItem(self.labelClientesProdutos, [f'{resultado[2]}', f'{resultado[0]}', f'{resultado[1]}'])
                for i, _ in enumerate(resultado):
                    child.setFont(i , font)
            self.create_file_excel(lista_dados=resultados, tabela='Clientes')
        except TypeError as e:
            pass


    def mostrar_produto(self):
        try:
            self.labelClienteProduto.clear()
            self.labelClienteProduto.headerItem().setText(0, "O.S")
            self.labelClienteProduto.headerItem().setText(1, "Produto")
            self.labelClienteProduto.headerItem().setText(2, "Cor")
            self.labelClienteProduto.headerItem().setText(3, "Serviço")
            self.labelClienteProduto.headerItem().setText(4, "Par/Pé")
            self.labelClienteProduto.headerItem().setText(5, "Data Prazo")
            self.labelClienteProduto.headerItem().setText(6, "Hora Entrada")
            self.labelClienteProduto.headerItem().setText(7, "Hora Saida")
            self.labelClienteProduto.headerItem().setText(8, "Preço")
            self.labelClienteProduto.headerItem().setText(9, "Sinal")
            self.labelClienteProduto.headerItem().setText(10, "Nome")

            telefone_cliente = self.InputTelefoneClienteProduto.text()
            produtos = CrudLoja(telefone=telefone_cliente, produtos='1')
            produtos.mostrar_produtos()
            resultados = produtos.resultados
            if len(resultados) > 0:
                font=QtGui.QFont() 
                font.setPointSize(12) 
                for resultado in resultados:
                    child = QTreeWidgetItem(self.labelClienteProduto, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
                    for i, _ in enumerate(resultado):
                        child.setFont(i , font)
        except TypeError as e:
            pass
    def mostrar_cliente(self):
        try:
            self.labelClienteProduto.clear()
            self.labelClienteProduto.headerItem().setText(0, "ID")
            self.labelClienteProduto.headerItem().setText(1, "Telefone")
            self.labelClienteProduto.headerItem().setText(2, "Nome")
            self.labelClienteProduto.setColumnCount(3)
            telefone_cliente = self.InputTelefoneClienteProduto.text()
            cliente = CrudLoja(telefone=telefone_cliente, clientes='1')
            cliente.mostrar_clientes()

            resultados = cliente.resultados
            
            if len(resultados) > 0:
                font=QtGui.QFont() 
                font.setPointSize(12) 
                for resultado in resultados:
                    child = QTreeWidgetItem(self.labelClienteProduto, [f'{resultado[2]}', f'{resultado[0]}', f'{resultado[1]}'])
                    for i, _ in enumerate(resultado):
                        child.setFont(i , font)
            else:
                self.labelClienteProduto.headerItem().setText(0, "ID")
                self.labelClientesProdutos.setColumnCount(1)
                child = QTreeWidgetItem(self.labelClienteProduto, [f'Não encontrado'])
        except TypeError as e:
            pass
    
    def pesquisa_finalizar(self):
        try:
            self.LabelFinalizar.clear()
            os_produto = self.inputFinalizar.text()
            if not os_produto.isdigit():
                child = QTreeWidgetItem(self.LabelFinalizar, ['Apenas Numero'])
                return
            os_produto = int(os_produto)

            pesquisa_produto = CrudLoja(o_s=os_produto, produtos='os')
            pesquisa_produto.mostrar_produtos()
            resultado = pesquisa_produto.resultados
            if len(resultado) > 0:
                resultado = resultado[-1]
                font=QtGui.QFont() 
                font.setPointSize(12)
                child = QTreeWidgetItem(self.LabelFinalizar, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
                for i, _ in enumerate(resultado):
                    child.setFont(i , font)
                return
            child = QTreeWidgetItem(self.LabelFinalizar, ['O.S Incorreto ou nao existe'])
            return
        except TypeError as e:
            pass
        

    def finalizar_produto(self):
        try:
            self.LabelFinalizar.clear()
            os_produto = self.inputFinalizar.text()
            if not os_produto.isdigit():
                child = QTreeWidgetItem(self.LabelFinalizar, ['Apenas Numero'])
                return
            os_produto = int(os_produto)

            finalizar = CrudLoja(o_s=os_produto, produtos='os')
            finalizar.update()
            finalizar.mostrar_produtos()
            resultado = finalizar.resultados
            if not finalizar.check_os:
                child = QTreeWidgetItem(self.LabelFinalizar, ['O.S Incorreto ou nao existe'])
                return
            
            resultado = resultado[-1]
            font=QtGui.QFont() 
            font.setPointSize(12)
            child = QTreeWidgetItem(self.LabelFinalizar, [
                f'{resultado[0]}', f'{resultado[1]}', 
                f'{resultado[2]}', f'{resultado[3]}', f'{resultado[4]}', 
                f'{resultado[5]}', f'{resultado[6]}', f'{resultado[7]}', 
                f'{resultado[8]}', f'{resultado[9]}', f'{resultado[10]}', 
                f'{resultado[11]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)
            child = QTreeWidgetItem(self.LabelFinalizar, ['Finalizado!!'])
            return
        except TypeError as e:
            pass
    
    def create_file_excel(self, lista_dados, tabela):
        try:
            arquivo = fr'C:\Users\{self.username}\Desktop'
            if os.path.exists(arquivo):
                pass
            else:
                arquivo = fr'C:\Users\{self.username}\OneDrive\Área de Trabalho'
            if tabela == 'Clientes':
                Nome = []
                Telefone = []
                ID_cliente = []
                for row in lista_dados:
                    Nome.append(row[0])
                    Telefone.append(row[1])
                    ID_cliente.append(row[2])
                dict_dados = {
                "ID": ID_cliente, "Nome": Nome, "Telefone": Telefone}
                dados = pd.DataFrame(data=dict_dados)
                arquivo = arquivo + '\Clientes.xls'
                dados.to_excel(arquivo)
                return
            else:
                id_produto = []
                produto = []
                cor = []
                servico = []
                par_pe = []
                data_prazo = []
                hora_entrada = []
                hora_saida = []
                preco = []
                sinal = []
                cliente = []
                funcionario = []
                telefone = []
                for row in lista_dados:
                    id_produto.append(row[0])
                    produto.append(row[1])
                    cor.append(row[2])
                    servico.append(row[3])
                    par_pe.append(row[4])
                    data_prazo.append(row[5])
                    hora_entrada.append(row[6])
                    hora_saida.append(row[7])
                    preco.append(row[8])
                    sinal.append(row[9])
                    cliente.append(row[10])
                    funcionario.append(row[11])
                    telefone.append(row[12])

                dict_dados = {
                    "ID": id_produto, "Produto": produto, "Cor": cor,
                    "Serviço": servico, "Par_pe": par_pe, "Data_Prazo": data_prazo,
                    "Hora_entrada":	hora_entrada, "Hora_saida": hora_saida,
                    "Preço": preco, "Sinal": sinal, "Cliente": cliente,
                    "Funcionario": funcionario, "Telefone":	telefone}
            

                dados = pd.DataFrame(data=dict_dados)
                arquivo = arquivo + '\produtos.xls'
                dados.to_excel(arquivo)
                return
            # try:
            #     dados.to_excel('produtos.xls')
            # except:
            #     os.remove('produtos.xls')
            #     dados.to_excel('produtos.xls')
        except TypeError as e:
            pass
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    loja = Interface()
    loja.show()
    app.exec()
