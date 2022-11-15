import os.path
import sys
from os import getenv

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

from crud import CrudLoja
from dialogo import *
from interface import *


class DialogBox(QDialog, Ui_Dialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.check = False

        self.btnImprimirLogar.clicked.connect(self.loginImprimir)

    def loginImprimir(self):
        login = self.inputImprimirLogin.text()
        senha = self.inputImprimirSenha.text()

        logar = CrudLoja(login=login, senha=senha)
        logar.read_funcionario()
        resultado = logar.resultados
        
        if len(resultado) == 1:
            self.login = login
            self.senha = senha
            self.check = True
            self.close()
            return 'Ola'
        self.labelCheck.setStyleSheet('color: rgb(255, 0, 0)')
        self.labelCheck.setText('Login Invalido')

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

        self.btnCliente.clicked.connect(self.mostrar_cliente)
        self.btnProduto.clicked.connect(self.mostrar_produto)

        self.btnPesquisarFinalizar.clicked.connect(self.pesquisa_finalizar)
        self.btnFinalizar.clicked.connect(self.finalizar_produto)

        # user
        self.login = None
        self.senha = None
        
        self.dialogo = DialogBox()
    
    def logar_usuario(self):
        login = self.inputLogin.text()
        senha = self.inputSenha.text()

        logar = CrudLoja(login=login, senha=senha)
        logar.read_funcionario()
        resultado = logar.resultados
        
        if len(resultado) == 1:
            self.login = login
            self.senha = senha
            self.labelNomeFuncionario.setText(self.login)
            self.inputLogin.setText('')
            self.inputSenha.setText('')
            if resultado[0][6] == 'True':
                self.btnPageCadastro.setEnabled(True)
                self.btnPageCadastro.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            return self.PaginaCentral.setCurrentWidget(self.pageHome)
        self.labelLoginMensageBox.setText('Usuario ou Senha invalidos')

    def cadastrar_usuario(self):
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
    
    def logout(self):
        sair = CrudLoja(login=self.login, senha=self.senha)
        sair.update_saida_funcionario()
        self.login = None
        self.senha = None
        self.btnPageCadastro.setEnabled(False)
        self.btnFuncionarioCadastrar.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        return self.PaginaCentral.setCurrentWidget(self.pageLogin)
        
    def checarloginParaImprimir(self):
        self.dialogo.show()

    def cadastrar_cliente(self):
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
    
    def cadastrar_produto(self):
        check = self.dialogo.check
        if not check:
            return
        self.btnCadastrarImprimir.setEnabled(True)
        servico_cliente = self.InputServico.text()
        produto_cliente = self.InputProduto.text()
        telefone_cliente = self.InputTelefoneProduto.text()
        preco_cliente = self.InputPreco.text()
        
        prazo = self.InputPrazo.text()

        radios = [self.radioBranco, self.radioPreto, self.radioAzul, self.radioVermelho, self.radioMarron,
        self.radioRosa, self.radioVerde, self.radioAmarelo, self.radioBege, self.radioLaranja, self.radioRoxo, self.radioBicolor,]
        cor_cliente = []
        for radio in radios:
            if radio.isChecked():
                cor_cliente.append(radio.text())

        preco_lista = preco_cliente.split('/')
        preco = 0
        for p in preco_lista:
            preco += float(p.replace(',','.'))
        cor_cliente = ','.join(cor_cliente)
        cadastrar = CrudLoja(produto=produto_cliente, servico=servico_cliente, cor=cor_cliente, telefone=telefone_cliente, preco=preco, prazo=prazo, produtos='1')
        cadastrar.create_produto()
        cadastrar.mostrar_produtos()

        if len(cadastrar.resultados) == 0:
            self.labelCadastrarProduto.setText(f'Cliente nao existe ou Telefone incorreto')
            return

        resultado = cadastrar.resultados[-1]
        servicos = resultado[3].split('/')
        
        font=QtGui.QFont() 
        font.setPointSize(12)        
        child = QTreeWidgetItem(self.labelCadastrarProduto, [f'{resultado[0]}', f'{resultado[1]}', f'{resultado[2]}', f'{resultado[3]}', f'{resultado[8]}', f'{resultado[5]}', f'{resultado[6]}', f'{resultado[9]}', f'{resultado[4]}'])
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
        pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
        cnv.setFont('Vera', 9)
        # cnv.setFontSize(size='8')
        cnv.scale(0.9, 0.65)
        
        cnv.drawString(0, 1285, f'RÁPIDO DOS CALÇADOS            Cliente')
        cnv.drawString(0, 1270, f'Praça D.Pedro, 31-LOJA 5/Centro')
        cnv.drawString(0, 1255, f'Petrópolis')
        cnv.drawString(0, 1240, f'Telef.: (00) 0000-0000')
        cnv.drawString(0, 1225, f'Horário: de Seg. a Sex. das 8:00 as 18:00')
        cnv.drawString(0, 1210, f'____________________________________________')
        cnv.drawString(0, 1195, f'{resultado[4]} {resultado[7]}')
        cnv.drawString(0, 1180, f'____________________________________________')
        cnv.drawString(0, 1165, f'O.S: {resultado[0]}')
        cnv.drawString(0, 1150, f'{resultado[5]}   Prazo: {resultado[8]}')
        cnv.drawString(0, 1135, f'____________________________________________')
        cnv.drawString(0, 1120, f'Impresso por: {self.login}')
        cnv.drawString(0, 1105, f'____________________________________________')
        cnv.drawString(0, 1090, f'It Objeto    Cor: {resultado[2]}')
        cnv.drawString(0, 1075, f'Serviço: {resultado[3]}')
        cnv.drawString(0, 1060, f'____________________________________________')
        cnv.drawString(0, 1045, f'{resultado[1]} {resultado[2]}')
        contador = 0
        y = 1030
        for servico in servicos:
            if len(preco_lista) > 1:
                cnv.drawString(0, y, f'                    {servico}           R${preco_lista[contador]}')
            else:
                cnv.drawString(0, y, f'                    {servico}           R${preco_cliente}')
            y -= 15
            contador += 1
        cnv.drawString(0, y - 15, f' ___________________________________________')
        cnv.drawString(0, y - 30, f'                     SubTotal: R$ {preco}')
        cnv.drawString(0, y - 45, f'                     Desconto: R$ 00.00')
        cnv.drawString(0, y - 60, f'                     Sinal:    R$ 00.00')
        cnv.drawString(0, y - 75, f'                     Total:    R$ {preco}')
        cnv.drawString(0, y - 90, f' ___________________________________________')
        cnv.drawString(0, y - 105, f' Prezado cliente! Preserve este Documento')
        cnv.drawString(0, y - 120, f' pois será através dele, que nossos')
        cnv.drawString(0, y - 135, f' funcionarios os indentificarão os objetos')
        cnv.drawString(0, y - 150, f' aqui deixados.')
        cnv.drawString(0, y - 165, f' ___________________________________________')
        cnv.drawString(0, y - 180, f'     NÃO NOS RESPONSABILIZAMOS POR')
        cnv.drawString(0, y - 195, f'     OBJETOS DEIXADOS POR MAIS')
        cnv.drawString(0, y - 210, f'         DE 30 DIAS.')
        cnv.drawString(0, y - 225, f' DATA: __/__/____     Ass.:__________________')
        cnv.setPageSize((300, 840))
        cnv.save()

        cnvloja = canvas.Canvas(notaLoja)
        pdfmetrics.registerFont(TTFont('Vera', 'Vera.ttf'))
        cnvloja.setFont('Vera', 9)
        # cnv.setFontSize(size='8')
        cnvloja.scale(0.9, 0.65)
        
        cnvloja.drawString(0, 1285, f'RÁPIDO DOS CALÇADOS            Balcão')
        cnvloja.drawString(0, 1270, f'Praça D.Pedro, 31-LOJA 5/Centro')
        cnvloja.drawString(0, 1255, f'Petrópolis')
        cnvloja.drawString(0, 1240, f'Telef.: (00) 0000-0000')
        cnvloja.drawString(0, 1225, f'Horário: de Seg. a Sex. das 8:00 as 18:00')
        cnvloja.drawString(0, 1210, f'____________________________________________')
        cnvloja.drawString(0, 1195, f'{resultado[4]} {resultado[7]}')
        cnvloja.drawString(0, 1180, f'____________________________________________')
        cnvloja.drawString(0, 1165, f'O.S: {resultado[0]}')
        cnvloja.drawString(0, 1150, f'{resultado[5]}   Prazo: {resultado[8]}')
        cnvloja.drawString(0, 1135, f'____________________________________________')
        cnvloja.drawString(0, 1120, f'Impresso por: {self.login}')
        cnvloja.drawString(0, 1105, f'____________________________________________')
        cnvloja.drawString(0, 1090, f'It Objeto    Cor: {resultado[2]}')
        cnvloja.drawString(0, 1075, f'Serviço: {resultado[3]}')
        cnvloja.drawString(0, 1060, f'____________________________________________')
        cnvloja.drawString(0, 1045, f'{resultado[1]} {resultado[2]}')
        contador = 0
        y = 1030
        for servico in servicos:
            if len(preco_lista) > 1:
                cnvloja.drawString(0, y, f'                    {servico}           R${preco_lista[contador]}')
            else:
                cnvloja.drawString(0, y, f'                    {servico}           R${preco_cliente}')
            y -= 15
            contador += 1
        cnvloja.drawString(0, y - 15, f' ___________________________________________')
        cnvloja.drawString(0, y - 30, f'                     SubTotal: R$ {preco}')
        cnvloja.drawString(0, y - 45, f'                     Desconto: R$ 00.00')
        cnvloja.drawString(0, y - 60, f'                     Sinal:    R$ 00.00')
        cnvloja.drawString(0, y - 75, f'                     Total:    R$ {preco}')
        cnvloja.drawString(0, y - 90, f' ___________________________________________')
        cnvloja.drawString(0, y - 105, f' Prezado cliente! Preserve este Documento')
        cnvloja.drawString(0, y - 120, f' pois será através dele, que nossos')
        cnvloja.drawString(0, y - 135, f' funcionarios os indentificarão os objetos')
        cnvloja.drawString(0, y - 150, f' aqui deixados.')
        cnvloja.drawString(0, y - 165, f' ___________________________________________')
        cnvloja.drawString(0, y - 180, f'     NÃO NOS RESPONSABILIZAMOS POR')
        cnvloja.drawString(0, y - 195, f'     OBJETOS DEIXADOS POR MAIS')
        cnvloja.drawString(0, y - 210, f'         DE 30 DIAS.')
        cnvloja.drawString(0, y - 225, f' DATA: __/__/____     Ass.:__________________')
        cnvloja.setPageSize((300, 840))
        cnvloja.save()
        impressora = win32print.EnumPrinters(2)[1]
        win32print.SetDefaultPrinter(impressora[2])
        win32api.ShellExecute(0, "print", notaCliente, None, diretorio[:-16], 0)

        impressora = win32print.EnumPrinters(2)[1]
        win32print.SetDefaultPrinter(impressora[2])
        win32api.ShellExecute(0, "print", notaLoja, None, notaLoja[:-13], 0)


    def mostrar_todos_produtos(self):
        self.labelClientesProdutos.clear()
        self.labelClientesProdutos.headerItem().setText(0, "O.S")
        self.labelClientesProdutos.headerItem().setText(1, "Produto")
        self.labelClientesProdutos.headerItem().setText(2, "Cor")
        self.labelClientesProdutos.headerItem().setText(3, "Serviço")
        self.labelClientesProdutos.headerItem().setText(4, "Data Prazo")
        self.labelClientesProdutos.headerItem().setText(5, "Hora Entrada")
        self.labelClientesProdutos.headerItem().setText(6, "Hora Saida")
        self.labelClientesProdutos.headerItem().setText(7, "Preço")
        self.labelClientesProdutos.headerItem().setText(8, "Nome")
        produtos = CrudLoja(produtos='todos')
        produtos.mostrar_produtos()
        resultados = produtos.resultados

        arquivo = fr'C:\Users\{self.username}\Desktop\relatorio_Produtos.txt'
        if os.path.isfile(arquivo):
            pass
        else:
            arquivo = fr'C:\Users\{self.username}\OneDrive\Área de Trabalho\relatorio_Produtos.txt'
        font=QtGui.QFont() 
        font.setPointSize(12) 
        for resultado in resultados:
            child = QTreeWidgetItem(self.labelClientesProdutos, [f'{resultado[0]}', f'{resultado[1]}', f'{resultado[2]}', f'{resultado[3]}', f'{resultado[8]}', f'{resultado[5]}', f'{resultado[6]}', f'{resultado[9]}', f'{resultado[4]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)
        
        

        # with open(arquivo, 'r') as relatorio:
        #     self.labelClientesProdutos.setText(f'{relatorio.read()}\n'
        #                                      f'\n'
        #                                      f'\n'
        #                                      f'Pré vizualização(O arquivo foi enviado para sua aréa de trabalho '
        #                                      f'relatorio_Produtos.txt)')
    def mostrar_todos_clientes(self):
        self.labelClientesProdutos.clear()
        self.labelClientesProdutos.setColumnCount(3)
        self.labelClientesProdutos.headerItem().setText(0, "ID")
        self.labelClientesProdutos.headerItem().setText(1, "Nome")
        self.labelClientesProdutos.headerItem().setText(2, "Telefone")
        
        clientes = CrudLoja(clientes='todos')
        clientes.mostrar_clientes()
        resultados = clientes.resultados

        arquivo = fr'C:\Users\{self.username}\Desktop\relatorio_Clientes.txt'
        if os.path.isfile(arquivo):
            pass
        else:
            arquivo = fr'C:\Users\{self.username}\OneDrive\Área de Trabalho\relatorio_Clientes.txt'

        font=QtGui.QFont() 
        font.setPointSize(12) 
        for resultado in resultados:
            child = QTreeWidgetItem(self.labelClientesProdutos, [f'{resultado[2]}', f'{resultado[0]}', f'{resultado[1]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)

        # with open(arquivo, 'r') as relatorio:
        #     self.labelClientesProdutos.setText(f'{relatorio.read()}\n'
        #                                      f'\n'
        #                                      f'\n'
        #                                      f'Pré vizualização(O arquivo foi enviado para sua aréa de trabalho '
        #                                      f'relatorio_Clientes.txt)')

    def mostrar_produto(self):
        self.labelClienteProduto.clear()
        self.labelClienteProduto.headerItem().setText(0, "O.S")
        self.labelClienteProduto.headerItem().setText(1, "Produto")
        self.labelClienteProduto.headerItem().setText(2, "Cor")
        self.labelClienteProduto.headerItem().setText(3, "Serviço")
        self.labelClienteProduto.headerItem().setText(4, "Data Prazo")
        self.labelClienteProduto.headerItem().setText(5, "Hora Entrada")
        self.labelClienteProduto.headerItem().setText(6, "Hora Saida")
        self.labelClienteProduto.headerItem().setText(7, "Preço")
        self.labelClienteProduto.headerItem().setText(8, "Nome")

        telefone_cliente = self.InputTelefoneClienteProduto.text()
        produtos = CrudLoja(telefone=telefone_cliente, produtos='1')
        produtos.mostrar_produtos()
        resultados = produtos.resultados
        if len(resultados) > 0:
            font=QtGui.QFont() 
            font.setPointSize(12) 
            for resultado in resultados:
                child = QTreeWidgetItem(self.labelClienteProduto, [f'{resultado[0]}', f'{resultado[1]}', f'{resultado[2]}', f'{resultado[3]}', f'{resultado[8]}', f'{resultado[5]}', f'{resultado[6]}', f'{resultado[9]}', f'{resultado[4]}'])
                for i, _ in enumerate(resultado):
                    child.setFont(i , font)
        # else:
            # self.labelClienteProduto.setText('Nada encontrado')
    def mostrar_cliente(self):
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
    
    def pesquisa_finalizar(self):
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
            child = QTreeWidgetItem(self.LabelFinalizar, [f'{resultado[0]}', f'{resultado[1]}', f'{resultado[2]}', f'{resultado[3]}', f'{resultado[8]}', f'{resultado[5]}', f'{resultado[6]}', f'{resultado[9]}', f'{resultado[4]}'])
            for i, _ in enumerate(resultado):
                child.setFont(i , font)
            return
        child = QTreeWidgetItem(self.LabelFinalizar, ['O.S Incorreto ou nao existe'])
        return
        

    def finalizar_produto(self):
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
        child = QTreeWidgetItem(self.LabelFinalizar, [f'{resultado[0]}', f'{resultado[1]}', f'{resultado[2]}', f'{resultado[3]}', f'{resultado[8]}', f'{resultado[5]}', f'{resultado[6]}', f'{resultado[9]}', f'{resultado[4]}'])
        for i, _ in enumerate(resultado):
            child.setFont(i , font)
        child = QTreeWidgetItem(self.LabelFinalizar, ['Finalizado!!'])
        return 
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    estacionamento = Interface()
    estacionamento.show()
    app.exec()
