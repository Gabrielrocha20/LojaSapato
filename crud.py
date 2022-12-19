import os.path
import sqlite3
import sys
from datetime import datetime
from hashlib import md5
from os import getenv
from random import choice, randint
from unittest import result

from faker import Faker


class CrudLoja:
    def __init__(self, nome=None, telefone=None, servico=None, cor=None,
                 produto=None, preco=None, prazo=None, par_pe=None, filtro=None,
                 produtos=None, sinal=None, clientes=None, o_s=None, status=None, funcionario=None,
                 login=None, senha=None, adm=None):

        self.produto = produto
        self.prazo = prazo
        self.nome = nome
        self.par_pe = par_pe
        self.telefone = telefone
        self.cor = cor
        self.servico = servico
        self.preco = preco
        self.produtos = produtos
        self.clientes = clientes
        self.o_s = o_s
        self.sinal = sinal
        self.funcionario = funcionario
        self.adm = adm
        self.status = status

        # filtros
        self.filtro = filtro

        self.login = login
        self.senha = senha

        self.check_os = True

        self.cliente = ''
        self.resultados = None

        self.username = getenv("USERNAME")

        self.con = sqlite3.connect('dados.db')
        self.cursor = self.con.cursor()
        try:
            sql_Cliente = 'SELECT * FROM Cliente'
            self.cursor.execute(sql_Cliente)
        except sqlite3.OperationalError as e:
            self.cursor.execute(
                """
                CREATE TABLE "Cliente" (
                "Nome"	TEXT,
                "Telefone"	TEXT,
                "ID"	INTEGER UNIQUE,
                PRIMARY KEY("ID" AUTOINCREMENT)
                );
                """)

        try:
            sql_Produto = 'SELECT * FROM Produto'
            self.cursor.execute(sql_Produto)
        except sqlite3.OperationalError as e:
            self.cursor.execute(
                """
                CREATE TABLE "Produto" (
                "ID"	INTEGER UNIQUE,
                "Produto"	TEXT,
                "Cor"	TEXT,
                "Serviço"	TEXT,
                "Par_pe"	TEXT,
                "Data_Prazo"	TEXT,
                "Hora_entrada"	TEXT,
                "Hora_saida"	TEXT,
                "Preço"	TEXT,
                "Sinal"	TEXT,
                "Cliente"	TEXT,
                "Funcionario"	TEXT,
                "Telefone"	TEXT,
                "Status"	TEXT,
                PRIMARY KEY("ID" AUTOINCREMENT)
            )
                """
            )
        try:
            sql_Produto = 'SELECT * FROM Funcionario'
            self.cursor.execute(sql_Produto)
        except sqlite3.OperationalError as e:
            self.cursor.execute(
                """
                CREATE TABLE "Funcionario" (
                "ID"	INTEGER UNIQUE,
                "Login"	TEXT,
                "Senha"	TEXT,
                "Telefone"	TEXT,
                "Hora_entrada"	TEXT,
                "Hora_saida"	INTEGER,
                PRIMARY KEY("ID" AUTOINCREMENT)
            )
                """
            )
        try:
            sql_Produto = 'SELECT * FROM Horarios_acesso'
            self.cursor.execute(sql_Produto)
        except sqlite3.OperationalError as e:
            self.cursor.execute(
                """
                CREATE TABLE "Horarios_acesso" (
                    "ID"	INTEGER UNIQUE,
                    "ID_Funcionario"	INTEGER,
                    "Hora_acesso"	TEXT,
                    PRIMARY KEY("ID" AUTOINCREMENT)
                )
                """
            )
        try:
            sql_Produto = 'SELECT * FROM Contas'
            self.cursor.execute(sql_Produto)
        except sqlite3.OperationalError as e:
            self.cursor.execute(
                """
                CREATE TABLE "Contas" (
                "ID"	INTEGER,
                "Login"	TEXT,
                "Senha"	TEXT,
                "Administrador"	TEXT,
                PRIMARY KEY("ID" AUTOINCREMENT)
            )
                """
            )

    def cadastrar_conta(self):
        cursor = self.con.cursor()

        senha = self.senha.encode("utf8")
        senha_hash = md5(senha).hexdigest()

        checar_funcionario = f'SELECT * FROM Contas WHERE Login = "{self.login}"'
        cursor.execute(checar_funcionario)
        checar_funcionario = cursor.fetchall()

        if len(checar_funcionario) == 0:
            cadastro = f'INSERT INTO Contas (Login, Senha, Administrador) VALUES ("{self.login}", "{senha_hash}", "{self.adm}")'
            cursor.execute(cadastro)
            self.cliente = True
        else:
            self.cliente = False

        self.con.commit()

    def cadastrar_funcionario(self):
        cursor = self.con.cursor()

        data = datetime.now()
        data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")
        senha = self.senha.encode("utf8")
        senha_hash = md5(senha).hexdigest()

        checar_funcionario = f'SELECT * FROM Funcionario WHERE Login = "{self.login}"'
        cursor.execute(checar_funcionario)
        checar_funcionario = cursor.fetchall()

        if len(checar_funcionario) == 0:
            cadastro = f'INSERT INTO Funcionario (Login, Senha, Telefone, Hora_entrada, Hora_saida) VALUES ("{self.login}", "{senha_hash}", "{self.telefone}", "{data_formatada}", "0")'
            cursor.execute(cadastro)
            self.cliente = 'Funcionario Cadastrado'
        else:
            self.cliente = 'Funcionario Ja cadastrado'

        checar_funcionario = f'SELECT * FROM Funcionario WHERE Telefone = "{self.telefone}"'
        cursor.execute(checar_funcionario)
        checar_funcionario = cursor.fetchall()
        if len(checar_funcionario) == 0:
            id_funcionario = checar_funcionario[0][0]
            cadastro = f'INSERT INTO Horarios_acesso (ID_Funcionario, Hora_acesso) VALUES ("{id_funcionario}", "{data_formatada}")'
            cursor.execute(cadastro)
        self.con.commit()

    def read_funcionario(self):
        if self.funcionario:
            senha = self.senha.encode("utf8")
            senha_hash = md5(senha).hexdigest()
            cursor = self.con.cursor()
            checar_funcionario = f'SELECT * FROM Funcionario WHERE Login = "{self.login}" AND Senha = "{senha_hash}"'
            cursor.execute(checar_funcionario)
            checar_funcionario = cursor.fetchall()
            if len(checar_funcionario) == 0:
                self.resultados = []
                return
            resultados = []
            for i in checar_funcionario:
                resultados.append(i)
            self.resultados = resultados

            identificador = checar_funcionario[0][0]
            data = datetime.now()
            data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")

            sql = f'UPDATE Funcionario SET Hora_entrada = "{data_formatada}" WHERE ID = {identificador}'
            up_hora = f'INSERT INTO Horarios_acesso (ID_Funcionario, Hora_acesso) VALUES ("{identificador}", "{data_formatada}")'
            cursor.execute(up_hora)
            self.con.commit()
            return
        else:
            senha = self.senha.encode("utf8")
            senha_hash = md5(senha).hexdigest()
            cursor = self.con.cursor()
            checar_funcionario = f'SELECT * FROM Contas WHERE Login = "{self.login}" AND Senha = "{senha_hash}"'
            cursor.execute(checar_funcionario)
            checar_funcionario = cursor.fetchall()
            if len(checar_funcionario) == 0:
                self.resultados = []
                return
            resultados = []
            for i in checar_funcionario:
                resultados.append(i)
            self.resultados = resultados
            self.con.commit()
            return

    def update_saida_funcionario(self):
        cursor = self.con.cursor()
        o_s = self.o_s
        checar_funcionario = f'SELECT * FROM Funcionario WHERE Login = "{self.login}" AND Senha = "{self.senha}"'
        cursor.execute(checar_funcionario)
        checar_funcionario = cursor.fetchall()
        if len(checar_funcionario) < 1:
            self.check_os = False
            return
        identificador = checar_funcionario[0][0]
        data = datetime.now()
        data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")
        sql = f'UPDATE Funcionario SET Hora_saida = "{data_formatada}" WHERE ID = {identificador}'
        cursor.execute(sql)
        self.con.commit()

    def mostrar_clientes(self):
        cursor = self.con.cursor()
        if self.clientes == '1':
            sql = f'SELECT * FROM Cliente WHERE Telefone = "{self.telefone}"'
            cursor.execute(sql)
            results = cursor.fetchall()
            self.resultados = results
            return
        sql = 'SELECT * FROM Cliente'
        cursor.execute(sql)
        results = cursor.fetchall()
        self.resultados = results

    def mostrar_produtos(self):
        cursor = self.con.cursor()
        if self.produtos == '1':
            sql = f'SELECT * FROM Produto WHERE Telefone = "{self.telefone}"'
            cursor.execute(sql)
            results = cursor.fetchall()
            self.resultados = results
            return
        elif self.produtos == 'os':
            sql = f'SELECT * FROM Produto WHERE ID = {self.o_s}'
            cursor.execute(sql)
            results = cursor.fetchall()
            self.resultados = results
            return
        elif self.produtos == 'filtro':
            sql = f"""SELECT * FROM Produto WHERE
            ID LIKE '{self.filtro}' OR
            Produto LIKE '{self.filtro}' OR
            Cor LIKE '{self.filtro}' OR
            Cliente LIKE '%{self.filtro}%' OR
            Serviço LIKE '{self.filtro}' OR
            Hora_entrada LIKE '{self.filtro}' OR
            Hora_saida LIKE '{self.filtro}' OR
            Data_Prazo LIKE '{self.filtro}' OR
            Par_pe LIKE '{self.filtro}'OR
            Status LIKE '{self.filtro}
            """
            cursor.execute(sql)
            results = cursor.fetchall()
            self.resultados = results
            return
        else:
            sql = 'SELECT * FROM Produto'
            cursor.execute(sql)
            results = cursor.fetchall()
            self.resultados = results
            return

    def create_cliente(self):
        cursor = self.con.cursor()

        nome = self.nome
        telefone = self.telefone

        checar_cliente = f'SELECT * FROM Cliente WHERE telefone = "{telefone}"'
        cursor.execute(checar_cliente)
        checar_cliente = cursor.fetchall()

        if (len(nome) == 0) or (len(telefone) == 0):
            return
        elif len(checar_cliente) == 0:
            cadastro = f'INSERT INTO Cliente (Nome, Telefone) VALUES ("{nome}", "{telefone}")'
            cursor.execute(cadastro)
            self.cliente = 'Cliente Cadastrado'
        else:
            self.cliente = 'Cliente Ja cadastrado'

        self.con.commit()

    def create_produto(self):
        cursor = self.con.cursor()

        produto = self.produto
        prazo = self.prazo
        cor = self.cor
        telefone = self.telefone
        servico = self.servico
        preco = self.preco
        sinal = self.sinal
        funcionario = self.funcionario
        par_pe = self.par_pe
        status = self.status

        data = datetime.now()
        data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")

        checar_cliente = f'SELECT * FROM Cliente WHERE Telefone = "{telefone}"'
        cursor.execute(checar_cliente)
        checar_cliente = cursor.fetchall()

        for cliente in checar_cliente:
            nome = cliente[0]

        checar_produto = f'SELECT * FROM Produto WHERE Telefone = "{telefone}"'
        cursor.execute(checar_produto)
        checar_produto = cursor.fetchall()

        if (len(produto) == 0) or (len(telefone) == 0) or (len(cor) == 0) or (len(servico) == 0) or (len(prazo) == 0) or (len(par_pe) == 0) or (len(status)) == 0:
            return
        elif len(checar_cliente) == 0:
            self.cliente = 'Cliente Não existe ou Telefone invalido'

        else:
            cadastro = f'INSERT INTO Produto (Produto, Cor, Serviço, Cliente, Hora_entrada, Hora_saida, Telefone, Data_Prazo, Preço, Funcionario, Sinal, Par_pe, Status) VALUES ("{produto}",\
            "{cor}", "{servico}","{nome}", "{data_formatada}", "{0}", "{telefone}", "{prazo}", "{preco}", "{funcionario}", "{sinal}", "{par_pe}", "{status}")'
            cursor.execute(cadastro)
            self.cliente = 'Produto registrado'

        self.con.commit()

    def update(self):
        cursor = self.con.cursor()
        o_s = self.o_s
        checar_cliente = f'SELECT * FROM Produto WHERE ID = {o_s}'
        cursor.execute(checar_cliente)
        checar_cliente = cursor.fetchall()
        if len(checar_cliente) < 1:
            self.check_os = False
            return
        data = datetime.now()
        data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")
        sql = f'UPDATE Produto SET Hora_saida = "{data_formatada}" WHERE ID = {o_s} '
        cursor.execute(sql)
        self.con.commit()

    def update_status(self):
        cursor = self.con.cursor()
        o_s = self.o_s
        status = self.status
        checar_cliente = f'SELECT * FROM Produto WHERE ID = {o_s}'
        cursor.execute(checar_cliente)
        checar_cliente = cursor.fetchall()
        if len(checar_cliente) < 1:
            self.check_os = False
            return

        sql = f'UPDATE Produto SET Status = "{status}" WHERE ID = {o_s} '
        cursor.execute(sql)
        self.con.commit()

# if __name__ == '__main__':
#     for i in range(1000):
#         fake = Faker()
#         nome = f"{fake.name()} {i}"
#         telefone = randint(10000000000, 99999999999)
#         data = datetime.now()
#         data_formatada = datetime.strftime(data, "%d/%m/%Y %H:%M")

#         produto = 'Tenis'
#         preco = i + 20
#         sinal = preco / 2
#         cor = 'preto'
#         servico = 'Colar'
#         prazo = '22/12/2022 09:00'
#         funcionario = 'Gabriel Rocha'

#         crud = CrudLoja(nome=nome, telefone=str(telefone), servico=servico, cor=cor,
#                 produto=produto, preco=preco, prazo=prazo, par_pe='Par', filtro=None,
#                 produtos=None, sinal=sinal, clientes=None, o_s=None, funcionario=funcionario,
#                 login=None, senha=None, adm=None)
#         crud.create_cliente()
#         crud.create_produto()
