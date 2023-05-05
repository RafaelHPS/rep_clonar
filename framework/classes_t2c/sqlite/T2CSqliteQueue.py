import sqlite3
import datetime
from pathlib import Path
from framework.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel

ROOT_DIR = Path(__file__).parent.parent.parent.__str__()

"""
ESTRUTURA ESPERADA POR ESSA CLASSE:

Create table tbl_Fila_Processamento(
    id integer primary key,
    referencia varchar(200),
    datahora_criado varchar(50),
    ultima_atualizacao varchar(50),
    nome_maquina varchar(200),
    status varchar(100),
    obs varchar(500));

COLUNAS PODEM SER ADICIONADAS, MAS NUNCA EXCLUÍDAS OU MOVIDAS
O ARQUIVO DO BANCO LOCALIZADO EM resources/sqlite/banco_dados.db JÁ POSSUI ESSA TABELA CRIADA E VAZIA
"""

#Classe responsável para manipulação do sqlite e controle de fila
class T2CSqliteQueue:
    #Cria a conexão com o banco
    #Nome da máquina precisa ser algum identificador único por execução
    #CaminhoBD e TabelaFila não precisam ser informados por padrão
    def __init__(self, arg_clssMaestro:T2CMaestro, arg_strCaminhoBd:str=None, arg_strTabelaFila:str=None, arg_strNomeMaquina:str=None):
        self.var_clssMaestro = arg_clssMaestro
        #Se o caminho não for especificado, usa a configuração padrão
        self.var_strTabelaFila = "tbl_Fila_Processamento" if(arg_strTabelaFila is None) else arg_strTabelaFila
        self.var_strPathToDb = ROOT_DIR + "\\resources\\sqlite\\banco_dados.db" if(arg_strCaminhoBd is None) else arg_strCaminhoBd
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE status = 'NEW'")

        self.var_intItemsQueue = len(var_csrCursor.fetchall())
        self.var_strNomeMaquina = arg_strNomeMaquina
        var_csrCursor.close()

    #Atualiza a própria classe, usado em vários pontos do projeto para atualziar os itens na fila
    def update(self):
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE status = 'NEW'")
        self.var_intItemsQueue = len(var_csrCursor.fetchall())
        var_csrCursor.close()

    #Insere um item na tabela especificada, com a referência e com os valores adicionais
    #IMPORTANTE: Criar colunas extras para informações a mais informadas em arg_listInfAdicional, não cria colunas sozinho
    def insert_new_queue_item(self, arg_strReferencia:str, arg_listInfAdicional:list = None):
        #Aqui eu insiro com o nome da máquina vazio, para que qualquer máquina possa processar em paralelo mais a frente
        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        var_listValues = [arg_strReferencia, var_strNow, var_strNow, "", "NEW", ""]
        var_listColumns = []
        
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE id = 0")
        for description in var_csrCursor.description:
            if(description[0] != "id"): var_listColumns.append(description[0])

        if(arg_listInfAdicional is not None): var_listValues.extend(arg_listInfAdicional)

        self.update()

        #Construindo o comando insert
        var_strInsert = "INSERT INTO " + self.var_strTabelaFila + " (" + var_listColumns.__str__() + ") VALUES (" + var_listValues.__str__() + ")"
        var_strInsert = var_strInsert.replace("[", "").replace("]", "")
        
        #Executando o comando insert
        try:
            var_csrCursor.execute(var_strInsert)
            var_csrCursor.connection.commit()
        except Exception:
            self.var_clssMaestro.write_log(arg_strMensagemLog="Erro ao inserir linhas: " + str(Exception), arg_enumLogLevel=LogLevel.ERROR)
            raise

        var_csrCursor.close()
        self.update()

    #Retorna um item específico
    def get_specific_queue_item(self, arg_intIndex:int) -> tuple:
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE id = " + arg_intIndex.__str__())
        var_tplRow = var_csrCursor.fetchone()
        var_csrCursor.close()

        self.update()
        return var_tplRow
   
    #Retorna o próximo item não processado e com nenhuma máquina alocada, retorna None se não existe nenhum item assim
    def get_next_queue_item(self) -> tuple:
        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET ultima_atualizacao = '" + var_strNow + "', nome_maquina = '" + self.var_strNomeMaquina + "', status = 'ON QUEUE' WHERE id = (SELECT MIN(id) FROM " + self.var_strTabelaFila + " WHERE status = 'NEW')").connection.commit()
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("SELECT * FROM " + self.var_strTabelaFila + " WHERE nome_maquina = '" + self.var_strNomeMaquina + "' and status = 'ON QUEUE' ORDER BY id LIMIT 1")
        var_tplRow:tuple = var_csrCursor.fetchone()
        if(var_tplRow is not None): var_csrCursor.execute("UPDATE " + self.var_strTabelaFila + " SET status = 'RUNNING' WHERE id = " + var_tplRow[0].__str__()).connection.commit()

        var_csrCursor.close()
        self.update()

        return var_tplRow

    #Atualiza os dados de um item com um determinado index
    def update_status_item(self, arg_intIndex:int, arg_strNovoStatus:str, arg_strObs:str=""):
        #Tratando os casos onde obs vem com quotes, trocando por *
        arg_strObs = arg_strObs.replace('"', '*').replace("'", '*')

        var_strNow = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET ultima_atualizacao = '" + var_strNow + "', status = '" + arg_strNovoStatus + "', obs = '" + arg_strObs + "' WHERE id = " + arg_intIndex.__str__())
        var_csrCursor.connection.commit()
   
        var_csrCursor.close()
        self.update()
   
    #Marca todos os itens com status NEW como ABANDONED
    def abandon_queue(self):
        var_csrCursor = sqlite3.connect(self.var_strPathToDb).execute("UPDATE " + self.var_strTabelaFila + " SET status = 'ABANDONED' WHERE status = 'NEW'")
        var_csrCursor.connection.commit()

        var_csrCursor.close()
        self.update()
   