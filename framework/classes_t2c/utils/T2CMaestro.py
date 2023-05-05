from botcity.maestro import *
from botcity.maestro.model import BotExecution, AutomationTask
import random, datetime, socket
from enum import Enum

class LogLevel(Enum):
    """
    Classe enum, usada para colocar o nível de um novo log.
    """

    INFO = "INFO"
    WARN = "WARN"
    ERROR = "ERROR"
    FATAL = "FATAL"


class T2CMaestro:
    """
    Classe responsável por todas as interações com o Maestro durante a execução.
    """

    def __init__(self, arg_clssExecution:BotExecution, arg_dictConfig:dict):
        """
        Inicia uma classe T2CMaestro, estabelecendo uma conexão com o Maestro, criando logs para o projeto e armazenando dados referente a execução em variáveis dentro da classe.

        Argumentos:
             - arg_clssExecution (BotExecution): Execução do projeto, nulo se estiver rodando localmente.
             - arg_dictConfig (dict): Dicionário de configuração do framework.
        """
        self.var_dictConfig = arg_dictConfig

        var_clssMaestroInit = BotMaestroSDK()
        var_clssMaestroInit.login(login=self.var_dictConfig["MaestroLogin"], 
                                  key=self.var_dictConfig["MaestroKey"], 
                                  server=self.var_dictConfig["MaestroServer"])
        self.var_clssMaestro = var_clssMaestroInit
        self.var_strNomeProcesso = arg_dictConfig["NomeProcesso"]

        try:
            #Se der certo, com tudo fornecido, significa que está rodando pelo maestro, atribuindo task e activity id
            self.var_clssTask = self.var_clssMaestro.get_task(arg_clssExecution.task_id)
            self.var_intActivityId = self.var_clssTask.activity_id
            self.var_strRunnerId = self.var_clssTask.machine_id
            self.var_boolIsRunningFromTask = True

        except Exception as exception:
            #No caso de não conseguir atribuir a task e activity id, gera uma task com none e um activity id randômico
            random.seed(datetime.datetime.now().microsecond)

            self.var_clssTask = None
            self.var_intActivityId = round(random.random()*1000000)
            self.var_strRunnerId = socket.gethostname()
            self.var_boolIsRunningFromTask = False

        #Criando o log do processo
        try:
            #Essa linha gera um erro em caso do log não existir
            self.var_clssMaestro.get_log(self.var_strNomeProcesso)
        except:
            #No caso de erro (log não existe), cria um novo log do zero
            var_listColumns = [Column(name="Reference", label="Reference", width=200), Column(name="LogLevel", label="LogLevel", width=25), Column(name="Message", label="Message", width=2000)]
            self.var_clssMaestro.new_log(activity_label=self.var_strNomeProcesso, columns=var_listColumns)


    def finish_task(self, arg_boolSucesso:bool, arg_strMensagem:str):
        """
        Finaliza uma task no Maestro.
        Se não estiver conectado com o Maestro, o método não faz nada.

        Argumentos:
            - arg_boolSucesso (bool): Indica se a task deve ser finalizada como sucesso ou erro.
            - arg_strMensagem (str): Mensagem a ser adicionada no final da execução.
        """
        
        if(self.var_boolIsRunningFromTask):
            self.var_clssMaestro.finish_task(
                task_id=self.var_clssTask.id,
                status=AutomationTaskFinishStatus.SUCCESS if arg_boolSucesso else AutomationTaskFinishStatus.FAILED,
                message=arg_strMensagem
            )


    def write_log(self, arg_strMensagemLog:str, arg_strReferencia:str="-", arg_enumLogLevel:LogLevel=LogLevel.INFO):
        """
        Realiza um print e grava uma nova entrada nos logs do projeto no Maestro com os dados fornecidos em argumentos.

        Argumentos:
            - arg_strMensagemLog (str): A mensagem principal a ser gravada no log.
            - arg_strReferencia (str): Referência do item da fila. `(DEFAULT = "-")`
            - arg_enumLogLevel (LogLevel): Enum LogLevel usado para indicar a seriedade log. `(DEFAULT = LogLevel.INFO)`
        """

        if (arg_strReferencia != "-"): 
            print(arg_enumLogLevel.value + " - " + arg_strReferencia + " - " + arg_strMensagemLog)
        else:
            print(arg_enumLogLevel.value + " - " + arg_strMensagemLog)

        #Criando e enviando log para o maestro
        var_jsonValues = {
            "Reference": arg_strReferencia,
            "LogLevel": arg_enumLogLevel.value,
            "Message": arg_strMensagemLog
        }
        self.var_clssMaestro.new_log_entry(activity_label=self.var_strNomeProcesso, values=var_jsonValues)


    def init_new_task(self, arg_strLabelTask:str, arg_boolEhTeste:bool=False, arg_dictParametros:dict=None):
        """
        Inicia uma nova task no Maestro a partir da sua label.
        Não inicia uma nova task se o projeto não estiver rodando pelo Maestro.

        Argumentos:
            - arg_strLabelTask (str): Label do projeto, usada para iniciar uma nova task.
            - arg_boolEhteste (bool): Indica se a task deve ser iniciado como teste. `(DEFAULT = False)`
            - arg_dictParametros (dict): Argumentos para a nova task, organizados em um dicionário. `(DEFAULT = None)`
        """
        
        if(self.var_boolIsRunningFromTask):
            self.var_clssMaestro.create_task(activity_label=arg_strLabelTask, test=arg_boolEhTeste, parameters=arg_dictParametros)


    #Pega uma credencial especificada do maestro a partir de uma label e uma key, volta nulo se der erro
    def get_credential(self, arg_strKey:str) -> str:
        """
        Retorna uma credencial guardada no Maestro para uso durante a execução.
        Label do Maestro deve ser o nome do projeto. No caso de labels diferentes, alterar o código.
        
        Argumentos:
            - arg_strKey (str): Key de uma credencial armazenada no maestro.
        """

        try:
            var_strCredential = self.var_clssMaestro.get_credential(label=self.var_dictConfig["NomeProcesso"], key=arg_strKey)
            return var_strCredential
        except Exception as exception:
            self.write_log(arg_strMensagemLog="Não foi possível pegar a credencial com a key " + arg_strKey + " (" + str(exception) + ")")
            return None