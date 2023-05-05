from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
from framework.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel
from framework.classes_t2c.utils.T2CExceptions import BusinessRuleException

#Classe responsável pelo processamento principal, necessário preencher com o seu código no método execute
class T2CProcess:
    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None):
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            
    #Parte principal do código, deve ser preenchida pelo desenvolvedor
    #Acesse o item a ser processado pelo self.queue_item
    def execute(self, arg_tplQueueItem:tuple):
        #Implemente aqui seu código para execução, o print() é apenas placeholder
        print()
        