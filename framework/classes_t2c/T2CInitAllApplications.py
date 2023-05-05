from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
from framework.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel
from framework.classes_t2c.utils.T2CExceptions import BusinessRuleException
from framework.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue

#Classe feita para ser invocada principalmente no começo de um processo, para iniciar os processos necessários para a automação
class T2CInitAllApplications:
    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None, arg_clssSqliteQueue:T2CSqliteQueue=None):
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro
            self.var_clssSqliteQueue = arg_clssSqliteQueue

    #Subindo itens para a fila no começo do processo, apenas se necessário
    #Código placeholder
    #Se o seu projeto precisa de mais do que um método simples para subir a sua fila, considere fazer um projeto dispatcher
    def add_to_queue(self, arg_clssSqliteQueue):
        True


    #Método que roda para iniciar os programas necessários, apenas com a estrutura em código
    def execute(self, arg_boolFirstRun=False, arg_clssSqliteQueue=None):
        #Chama o método para subir a fila, apenas se for a primeira vez
        if(arg_boolFirstRun):
            self.add_to_queue(arg_clssSqliteQueue)

        #Edite o valor dessa variável a no arquivo Config.xlsx
        var_intMaxTentativas = self.var_dictConfig["MaxRetryNumber"]
        
        for var_intTentativa in range(var_intMaxTentativas):
            try:
                self.var_clssMaestro.write_log("Iniciando aplicativos, tentativa " + (var_intTentativa+1).__str__())
                #Insira aqui seu código para iniciar os aplicativos
            except BusinessRuleException as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro de negócio: " + str(exception), arg_enumLogLevel=LogLevel.ERROR)

                raise
            except Exception as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro, tentativa " + (var_intTentativa+1).__str__() + ": " + str(exception), arg_enumLogLevel=LogLevel.ERROR)

                if(var_intTentativa+1 == var_intMaxTentativas): raise
                else: 
                    #inclua aqui seu código para tentar novamente
                    
                    continue
            else:
                self.var_clssMaestro.write_log("Aplicativos iniciados, continuando processamento...")
                break
            