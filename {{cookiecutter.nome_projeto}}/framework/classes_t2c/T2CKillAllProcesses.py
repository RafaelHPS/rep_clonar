from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
from framework.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel
from framework.classes_t2c.utils.T2CExceptions import BusinessRuleException

#Classe feita para ser invocada em casos de system exceptions no processamento, para resetar o processamento
#Pode ser usado em outras partes do processo também, dependendo de como a automação for programada
class T2CKillAllProcesses:
    #Iniciando a classe, pedindo um dicionário config e o bot que vai ser usado e enviando uma exceção caso nenhum for informado
    def __init__(self, arg_dictConfig:dict, arg_clssMaestro:T2CMaestro, arg_botWebbot:WebBot=None, arg_botDesktopbot:DesktopBot=None):
        if(arg_botWebbot is None and arg_botDesktopbot is None): raise Exception("Não foi possível inicializar a classe, forneça pelo menos um bot")
        else:
            self.var_botWebbot = arg_botWebbot
            self.var_botDesktopbot = arg_botDesktopbot
            self.var_dictConfig = arg_dictConfig
            self.var_clssMaestro = arg_clssMaestro

    #Método que roda para finalizar os processos necessários, apenas com a estrutura em código
    def execute(self):
        #Edite o valor dessa variável a no arquivo Config.xlsx
        var_intMaxTentativas = self.var_dictConfig["MaxRetryNumber"]

        for var_intTentativa in range(var_intMaxTentativas):
            try:
                self.var_clssMaestro.write_log("Finalizando processos, tentativa " + (var_intTentativa+1).__str__())
                #Insira aqui seu código para finalizar processos
            except BusinessRuleException as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro de negócio: " + str(exception), arg_enumLogLevel=LogLevel.ERROR)

                raise
            except Exception as exception:
                self.var_clssMaestro.write_log(arg_strMensagemLog="Erro, tentativa " + (var_intTentativa+1).__str__() + ": " + str(exception), arg_enumLogLevel=LogLevel.ERROR)
                
                if(var_intTentativa+1 == var_intMaxTentativas): raise
                else: 
                    #Incluir aqui seu código para tentar novamente
                    
                    continue
            else:
                self.var_clssMaestro.write_log("Aplicativos finalizados, continuando processamento...")
                break
            