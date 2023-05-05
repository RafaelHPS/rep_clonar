"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install -e .`
- Use the same interpreter as the one used to install the bot (`pip install -e .`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

from botcity.web import WebBot, Browser
from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
from botcity.maestro import *

from framework.classes_t2c.email.T2CSendEmail import T2CSendEmail
from framework.classes_t2c.email.T2CSendEmailOutlook import T2CSendEmailOutlook
from framework.classes_t2c.T2CInitAllSettings import T2CInitAllSettings
from framework.classes_t2c.sqlite.T2CSqliteQueue import T2CSqliteQueue
from framework.classes_t2c.relatorios.T2CRelatorios import T2CRelatorios
from framework.classes_t2c.T2CInitAllApplications import T2CInitAllApplications
from framework.classes_t2c.T2CCloseAllApplications import T2CCloseAllApplications
from framework.classes_t2c.T2CKillAllProcesses import T2CKillAllProcesses
from framework.classes_t2c.T2CProcess import T2CProcess
from framework.classes_t2c.utils.T2CExceptions import BusinessRuleException
from framework.classes_t2c.utils.T2CMaestro import T2CMaestro, LogLevel
from framework.classes_t2c.sqlserver.T2CSqlAnaliticoSintetico import T2CSqlAnaliticoSintetico
from webdriver_manager import chrome, firefox, microsoft
import datetime

class Bot(WebBot):
    def action(self, execution=None):
        # Configure whether or not to run on headless mode
        self.headless = False
        
        #Carregando o arquivo Config num dicionário
        var_dictConfig = T2CInitAllSettings().load_config()

        #Iniciando a classe para controle do maestro
        var_clssMaestro = T2CMaestro(arg_clssExecution=execution, arg_dictConfig=var_dictConfig)
        
        #RunnerId é o nome do runner se estiver rodando pelo maestro, mas se estiver rodando localmente, ele pega o nome da máquina
        var_strNomeMaquina:str = var_clssMaestro.var_strRunnerId

        #Iniciando variáveis de controle
        var_dateDatahoraInicio = datetime.datetime.now()
        var_strDatahoraInicio = var_dateDatahoraInicio.strftime("%d/%m/%Y %H:%M:%S")
        var_strNomeProcesso:str = var_dictConfig["NomeProcesso"]
       
        #Variáveis contadores já levadas em conta no framework
        var_intQtdeItensProcessados = 0
        var_intQtdeItensAppException = 0
        var_intQtdeItensBusinessException = 0
        var_intQtdeItensSucesso = 0
        
        #Variáveis que precisam ser levadas em conta pelo desenvolvedor (somar contadores e indicar uso ou não)
        var_boolUsaCaptcha = False
        var_boolUsaOcr = False
        var_boolUsaApi = False
        var_intQtdeCaptcha = 0
        var_intQtdeOcr = 0
        var_intQtdeApi = 0

        # Instantiate a DesktopBot
        desktop_bot = DesktopBot()

        #CONFIGURAÇÕES DE BROWSER
        #Descomente para usar o browser escolhido
        #self.browser = Browser.CHROME
        #self.browser = Browser.FIREFOX
        #self.browser = Browser.EDGE

        #Descomente para baixar e usar a versão do webdriver do seu navegador escolhido
        #self.driver_path = chrome.ChromeDriverManager().install()              #Chrome
        #self.driver_path = firefox.GeckoDriverManager().install()              #Firefox
        #self.driver_path = microsoft.EdgeChromiumDriverManager().install()     #Edge

        #Iniciando classes
        var_strCaminhoBd = var_dictConfig["CaminhoBancoSqlite"] if(var_dictConfig.__contains__("CaminhoBancoSqlite")) else None
        var_strTabelaFila = var_dictConfig["FilaProcessamento"] if(var_dictConfig.__contains__("FilaProcessamento")) else "tbl_Fila_Processamento"
        var_clssSqliteQueue = T2CSqliteQueue(arg_strCaminhoBd=var_strCaminhoBd, arg_strTabelaFila=var_strTabelaFila, arg_strNomeMaquina=var_strNomeMaquina, arg_clssMaestro=var_clssMaestro)
        var_clssSqlAnaliticoSintetico = T2CSqlAnaliticoSintetico(arg_clssMaestro=var_clssMaestro, arg_dictConfig=var_dictConfig)

        var_clssRelatorios = T2CRelatorios(arg_dictConfig=var_dictConfig)
        var_clssInitAllApplications = T2CInitAllApplications(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro, arg_clssSqliteQueue=var_clssSqliteQueue)
        var_clssCloseAllApplications = T2CCloseAllApplications(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)
        var_clssKillAllProcesses = T2CKillAllProcesses(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)
        var_clssProcess = T2CProcess(arg_dictConfig=var_dictConfig, arg_botWebbot=self, arg_botDesktopbot=desktop_bot, arg_clssMaestro=var_clssMaestro)

        #Se for indicando que é necessário incluir linhas no SQL Server, insere a primeira aqui
        if(var_dictConfig["NecesSQLServer"] == "Sim"):
            var_clssSqlAnaliticoSintetico.insert_linha_inicio_sintetico(arg_strNomeMaquina=var_strNomeMaquina, arg_boolUsaCaptcha=var_boolUsaCaptcha, 
                                                                        arg_boolUsaOCR=var_boolUsaOcr, arg_boolUsaAPI=var_boolUsaApi, arg_dateInicioExecucao=var_dateDatahoraInicio)

        #CLASSE COMENTADA POR PRECISAR DE PARÂMETROS CUSTOMIZADOS
        #DESCOMENTAR EM DESENVOLVIMENTO, apenas o que for necessário
        #var_clssSendEmail = T2CSendEmail(arg_strEmailServerSmtp="", arg_intEmailPortaSmtp="", arg_strUsuario="", arg_strSenha="",arg_strNomeProcesso=var_strNomeProcesso, arg_clssMaestro=var_clssMaestro)
        #var_clssSendEmailOutlook = T2CSendEmailOutlook(arg_strNomeProcesso=var_strNomeProcesso, arg_clssMaestro=var_clssMaestro)

        #Enviando email de inicialização
        #if(var_dictConfig["EmailInicial"] == "Sim"):
            #var_clssSendEmail.send_email_inicial(arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
            #var_clssSendEmailOutlook.send_email_inicial(arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])

        var_clssMaestro.write_log("Iniciando processamento: " + var_strNomeProcesso)
        
        try:
            var_clssInitAllApplications.execute(arg_boolFirstRun=True)
            var_clssMaestro.write_log("Itens encontrados na fila: " + var_clssSqliteQueue.var_intItemsQueue.__str__())
        except BusinessRuleException as exception:
            #incluindo erro no relatório sintético, com horário de finalização
            var_dateDatahoraFim = datetime.datetime.now()
            var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
            var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
            var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]
            var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, var_strDatahoraInicio, var_strDatahoraFim, var_strTotalProcessamento, var_intQtdeItensProcessados, var_intQtdeItensSucesso, var_intQtdeItensBusinessException, var_intQtdeItensAppException, var_strNomeMaquina])

            #Tirando print do erro antes de fechar
            var_strCaminhoScreeshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreeshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
            desktop_bot.save_screenshot(path=var_strCaminhoScreeshot)
            
            #Se for indicando que é necessário incluir linhas no SQL Server, faz update com os dados finais aqui
            if(var_dictConfig["NecesSQLServer"] == "Sim"):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, arg_intQtdeCaptcha=var_intQtdeCaptcha, arg_intQtdeApi=var_intQtdeApi, arg_intQtdeOcr=var_intQtdeApi, 
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException, 
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, arg_intTotalItensSucesso=var_intQtdeItensSucesso)

            #Enviando o email final por causa do erro na inicialização
            #COMENTADA PARA EVITAR ERROS QUANDO CLASSE NÂO INICIALIZADA, DESCOMENTAR A QUE PRECISAR
            #if(var_dictConfig["EmailFinal"] == "Sim"):
                #var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strNomeProcesso=var_strNomeProcesso, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=False)
                #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico, var_strCaminhoScreeshot], arg_boolSucesso=False)

            var_clssMaestro.finish_task(arg_boolSucesso=False, arg_strMensagem="Task finalizada na inicialização por erro de negócio, verifique os logs de execução.")
            raise
        except Exception as exception:
            #incluindo erro no relatório sintético, com horário de finalização
            var_dateDatahoraFim = datetime.datetime.now()
            var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
            var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
            var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]
            var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, var_strDatahoraInicio, var_strDatahoraFim, var_strTotalProcessamento, var_intQtdeItensProcessados, var_intQtdeItensSucesso, var_intQtdeItensBusinessException, var_intQtdeItensAppException, var_strNomeMaquina])

            #Tirando print do erro antes de fechar
            var_strCaminhoScreeshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreeshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
            desktop_bot.save_screenshot(path=var_strCaminhoScreeshot)
            
            #Se for indicando que é necessário incluir linhas no SQL Server, faz update com os dados finais aqui
            if(var_dictConfig["NecesSQLServer"] == "Sim"):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, arg_intQtdeCaptcha=var_intQtdeCaptcha, arg_intQtdeApi=var_intQtdeApi, arg_intQtdeOcr=var_intQtdeApi, 
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException, 
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, arg_intTotalItensSucesso=var_intQtdeItensSucesso)


            #Enviando o email final por causa do erro na inicialização
            #COMENTADA PARA EVITAR ERROS QUANDO CLASSE NÂO INICIALIZADA, DESCOMENTAR A QUE PRECISAR
            #if(var_dictConfig["EmailFinal"] == "Sim"):
                #var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strNomeProcesso=var_strNomeProcesso, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=False)
                #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico, var_strCaminhoScreeshot], arg_boolSucesso=False)

            var_clssMaestro.finish_task(arg_boolSucesso=False, arg_strMensagem="Task finalizada na inicialização por erro desconhecido, verifique os logs de execução.")
            raise

        #PROCESS

        #Processamento continua até a classe informar que não existe itens novos para processar
        var_tplQueueItem = var_clssSqliteQueue.get_next_queue_item()

        if(var_tplQueueItem is not None): var_clssMaestro.write_log("Iniciando processamento dos itens encontrados na fila")
        else: var_clssMaestro.write_log("Não existem itens para ser processados")

        while(var_tplQueueItem is not None):
            for var_intTentativa in range(var_dictConfig["MaxRetryNumber"]):
                try:
                    var_intQtdeItensProcessados += 1
                    var_dateDatahoraInicio_Item = datetime.datetime.now()
                    var_strDatahoraInicio_Item = var_dateDatahoraInicio_Item.strftime("%d/%m/%Y %H:%M:%S")
                    var_clssMaestro.write_log(arg_strMensagemLog="Executando item, referência: " + var_tplQueueItem[1] + ", tentativa: " + (var_intTentativa+1).__str__(), arg_strReferencia=var_tplQueueItem[1])
                    #Executando o process
                    var_clssProcess.execute(arg_tplQueueItem=var_tplQueueItem)
                except BusinessRuleException as exception:
                    var_intQtdeItensBusinessException += 1
                    var_clssMaestro.write_log(arg_strMensagemLog="Erro de negócio: " + str(exception), arg_strReferencia=var_tplQueueItem[1], arg_enumLogLevel=LogLevel.ERROR)
                    
                    #Chamando a classe KillAllProcesses em caso de erro de business
                    var_clssKillAllProcesses.execute()
                    var_clssInitAllApplications.execute()

                    #Incluindo linha no relatório analítico sinalizando erro de aplicação
                    var_dateDatahoraFim_Item = datetime.datetime.now()
                    var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")
                    var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, var_strDatahoraFim_Item, var_tplQueueItem[0], var_tplQueueItem[1], var_strNomeMaquina, "ERRO - REGRA DE NEGÓCIO", exception.__str__()])

                    #Marcando item como erro de business
                    var_clssSqliteQueue.update_status_item(arg_intIndex=var_tplQueueItem[0], arg_strNovoStatus="BUSINESS ERROR", arg_strObs=exception.__str__())

                    #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                    if(var_dictConfig["NecesSQLServer"] == "Sim"):
                        var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_tplItemFila=var_tplQueueItem, arg_dateInicioItem=var_dateDatahoraInicio_Item, arg_dateFimItem=var_dateDatahoraFim_Item, 
                                                                             arg_strNomeFila=var_strTabelaFila, arg_strStatusItem="BUSINESS ERROR", arg_strTipoExcecao="REGRA DE NEGOCIO", arg_strDescricaoExcecao=exception.__str__())

                    #Tirando print
                    var_strCaminhoScreeshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreeshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
                    desktop_bot.save_screenshot(path=var_strCaminhoScreeshot)
                        
                    #Enviando email com print do erro
                    #if(var_dictConfig["EmailCadaErro"] == "Sim"):
                        #var_clssSendEmail.send_email_erro(arg_boolBusiness=True, arg_listAnexos=[var_strCaminhoScreeshot], arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                        #var_clssSendEmailOutlook.send_email_erro(arg_boolBusiness=True, arg_listAnexos=[var_strCaminhoScreeshot], arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                    
                    break
                except Exception as exception:
                    var_intQtdeItensAppException +=1

                    var_clssMaestro.write_log(arg_strMensagemLog="Erro, tentativa " + (var_intTentativa+1).__str__() + ": " + str(exception), arg_strReferencia=var_tplQueueItem[1], arg_enumLogLevel=LogLevel.ERROR)
                    
                    #Chamando a classe KillAllProcesses em caso de erro de aplicação
                    var_clssKillAllProcesses.execute()

                    #Incluindo linha no relatório analítico sinalizando erro de aplicação
                    var_dateDatahoraFim_Item = datetime.datetime.now()
                    var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")
                    var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, var_strDatahoraFim_Item, var_tplQueueItem[0], var_tplQueueItem[1], var_strNomeMaquina, "ERRO - APLICAÇÃO", exception.__str__()])

                    #Marcando item como erro
                    var_clssSqliteQueue.update_status_item(arg_intIndex=var_tplQueueItem[0], arg_strNovoStatus="APP ERROR", arg_strObs=exception.__str__())

                    #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                    if(var_dictConfig["NecesSQLServer"] == "Sim"):
                        var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_tplItemFila=var_tplQueueItem, arg_dateInicioItem=var_dateDatahoraInicio_Item, arg_dateFimItem=var_dateDatahoraFim_Item, 
                                                                             arg_strNomeFila=var_strTabelaFila, arg_strStatusItem="APP ERROR", arg_strTipoExcecao="APLICACAO", arg_strDescricaoExcecao=exception.__str__())

                    #Iniciando novamente o processamento
                    var_clssInitAllApplications.execute()

                    if(var_intTentativa+1 == var_dictConfig["MaxRetryNumber"]):
                        #Tirando print
                        var_strCaminhoScreeshot = var_dictConfig["CaminhoExceptionScreenshots"] + "ExceptionScreeshot_" + datetime.datetime.now().strftime("%d%m%Y_%H%M%S%f") + ".png"
                        desktop_bot.save_screenshot(path=var_strCaminhoScreeshot)

                        #Enviando email com print do erro
                        #if(var_dictConfig["EmailCadaErro"] == "Sim"):           
                            #var_clssSendEmail.send_email_erro(arg_boolBusiness=False, arg_listAnexos=[var_strCaminhoScreeshot], arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])
                            #var_clssSendEmailOutlook.send_email_erro(arg_boolBusiness=False, arg_listAnexos=[var_strCaminhoScreeshot], arg_strDetalhesErro=str(exception), arg_strEnvioPara=var_dictConfig["EmailDestinatarios"])

                    else: continue
                else:
                    #Insira aqui qualquer cóigo necessário para voltar ao estado inicial em caso de sucesso, para executar um possivel próximo item



                    #Inluindo linha no relatório analítico sinalizando sucesso
                    var_intQtdeItensSucesso += 1
                    var_dateDatahoraFim_Item = datetime.datetime.now()
                    var_strDatahoraFim_Item = var_dateDatahoraFim_Item.strftime("%d/%m/%Y %H:%M:%S")
                    var_clssRelatorios.inserir_linha_analitico(arg_listValores=[var_strDatahoraInicio_Item, var_strDatahoraFim_Item, var_tplQueueItem[0], var_tplQueueItem[1], var_strNomeMaquina, "SUCESSO", ""])
                    
                    #Marcando item como finalizado com sucesso
                    var_clssSqliteQueue.update_status_item(arg_intIndex=var_tplQueueItem[0], arg_strNovoStatus="SUCCESS")
                    
                    #Se for indicando que é necessário incluir linhas no SQL Server, faz o insert com os dados do item aqui
                    if(var_dictConfig["NecesSQLServer"] == "Sim"):
                        var_clssSqlAnaliticoSintetico.insert_linha_analitico(arg_tplItemFila=var_tplQueueItem, arg_dateInicioItem=var_dateDatahoraInicio_Item, 
                                                                             arg_dateFimItem=var_dateDatahoraFim_Item, arg_strNomeFila=var_strTabelaFila, arg_strStatusItem="SUCCESSO")

                    break
            
            var_tplQueueItem = var_clssSqliteQueue.get_next_queue_item()

        #END

        #Fechando aplicativos no final do processamento
        try:
            var_clssCloseAllApplications.execute()
        except Exception:
            var_clssMaestro.write_log(arg_strMensagemLog="Fechando aplicativos pelo KillAllProcesses", arg_enumLogLevel=LogLevel.WARN)
            var_clssKillAllProcesses.execute()

        #Preenchendo linha no relatório sintético, incluindo horário de finalização
        var_clssMaestro.write_log("Inserindo linhas no relatório analítico e enviando email")
        var_dateDatahoraFim = datetime.datetime.now()
        var_strDatahoraFim = var_dateDatahoraFim.strftime("%d/%m/%Y %H:%M:%S")
        var_strTotalProcessamento = str(var_dateDatahoraFim - var_dateDatahoraInicio)
        var_strTotalProcessamento = var_strTotalProcessamento.split('.')[0]
        var_clssRelatorios.inserir_linha_sintetico(arg_listValores=[var_strNomeProcesso, var_strDatahoraInicio, var_strDatahoraFim, var_strTotalProcessamento, var_intQtdeItensProcessados, var_intQtdeItensSucesso, var_intQtdeItensBusinessException, var_intQtdeItensAppException, var_strNomeMaquina])

        #Se for indicando que é necessário incluir linhas no SQL Server, faz update com os dados finais aqui
        if(var_dictConfig["NecesSQLServer"] == "Sim"):
                var_clssSqlAnaliticoSintetico.update_linha_fim_sintetico(arg_dateFimExecucao=var_dateDatahoraFim, arg_intQtdeCaptcha=var_intQtdeCaptcha, arg_intQtdeApi=var_intQtdeApi, arg_intQtdeOcr=var_intQtdeApi, 
                                                                         arg_intTotalItens=var_intQtdeItensProcessados, arg_intTotalItensBusinessEx=var_intQtdeItensBusinessException, 
                                                                         arg_intTotalItensAppEx=var_intQtdeItensAppException, arg_intTotalItensSucesso=var_intQtdeItensSucesso)

        #Enviando email final com os relatórios analítico e sintético
        #if(var_dictConfig["EmailFinal"] == "Sim"):    
            #var_clssSendEmail.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=True)
            #var_clssSendEmailOutlook.send_email_final(arg_strHorarioInicio=var_strDatahoraInicio, arg_strHorarioFim=var_strDatahoraFim, arg_strEnvioPara=var_dictConfig["EmailDestinatarios"], arg_listAnexos=[var_clssRelatorios.var_strPathRelatorioAnalitico, var_clssRelatorios.var_strPathRelatorioSintetico], arg_boolSucesso=True)
    
        var_clssMaestro.write_log("Finalizando processamento.")
        var_clssMaestro.finish_task(arg_boolSucesso=True, arg_strMensagem="Task finished OK.")
        

if __name__ == '__main__':
    Bot.main()
