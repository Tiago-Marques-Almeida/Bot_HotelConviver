from botcity.core import DesktopBot
import os
import time
from datetime import datetime, timedelta
import pyautogui

import config
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def __init__(self):
        super().__init__()
        self.dia_anterior = datetime.now() - timedelta (1)
        self.data = self.dia_anterior.strftime('%d-%m-%Y')
        self.cria_diretorio()
        
    def cria_diretorio(self):
        caminho = (r'C:\Robo\BotHotelConviver\arquivos'+ '\\' + self.data)
        if not os.path.isdir(caminho):
            os.mkdir(caminho)


    def abrir_probrama(self):
        print(datetime.now())

        #Minimiza todos os programas e move o mause para o canto da tela
        pyautogui.keyDown('win')
        pyautogui.press('d')
        pyautogui.keyUp('win')
        time.sleep(5)
        pyautogui.moveTo(1,1)

        #abre Erbon Suite
        self.execute(r'C:\Users\Pc2\Desktop\Erbon Suite.appref-ms')
        time.sleep(15)
        
    def verifica_windows_protegeu(self):
       
        try:
            #clica em mais informações 
            if not self.find( "mais_informacoes", matching=0.97, waiting_time=10000):
                self.not_found("mais_informacoes")
            self.click_relative(56, 10)
             
        #clica em continuar assim mesmo
            if not self.find( "executar_assim_mesmo", matching=0.97, waiting_time=10000):
                self.not_found("executar_assim_mesmo")
            self.click_relative(91, 30)
            
        except:
            print('Hoje o Windows não protegeu!')
    
    
    def verifica_atualizacao(self):

        #verifica se ocorreu atualizaçao e aguarda atualizar  
        if not self.find( "icone_atualizacao", matching=0.97, waiting_time=10000):
            print('Nao Houve Atualizaçao Hoje!')
        else: 
            #aguardando a atualizaçao finalizar
            while True:
                if not self.find( "mais_informacoes", matching=0.97, waiting_time=10000):
                    print('atualizaçao em progresso...')
                    time.sleep(180)
            
            #apos finalizada a atualizaçao, clicar para aceitar a inicializaçao
            if not self.find( "mais_informacoes", matching=0.97, waiting_time=10000):
                self.not_found("mais_informacoes")
            self.click_relative(56, 10)
            
            if not self.find( "executar_assim_mesmo_2", matching=0.97, waiting_time=10000):
                self.not_found("executar_assim_mesmo_2")
            self.click_relative(77, 24)
                       
            time.sleep(10)
            


    def loguin(self):
     
        #loguin Suíte
        if not self.find( "usuario", matching=0.97, waiting_time=30000):
            self.not_found("usuario")        
        self.click()

        #insere o usuario
        self.paste('equipejuliano')
        time.sleep(2)
        #clica no campo de senha
        if not self.find( "senha", matching=0.97, waiting_time=30000):
            self.not_found("senha")
        self.click()

        #insere a senha
        self.paste('conviver5151')

        #clica em logar
        if not self.find( "loguin", matching=0.97, waiting_time=30000):
            self.not_found("loguin")
        self.click()


    def acessa_PMS(self):
        # def acessa_relatorios(self):27-04-2022
        dictPassos = config.dPassos
        for key in dictPassos:
            if not getattr(self,'find')(dictPassos[key]['imagem'], matching=0.97, waiting_time=30000):
                self.not_found(key)
            if dictPassos[key]['is_relativo']:
                self.click_relative(*dictPassos[key]['cordenadas'])
            else:
                self.click()


    def action(self, execution=None):
        self.abrir_probrama()
        self.verifica_windows_protegeu()
        self.verifica_atualizacao()
        self.loguin()
        self.acessa_PMS()
        self.Baixando_relatorio_PDV()
        self.Baixando_relatorio_RESUMO_DIARIO()
        self.Baixando_relatorio_VENDAS_DIARIAS_MENSAL()
        self.Baixando_VENDAS_PRODUÇAO_GERAl()
        self.Baixando_Relatorio_ESTORNOS()
        self.relatorio_financeiro()
        
        

    def Baixando_relatorio_PDV(self):

        #escreve o nome do relatório
        self.paste('relat')
        #clica no relatorio
        if not self.find( "relatorio_caixa_pdv", matching=0.97, waiting_time=30000):
            self.not_found("relatorio_caixa_pdv")
        self.click(clicks=2)

        #seleciona todos os usuários
        if not self.find( "selecionar_todos", matching=0.97, waiting_time=30000):
            self.not_found("selecionar_todos")
        self.click()

        #seleciona a inicial data
        if not self.find( "datainicialrelpdv", matching=0.97, waiting_time=30000):
            self.not_found("datainicialrelpdv")
        self.click_relative(11, 13)
        

        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        #clica data final
        if not self.find( "data_final", matching=0.97, waiting_time=30000):
            self.not_found("data_final")
        self.click_relative(8, 15)
        

        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        #clica em confirmar
        if not self.find( "confirmar", matching=0.97, waiting_time=30000):
            self.not_found("confirmar")
        self.click()

        self.exporta_arquivo('Relatorio_Caixa_PDV')

        #clica em pesquisar
        if not self.find( "pesquisar2", matching=0.97, waiting_time=30000):
            self.not_found("pesquisar2")
        self.click_relative(29, 15)

        self.control_a()
        self.delete()


    def Baixando_relatorio_RESUMO_DIARIO(self):

        #pesquisando novo relatório
        self.paste('resum')
        #abrindo novo relatório
        if not self.find( "resumo_diario", matching=0.97, waiting_time=30000):
            self.not_found("resumo_diario")
        self.click(clicks=2)

        #clica na data
        if not self.find( "data22", matching=0.97, waiting_time=30000):
            self.not_found("data22")
        self.click_relative(13, 11)
        


        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        #clica em confirmar
        if not self.find( "confirmar2", matching=0.97, waiting_time=30000):
            self.not_found("confirmar2")
        self.click_relative(57, 20)

        self.exporta_arquivo('Relatorio_Resumo_Avancado_Diario')

        #clica em pesquisar
        if not self.find( "pesquisa123", matching=0.97, waiting_time=30000):
            self.not_found("pesquisa123")
        self.click_relative(117, 186)

        self.control_a()
        self.delete()

    def Baixando_relatorio_VENDAS_DIARIAS_MENSAL(self):

        #pesquisando novo relatório
        self.paste('venda')
        #movendo o mouse
        if not self.find( "move_mause_1", matching=0.97, waiting_time=30000):
            self.not_found("move_mause_1")
        self.move()

        self.scroll_down(clicks=20)

        #abrindo relatório
        if not self.find( "vendas_diarias_mensal", matching=0.97, waiting_time=30000):
            self.not_found("vendas_diarias_mensal")
        self.click(clicks=2)


        #abrindo leque de meses
        if not self.find( "mes_leque", matching=0.97, waiting_time=30000):
            self.not_found("mes_leque")
        self.click_relative(95, 36, clicks=2)

        #selecionando o mês
        mes = datetime.today().strftime('%m')
        pyautogui.press('down', presses= int(mes))

        #Abrindo leque de anos
        if not self.find( "ano", matching=0.97, waiting_time=30000):
            self.not_found("ano")
        self.click_relative(80, 38, clicks=2)

        #sellecionando o Ano
        ano = datetime.today().strftime('%Y')
        pyautogui.write(str(ano))

        #clica em confirmar
        if not self.find( "confirmar2", matching=0.97, waiting_time=30000):
            self.not_found("confirmar2")
        self.click_relative(55, 23)

        self.exporta_arquivo('relatorio_VENDAS_DIARIAS_MENSAL')

        #clica em pesquisar
        if not self.find( "pesquisa123", matching=0.97, waiting_time=30000):
            self.not_found("pesquisa123")
        self.click_relative(117, 186)

        self.control_a()
        self.delete()

    def Baixando_VENDAS_PRODUÇAO_GERAl(self):

        #pesquisando novo relatório
        self.paste('venda')
        #abrindo novo relatório
        if not self.find( "relatoprodmen", matching=0.97, waiting_time=30000):
            self.not_found("relatoprodmen")
        self.click_relative(66, 18, clicks=2)
        
        
        #clicando na data inicial
        if not self.find( "data333", matching=0.97, waiting_time=30000):
            self.not_found("data333")
        self.click_relative(9, 10)

        #apagando a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        #clicando na data final
        if not self.find( "data4444", matching=0.97, waiting_time=30000):
            self.not_found("data4444")
        self.click_relative(24, 86)
        

        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.dia_anterior.strftime('%d/%m/%Y'))
        #clica em confirmar
        if not self.find( "confirmar345", matching=0.97, waiting_time=30000):
            self.not_found("confirmar345")
        self.click_relative(57, 11)


        self.exporta_arquivo('relatorio_VENDAS_PRODUÇAO_GERAl')

        if not self.find( "pesquisa123", matching=0.97, waiting_time=30000):
            self.not_found("pesquisa123")
        self.click_relative(117, 186)
        self.control_a()
        self.delete()


    def Baixando_Relatorio_ESTORNOS(self):
        
        #pesquisa estornos
        self.paste('estor')
        #clica para abrir o relatório
        if not self.find( "estorno", matching=0.97, waiting_time=30000):
            self.not_found("estorno")
        self.click(clicks=2)

        #clicando na data inicial
        if not self.find( "data12345", matching=0.97, waiting_time=30000):
            self.not_found("data12345")
        self.click_relative(15, 9)
        

        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        #clicando na data final
        if not self.find( "dataaa", matching=0.97, waiting_time=30000):
            self.not_found("dataaa")
        self.click_relative(13, 18)
        

        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.dia_anterior.strftime('%d/%m/%Y'))
        #clica em confirmar
        if not self.find( "confir", matching=0.97, waiting_time=30000):
            self.not_found("confir")
        self.click_relative(51, 16)

        self.exporta_arquivo('relatorio_ESTORNO')
        
        #fecha Erbon PMS       
        self.alt_f4()      
        
        
        #confirmar fechamento PMS
        if not self.find( "sairsomente", matching=0.97, waiting_time=30000):
            self.not_found("sairsomente")
        self.click_relative(62, 17)


    def relatorio_financeiro(self):
        
        #clica na seta para abrir as opções
        if not self.find( "setadown1", matching=0.97, waiting_time=30000):
            self.not_found("setadown1")
        self.click_relative(32, 3)
        
        
        '''if not self.find( "setadown", matching=0.97, waiting_time=30000):
            self.not_found("setadown")
        self.click_relative(22, 7)'''

        #clica no botão financeiro
        if not self.find( "financeirobotao", matching=0.97, waiting_time=30000):
            self.not_found("financeirobotao")
        self.click()

        #clica em contas a receber
        if not self.find( "contreceber", matching=0.97, waiting_time=30000):
            self.not_found("contreceber")
        self.click_relative(53, 13)

        #clica em pesquisar tudo
        if not self.find( "pesqTUdo", matching=0.97, waiting_time=30000):
            self.not_found("pesqTUdo")
        self.click()

        #espera modal desaparecer
        time.sleep(10)
        
        #clica na data
        if not self.find( "datafinaldobot", matching=0.97, waiting_time=10000):
            self.not_found("datafinaldobot")
        self.click_relative(7, 20)
        
        #apaga a data
        self.control_a()
        self.delete()
        #escreve a data de ontem
        self.paste(self.data)
        time.sleep(1)
        #clica em selecionar todos
        if not self.find( "selalltodosfin", matching=0.97, waiting_time=30000):
            self.not_found("selalltodosfin")
        self.click()

        #clica em excel
        if not self.find( "excelfinalfin", matching=0.97, waiting_time=30000):
            self.not_found("excelfinalfin")
        self.click_relative(32, 24)

        try:
            #clica no botão para exportar
            self.exporta_arquivo('relatorio_FINANCEIRO')
            
            self.alt_f4()        

            #CLICA PARA FECHAR eRBON
            if not self.find( "fechandotudodevez", matching=0.97, waiting_time=30000):
                self.not_found("fechandotudodevez")
            self.click_relative(106, 15)

            #fecha tudo 100 por cento
            if not self.find( "saidetudoagorafoi", matching=0.97, waiting_time=30000):
                self.not_found("saidetudoagorafoi")
            self.click()
        except:
            self.alt_f4()        

            #CLICA PARA FECHAR eRBON
            if not self.find( "fechandotudodevez", matching=0.97, waiting_time=30000):
                self.not_found("fechandotudodevez")
            self.click_relative(106, 15)

            #fecha tudo 100 por cento
            if not self.find( "saidetudoagorafoi", matching=0.97, waiting_time=30000):
                self.not_found("saidetudoagorafoi")
            self.click()


        


    def exporta_arquivo(self, nome_arquivo):
        time.sleep(3)
        #clica no botão para exportar
        if not self.find( "botao_exportar", matching=0.97, waiting_time=30000):
            self.not_found("botao_exportar")
        self.click()
        time.sleep(1)
        #clica em exportar
        if not self.find( "excel", matching=0.97, waiting_time=30000):
            self.not_found("excel")
        self.click()
        time.sleep(1)
        #clica no diretorio
        if not self.find( "diretoriopsalva", matching=0.97, waiting_time=30000):
            self.not_found("diretoriopsalva")
        self.click_relative(13, 19)
        
        time.sleep(1)
        self.paste(r'C:\Robo\BotHotelConviver\arquivos' + '\\' + self.data)
        self.enter()
        #renomeia o arquivo
        if not self.find( "camponomearq", matching=0.97, waiting_time=30000):
            self.not_found("camponomearq")
        self.click_relative(5, 37)
        
        time.sleep(1)
        self.control_a()
        self.delete()
        self.paste(nome_arquivo)
        time.sleep(1)
        #clica em salvar
        if not self.find( "salvaecanc", matching=0.97, waiting_time=30000):
            self.not_found("salvaecanc")
        self.click_relative(42, 11)
        time.sleep(1)
        #caso já haja arquivo na pasta
        if self.find( "simenao", matching=0.97, waiting_time=1000):
            self.click_relative(36, 12)
        else:
            self.not_found("simenao")
        
        #tempo de espera para fechar janela
        time.sleep(2)
        #clica em fechar relatório
        self.alt_f4()

        

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    tentativas = 0
    while (tentativas<10):
        try:
            Bot.main()
            tentativas = 10
        except Exception as e:
            print(e)
            tentativas += 1
            if (tentativas >=10):
                with open (r'C:\Robo\BotHotelConviver\log\log.txt', 'w') as arq:
                    arq.write(str(e))
            
        os.system("taskkill /F /IM HotelboxSuite.exe")
        time.sleep(30)
        

    











































































