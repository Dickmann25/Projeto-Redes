import email
from socket import *
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, uic
from PyQt5.QtWidgets import QWidget
from PyQt5 import QtWidgets
import ssl
from base64 import*
from email.header import decode_header


class MyGUi(QMainWindow):

    def __init__(self):
        super(MyGUi,self).__init__()
        uic.loadUi("email_client.ui", self)
        self.show()

        #atributos
        self.socket_application = None
        self.nome_application = None
        self.senha_application = None
        self.assunto_application = None
        self.corpo_application = None
        self.linhas_tabela = -1
        self.resultado_pesquisa = []

        #inicia botões desligadosna aplicação smtp
        self.enviar_botao.setEnabled(False)
        self.Destinatario.setEnabled(False)
        self.text_destinatario.setEnabled(False)
        self.Assunto.setEnabled(False)
        self.text_assunto.setEnabled(False)
        self.Corpo.setEnabled(False)
        self.text_corpo.setEnabled(False)
        self.deslogar_smtp.setEnabled(False)

        #inicia botões desligadosna aplicação imap
        self.comboBox.setEnabled(False)
        self.tableWidget.setEnabled(False)
        self.search_bar.setEnabled(False)
        self.pesquisar.setEnabled(False)
        self.mostrar.setEnabled(False)
        self.deslogar_imap.setEnabled(False)

        #funções
        header = self.tableWidget.horizontalHeader()       
        self.tableWidget.setColumnWidth(0,200)
        self.tableWidget.hideColumn(2)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        self.actionEnviar.triggered.connect(self.mEnviar)
        self.actionChecar.triggered.connect(self.mChecar)
        self.login_botao.clicked.connect(self.loginSmtp)
        self.login_botao_2.clicked.connect(self.loginImap)
        self.deslogar_smtp.clicked.connect(self.deslogarSmtp)
        self.deslogar_imap.clicked.connect(self.deslogarImap)
        self.mostrar.clicked.connect(self.mostrarMais)
        self.enviar_botao.clicked.connect(self.enviar)
        self.pesquisar.clicked.connect(self.cPesquisar)
        self.tableWidget.setRowCount(10)
        self.tableWidget.cellDoubleClicked.connect(self.getClickedCell)
        self.voltar.clicked.connect(self.retornar)
        
    def getClickedCell(self, row):
        #identificação do email
        id = (self.tableWidget.item(row, 2)).text()

        clientSSL = self.socket_application

        #garante que a caixa principal esta selecionada
        clientSSL.sendall('a002 SELECT INBOX\r\n'.encode('utf-8'))
        answer = self.resposta(clientSSL,"a002")

        clientSSL.sendall(f"a003 FETCH {id} RFC822\r\n".encode('utf-8'))
        answer = self.resposta_byte(clientSSL,"a003")

        answer = answer.removesuffix(b")\r\na003 OK Success\r\n")
        element_to_find = b'}\r\n'

        index = answer.find(element_to_find)

        answer = answer[index + len(element_to_find):]

        answer = email.message_from_bytes(answer) 

        answer = self.get_body(answer).decode('utf-8') 

        print(answer)

        self.stackedWidget.setCurrentWidget(self.Conteudo)
        self.menuBar.setVisible(False)
        self.email_text.setPlainText(answer)

    def retornar(self):
        self.stackedWidget.setCurrentWidget(self.Checar)

    def mEnviar(self):
        if self.stackedWidget.currentIndex() == 1:
            try:
                #fecha conexão
                clientSSL = self.socket_application
                quit = "a100 LOGOUT\r\n"
                clientSSL.send(quit.encode())
                answer = self.resposta(clientSSL,"a100")
                print(answer)
                
                self.socket_application = None
                self.nome_application = None
                clientSSL.close()

                self.comboBox.setEnabled(False)
                self.tableWidget.setEnabled(False)
                self.search_bar.setEnabled(False)
                self.pesquisar.setEnabled(False)
                self.mostrar.setEnabled(False)
                
            except:  
                pass  

        #troca de página    
        self.stackedWidget.setCurrentWidget(self.Enviar)
        print(self.stackedWidget.currentIndex())

    def mChecar(self):
        if self.stackedWidget.currentIndex() == 0:
            try:
                #fecha conexão
                clientSSL = self.socket_application
                quit = "QUIT\r\n"
                clientSSL.send(quit.encode())
                answer = clientSSL.recv(1024).decode()
                print(answer,"11")

                self.socket_application = None
                self.nome_application = None

                self.enviar_botao.setEnabled(False)
                self.Destinatario.setEnabled(False)
                self.text_destinatario.setEnabled(False)
                self.Assunto.setEnabled(False)
                self.text_assunto.setEnabled(False)
                self.Corpo.setEnabled(False)
                self.text_corpo.setEnabled(False)

                clientSSL.close()
            except:  
                pass 
                
        #troca de página        
        self.stackedWidget.setCurrentWidget(self.Checar)
        print(self.stackedWidget.currentIndex())

    def loginSmtp(self):
        #pega o email do usuario
        nome = self.text_endereco.text()
        self.nome_application = nome

        #pega a senha do usuario
        senha = self.text_senha.text()

        #servidor e port que usaremos
        email_server = 'smtp.gmail.com'
        email_port = 587

        #Cria conexão tcp
        client_socket = socket(AF_INET, SOCK_STREAM)
        client_socket.connect((email_server, email_port))
        answer = client_socket.recv(1024).decode()
        print(answer,"1")

        #Manda ola para o servidor
        helo = "HELO Juan\r\n"
        client_socket.send(helo.encode())
        answer = client_socket.recv(1024).decode()
        print(answer,"2")

        #inicia o TLS por segurança
        strttlscmd = "STARTTLS\r\n".encode()
        client_socket.send(strttlscmd)
        answer = client_socket.recv(1024).decode()
        print(answer,"3")

        clientSSL = ssl.wrap_socket(client_socket)

        nome = b64encode(nome.encode())
        senha = b64encode(senha.encode())

        #autentifica o usuario
        authorizationCMD = "AUTH LOGIN\r\n"
        
        self.socket_application = clientSSL
        
        clientSSL.send(authorizationCMD.encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"4")

        clientSSL.send(nome + "\r\n".encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"5")

        clientSSL.send(senha + "\r\n".encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"6")
        if "535-5.7.8 Username and Password not accepted. Learn more at" in answer :
            QMessageBox.about(self, "Erro", "Endereço de Email ou Senha incorretos")

            #fecha conexão
            quit = "QUIT\r\n"
            clientSSL.send(quit.encode())
            answer = clientSSL.recv(1024).decode()
            print(answer,"11")
            self.socket_application = None
            self.nome_application = None
            clientSSL.close()
        else:
            QMessageBox.about(self, "Success", "Endereço válido, escreva sua mensagem")
            self.enviar_botao.setEnabled(True)
            self.Destinatario.setEnabled(True)
            self.text_destinatario.setEnabled(True)
            self.Assunto.setEnabled(True)
            self.text_assunto.setEnabled(True)
            self.Corpo.setEnabled(True)
            self.text_corpo.setEnabled(True)
            self.deslogar_smtp.setEnabled(True)
            
    def loginImap(self):
        #pega o email do usuario
        nome = self.text_endereco_2.text()
        self.nome_application = nome

        #pega a senha do usuario
        senha = self.text_senha_2.text()

        #servidor e port que usaremos
        email_server = 'imap.gmail.com'
        email_port = 993
    
        #Cria conexão tcp
        client_socket = socket(AF_INET, SOCK_STREAM)
        client_socket.connect((email_server, email_port))

        #inicia o TLS por segurança
        context = ssl.create_default_context()
        clientSSL = context.wrap_socket(client_socket, server_hostname=email_server)
        answer = clientSSL.recv(4096).decode()
        print(answer)

        self.socket_application = clientSSL

        #autentifica o usuario
        clientSSL.sendall(f'a001 LOGIN {nome} {senha}\r\n'.encode())
        answer = self.resposta(clientSSL,"a001")
        print(answer)

        if "BAD" in answer or "NO" in answer :
            QMessageBox.about(self, "Erro", "Endereço de Email ou Senha Incorretos")

            #fecha conexão
            quit = "a100 LOGOUT\r\n"
            clientSSL.send(quit.encode())
            answer = self.resposta(clientSSL,"a100")
            print(answer)

            self.socket_application = None
            self.nome_application = None
            clientSSL.close()
        else:
            QMessageBox.about(self, "Success", "Endereço Válido")
            self.comboBox.setEnabled(True)
            self.tableWidget.setEnabled(True)
            self.search_bar.setEnabled(True)
            self.pesquisar.setEnabled(True)
            self.deslogar_imap.setEnabled(True)

    def deslogarSmtp(self):
        clientSSL = self.socket_application

        quit = "QUIT\r\n"
        clientSSL.send(quit.encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"11")
        self.socket_application = None
        self.nome_application = None
        clientSSL.close()
        
        self.enviar_botao.setEnabled(False)
        self.Destinatario.setEnabled(False)
        self.text_destinatario.setEnabled(False)
        self.Assunto.setEnabled(False)
        self.text_assunto.setEnabled(False)
        self.Corpo.setEnabled(False)
        self.text_corpo.setEnabled(False)
        self.deslogar_smtp.setEnabled(False)

    def deslogarImap(self):  
        clientSSL = self.socket_application

        quit = "a100 LOGOUT\r\n"
        clientSSL.send(quit.encode())
        answer = self.resposta(clientSSL,"a100")
        print(answer)

        self.socket_application = None
        self.nome_application = None
        clientSSL.close()
        
        self.comboBox.setEnabled(False)
        self.tableWidget.setEnabled(False)
        self.search_bar.setEnabled(False)
        self.pesquisar.setEnabled(False)
        self.mostrar.setEnabled(False)
        self.deslogar_imap.setEnabled(False)
        
        self.tableWidget.setRowCount(0)
        self.tableWidget.setRowCount(10)

    def enviar(self):
        clientSSL = self.socket_application

        #pega o email do usuario
        nome = self.nome_application

        #pega o email do destinatario
        destinatario = self.text_destinatario.text()

        #pega o assunto e o corpo da mensagem
        assunto = self.text_assunto.text()
        corpo = self.text_corpo.toPlainText()

        #indica quem esta mandando a mensagem
        remetente = f"Mail from: <{nome}>\r\n"
        clientSSL.send(remetente.encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"7")
        
        #indica quem deve receber a mensagem
        rcpt = f"RCPT TO: <{destinatario}>\r\n"
        clientSSL.send(rcpt.encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"8")

        #pede permissão para mandar mensagem
        conteudo = "DATA\r\n"
        clientSSL.send(conteudo.encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"9")

        #manda mensagem
        mensagem = f"Subject: {assunto}\n\n{corpo}"
        clientSSL.send(mensagem.encode())
        clientSSL.send("\r\n.\r\n".encode())
        answer = clientSSL.recv(1024).decode()
        print(answer,"10")
        if "250" in answer or "354" in answer:
            QMessageBox.about(self, "Success", "Mensagem enviada com sucesso")
        
        else:
            QMessageBox.about(self, "Error", "Ocorreu um erro no envio da mensagem")
        
    def cPesquisar(self, type_search):
        clientSSL = self.socket_application

        #seleciona a caixa principal
        clientSSL.sendall('a002 SELECT INBOX\r\n'.encode('utf-8'))
        answer = self.resposta(clientSSL,"a002")
        print(answer)

        #contem o parametro que o usuario gostaria de usar
        type_search = self.comboBox.currentText()

        if type_search == "Pesquise por Assunto":
            search_value = self.search_bar.text()
            clientSSL.sendall(f'a004 SEARCH SUBJECT "{search_value}"\r\n'.encode())
            answer = self.resposta(clientSSL,"a004")
            answer = answer.split()
            answer = [x for x in answer if x.isnumeric()]
            self.linhas_tabela = -1
            self.tableWidget.setRowCount(10)

            #disponibiliza os resultados na tabela
            self.showTable(answer, clientSSL)

            
        elif type_search == "Pesquise por Destinatario":
            search_value = self.search_bar.text()
            clientSSL.sendall(f'a004 SEARCH FROM "{search_value}"\r\n'.encode())
            answer = self.resposta(clientSSL,"a004")
            answer = answer.split(" ")
            answer = [x for x in answer if x.isnumeric()]
            self.linhas_tabela = -1
            self.tableWidget.setRowCount(10)

            #disponibiliza os resultados na tabela
            self.showTable(answer, clientSSL)

        elif type_search == "Pesquise por Conteudo":
            search_value = self.search_bar.text()
            clientSSL.sendall(f'a004 SEARCH BODY "{search_value}"\r\n'.encode())
            answer = self.resposta(clientSSL,"a004")
            answer = answer.split()
            answer = [x for x in answer if x.isnumeric()]
            self.linhas_tabela = -1
            self.tableWidget.setRowCount(10)

            #disponibiliza os resultados na tabela
            self.showTable(answer, clientSSL)

    def resposta(self, clientSSL, tag):
        #essa função é usada para garantir que recebemos a resposta completa do servidor

        texto =""
        while True:
            answer = clientSSL.recv(4096).decode()
            texto = texto + answer
            if f"{tag} OK" in answer or f"{tag} BAD" in answer or f"{tag} NO" in answer:
                return texto

    def showTable(self, texto, clientSSL):
        QMessageBox.about(self, "Loading", "Buscando Emails, Isso Pode Demorar um Pouco")

        #garante que a caixa principal esta selecionada
        clientSSL.sendall('a002 SELECT INBOX\r\n'.encode('utf-8'))
        answer = self.resposta(clientSSL,"a002")
        row = self.linhas_tabela + 1
        contador = 0
        self.resultado_pesquisa = texto[10:].copy()
        if self.resultado_pesquisa == []:
            self.mostrar.setEnabled(False)

        else:
            self.mostrar.setEnabled(True)

        #busca tanto o remetente quanto o assunto do email e disponibiliza na tabela
        for i in texto:
            if contador == 10:
                break

            try:
                clientSSL.sendall(f"a003 FETCH {i} BODY[HEADER.FIELDS (FROM)]\r\n".encode('utf-8'))
                answer = self.resposta(clientSSL,"a003")

            except:
                print("fetch failed")

            if "a003 OK" in answer:
                    #processo de formatação da mensagem
                    pivo = []
                    answer = answer.split("\n")
                    answer = [x for x in answer if "From: " in x]
                    answer = answer[0].replace("From: ", "")
                     
                    self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(answer))

                    #disponibiliza em uma coluna escondida o numero do email
                    self.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(i))
        
        
            clientSSL.sendall(f"a003 FETCH {i} BODY[HEADER.FIELDS (SUBJECT)]\r\n".encode('utf-8'))
            answer = self.resposta(clientSSL,"a003")
            if "a003 OK" in answer:
                    pivo = []
                    answer = answer.split("\n")
                    answer[1] = answer[1].replace("Subject: ","")
                    for x in answer[1:]:
                        if x == "\r":
                            answer = ' '.join(pivo)

                            break
                        pivo.append(x)
                        
                    if "UTF-8" in answer:
                        answer = " ".join(text.decode(encoding or 'utf-8') for text, encoding in decode_header(answer))
                    elif "ISO-8859-1" in answer:
                        answer = " ".join(text.decode(encoding or 'utf-8') for text, encoding in decode_header(answer))
            
                    self.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(answer))
                    row += 1
                    contador += 1
                    self.linhas_tabela += 1

    def mostrarMais(self):
        #função usada para mostrar os proximos 10 emails na pesquisa
        self.tableWidget.setRowCount(min(self.linhas_tabela + 11, self.linhas_tabela + len(self.resultado_pesquisa)))
        answer = self.resultado_pesquisa
        clientSSL = self.socket_application
        self.showTable(answer, clientSSL)               

    def get_body(self,msg):
            if msg.is_multipart():
                return self.get_body(msg.get_payload(0))
            else:
                return msg.get_payload(None,True)
    
    def resposta_byte(self, clientSSL, tag):
        texto =bytes()
        while True:
            answer = clientSSL.recv(4096)
            texto = texto + answer
            if f"{tag} OK".encode() in answer or f"{tag} BAD".encode() in answer or f"{tag} NO".encode() in answer:
                return texto


app = QApplication([])
window = MyGUi()
app.exec_()