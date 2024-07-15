from PyQt5 import QtCore, QtGui, QtWidgets
import EnviarEmail, Dados
import telaInicial


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(660, 650)

        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, 10, 640, 30))
        self.label.setObjectName("label")
        self.label.setText("Fornecedores selecionados")
        self.label.setFont(QtGui.QFont("Arial", 14))

        self.listView = QtWidgets.QListView(Form)
        self.listView.setGeometry(QtCore.QRect(10, 50, 640, 540))
        self.listView.setObjectName("listView")

        self.pushButtonCancelar = QtWidgets.QPushButton(Form)
        self.pushButtonCancelar.setGeometry(QtCore.QRect(0, 600, 330, 50))
        self.pushButtonCancelar.setObjectName("pushButtonCancelar")
        self.pushButtonCancelar.setText("Cancelar")

        self.pushButtonEnviar = QtWidgets.QPushButton(Form)
        self.pushButtonEnviar.setGeometry(QtCore.QRect(330, 600, 330, 50))
        self.pushButtonEnviar.setObjectName("pushButtonEnviar")
        self.pushButtonEnviar.setText("Enviar Emails")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Emails para fornecedores do Bloco K"))


class Enviar(QtWidgets.QWidget):
    def __init__(self, fornecedores):
        super().__init__()

        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.setFixedSize(660, 650)

        self.fornecedores = fornecedores

        self.model = QtGui.QStandardItemModel()
        for item_text in fornecedores:
            item = QtGui.QStandardItem(item_text)
            item.setEditable(False)
            self.model.appendRow(item)
        self.ui.listView.setModel(self.model)
        self.ui.pushButtonEnviar.clicked.connect(self.btnEnviar)
        self.ui.pushButtonCancelar.clicked.connect(self.btnCancelar)


    def btnEnviar(self):
        fornecedores_com_erro = []
        for fornecedor in self.fornecedores:
            try:
                dados_tabela_vencido = Dados.ver_informacoes_necessarias(Dados.ver_prazos_vencidos(fornecedor))
                dados_tabela_15 = Dados.ver_informacoes_necessarias(Dados.ver_prazos_15_dias_para_vencer(fornecedor))
                EnviarEmail.enviar_email(fornecedor, dados_tabela_vencido, dados_tabela_15)
            except Exception as e:
                fornecedores_com_erro.append(fornecedor)

        if fornecedores_com_erro:
            self.mensagem_enviada_com_erro(fornecedores_com_erro)

        else: 
            self.mensagem_enviada()

        self.primeira_tela = telaInicial.Inicial()
        self.primeira_tela.show()
        self.close()





    def btnCancelar(self):
        self.primeira_tela = telaInicial.Inicial()
        self.primeira_tela.show()
        self.close()
        

    def mensagem_enviada(self):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Information)
        msg_box.setText("Emails enviados com sucesso!")
        msg_box.setWindowTitle("Confirmação")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg_box.exec_()

    def mensagem_enviada_com_erro(self, fornecedores):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        fornecedores_str = "\n".join(fornecedores)
        msg_box.setText("Os emails foram enviados, porém alguns fornecedores tiveram problema:\n" + fornecedores_str)
        msg_box.setWindowTitle("Confirmação")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg_box.exec_()



