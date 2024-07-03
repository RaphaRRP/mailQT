from PyQt5 import QtCore, QtGui, QtWidgets
import telaInicial


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(660, 650)

        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, -25, 645, 650))
        self.label.setObjectName("label")
        self.label.setText(
"""
Está com problema?                       
Este Programa se guia por um arquivo excel e o nome de certos elementos,
certifique que esses nomes estão corretos:
                           
Pasta do arquivo: 
S:/PM/ter/tef/tef3/Inter_Setor/TEF3 Controles/Bloco K Planejamento
                           
Nome do arquivo:
BLOCO K - Planejamento.xlsx
                           
Nome da tabela:
FUP Prazos
                           
Coluna com nome dos fornecedorer:
Razão Social
                           
Coluna com deposito:
Depósito
                           
Coluna com o email do fornecedor:
E-mail Fornecedor
                           
Coluna com o prazo do fornecimento:
Prazo
                           
Se o problema perssistir, aqui esta meu contato: 
email: raphael.prates@bosch.com | chat:  prr8ca@bosch.com
""")
        self.label.setFont(QtGui.QFont("Arial", 9))

        self.pushButtonVoltar = QtWidgets.QPushButton(Form)
        self.pushButtonVoltar.setGeometry(QtCore.QRect(0, 600, 660, 50))
        self.pushButtonVoltar.setObjectName("pushButtonVoltar")
        self.pushButtonVoltar.setText("Voltar")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Emails para fornecedores do Bloco K"))


class Ajuda(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.setFixedSize(660, 650)
        self.ui.pushButtonVoltar.clicked.connect(self.btnVoltar)


    def btnVoltar(self):
        self.primeira_tela = telaInicial.Inicial()
        self.primeira_tela.show()
        self.close()
        
