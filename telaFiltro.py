from PyQt5 import QtCore, QtGui, QtWidgets
import Dados, telaInicial


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(660, 650)
        self.listView = QtWidgets.QListView(Form)
        self.listView.setGeometry(QtCore.QRect(0, 0, 660, 600))
        self.listView.setObjectName("listView")

        self.pushButtonLimpar = QtWidgets.QPushButton(Form)
        self.pushButtonLimpar.setGeometry(QtCore.QRect(0, 600, 330, 50))
        self.pushButtonLimpar.setObjectName("pushButtonLimpar")
        self.pushButtonLimpar.setText("Limpar filtros")

        self.pushButtonFiltrar = QtWidgets.QPushButton(Form)
        self.pushButtonFiltrar.setGeometry(QtCore.QRect(330, 600, 330, 50))
        self.pushButtonFiltrar.setObjectName("pushButtonFiltrar")
        self.pushButtonFiltrar.setText("Filtrar Fornecedores")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Emails para fornecedores do Bloco K", "Emails para fornecedores do Bloco K"))


class Filtro(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.setFixedSize(660, 650)
        data = Dados.buscar_depositos()
        self.model = QtGui.QStandardItemModel()
        for item_text in data:
            item = QtGui.QStandardItem(item_text)
            item.setCheckable(True)
            item.setEditable(False)
            self.model.appendRow(item)
        self.ui.listView.setModel(self.model)
        self.ui.pushButtonLimpar.clicked.connect(self.btnLimpar)
        self.ui.pushButtonFiltrar.clicked.connect(self.btnFiltrar)


    def btnLimpar(self):
        self.tela_inicial = telaInicial.Inicial()
        self.tela_inicial.show()
        self.close()


    def btnFiltrar(self):
        selected_indexes = [index for index in range(self.model.rowCount()) if self.model.item(index).checkState() == QtCore.Qt.Checked]
        selected_items = [self.model.item(index).text() for index in selected_indexes]
        if not selected_items:
            self.tela_inicial = telaInicial.Inicial()
        else:
            self.tela_inicial = telaInicial.Inicial(selected_items)
        self.tela_inicial.show()
        self.close()
        self.mensagem_filtros(selected_items)


    def mensagem_filtros(self, depositos):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Information)
        
        if not depositos:
            msg_box.setText("Nenhum filtro realizado")
        else:
            msg_box.setText(f"Filtros Realizados: {', '.join(depositos)}")
            
        msg_box.setWindowTitle("Confirmação")
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg_box.exec_()


