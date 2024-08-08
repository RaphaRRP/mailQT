from PyQt5 import QtCore, QtGui, QtWidgets
import Dados, telaEnviar, telaAjuda, telaFiltro 


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(660, 650)
        self.listView = QtWidgets.QListView(Form)
        self.listView.setGeometry(QtCore.QRect(0, 0, 660, 600))
        self.listView.setObjectName("listView")

        self.pushButtonAjuda = QtWidgets.QPushButton(Form)
        self.pushButtonAjuda.setGeometry(QtCore.QRect(0, 600, 110, 50))
        self.pushButtonAjuda.setObjectName("pushButtonAjuda")
        self.pushButtonAjuda.setText("Ajuda")

        self.pushButtonFiltrar = QtWidgets.QPushButton(Form)
        self.pushButtonFiltrar.setGeometry(QtCore.QRect(110, 600, 110, 50))
        self.pushButtonFiltrar.setObjectName("pushButtonFiltrar")
        self.pushButtonFiltrar.setText("Filtrar")

        self.pushButtonSelecionar = QtWidgets.QPushButton(Form)
        self.pushButtonSelecionar.setGeometry(QtCore.QRect(220, 600, 220, 50))
        self.pushButtonSelecionar.setObjectName("pushButtonSelecionar")
        self.pushButtonSelecionar.setText("Selecionar Todos")

        self.pushButtonEnviar = QtWidgets.QPushButton(Form)
        self.pushButtonEnviar.setGeometry(QtCore.QRect(440, 600, 220, 50))
        self.pushButtonEnviar.setObjectName("pushButtonEnviar")
        self.pushButtonEnviar.setText("Enviar Emails")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Emails para fornecedores do Bloco K", "Emails para fornecedores do Bloco K"))


class Inicial(QtWidgets.QWidget):
    def __init__(self, depositos=None):
        super().__init__()

        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.setFixedSize(660, 650)
        self.depositos = depositos
        data = Dados.buscar_fornecedores(depositos=self.depositos)
        self.model = QtGui.QStandardItemModel()
        for item_text in data:
            item = QtGui.QStandardItem(str(item_text))
            item.setCheckable(True)
            item.setEditable(False)
            self.model.appendRow(item)
        self.ui.listView.setModel(self.model)
        self.ui.pushButtonEnviar.clicked.connect(self.btnEnviar)
        self.ui.pushButtonSelecionar.clicked.connect(self.btnSelecionar)
        self.ui.pushButtonAjuda.clicked.connect(self.btnAjuda)
        self.ui.pushButtonFiltrar.clicked.connect(self.btnFiltrar)


    def btnSelecionar(self):
        all_selected = all(self.model.item(index).checkState() == QtCore.Qt.Checked for index in range(self.model.rowCount()))
        new_state = QtCore.Qt.Unchecked if all_selected else QtCore.Qt.Checked
        for index in range(self.model.rowCount()):
            item = self.model.item(index)
            item.setCheckState(new_state)
        self.ui.pushButtonSelecionar.setText("Selecionar Todos" if all_selected else "Desselecionar Todos")


    def btnEnviar(self):
        selected_indexes = [index for index in range(self.model.rowCount()) if self.model.item(index).checkState() == QtCore.Qt.Checked]
        selected_items = [self.model.item(index).text() for index in selected_indexes]
        if not selected_items:
            QtWidgets.QMessageBox.warning(self, "Aviso", "Selecione pelo menos um Fornecedor")
            return
        self.tela_enviar = telaEnviar.Enviar(selected_items)
        self.tela_enviar.show()
        self.close()


    def btnAjuda(self):
        self.tela_ajuda = telaAjuda.Ajuda()
        self.tela_ajuda.show()
        self.close()


    def btnFiltrar(self):
        self.tela_filtro = telaFiltro.Filtro()
        self.tela_filtro.show()
        self.close()



if __name__ == "__main__":
    import sys
    try:
        app = QtWidgets.QApplication(sys.argv)
        form = Inicial()
        form.show()
        sys.exit(app.exec_())
    except:
        app = QtWidgets.QApplication(sys.argv)
        form = telaAjuda.Ajuda()
        form.show()
        sys.exit(app.exec_())