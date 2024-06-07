from PyQt5 import QtCore, QtGui, QtWidgets
import Dados, EnviarEmail


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(700, 700)
        self.listView = QtWidgets.QListView(Form)
        self.listView.setGeometry(QtCore.QRect(10, 20, 600, 600))
        self.listView.setObjectName("listView")

        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(650, 650, 50, 30))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setText("Enviar")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))


class MyForm(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.ui = Ui_Form()
        self.ui.setupUi(self)

        # Adicionando dados à ListView com caixas de seleção
        data = Dados.buscar_fornecedores()
        model = QtGui.QStandardItemModel()
        for item_text in data:
            item = QtGui.QStandardItem(item_text)
            item.setCheckable(True)  
            item.setEditable(False)  
            model.appendRow(item)
        self.ui.listView.setModel(model)
        self.ui.pushButton.clicked.connect(self.print_selected_items)

    def print_selected_items(self):
        model = self.ui.listView.model()
        selected_indexes = [index for index in range(model.rowCount()) if model.item(index).checkState() == QtCore.Qt.Checked]
        selected_items = [model.item(index).text() for index in selected_indexes]
        print("Itens selecionados:")
        for item in selected_items:
            dados_tabela = Dados.ver_informacoes_necessarias(Dados.ver_prazos_vencidos(item))
            EnviarEmail.enviar_emil(item, dados_tabela)
        


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    form = MyForm()
    form.show()
    sys.exit(app.exec_())
