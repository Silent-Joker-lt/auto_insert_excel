# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'windowdesign.ui'
#
# Created by: PyQt5 UI code generator 5.12.3
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
import excelcool
from pandas import read_excel
import numpy as np
from pandas import DataFrame


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1200, 846)
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(30)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(40, 120, 311, 211))
        self.groupBox.setObjectName("groupBox")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setGeometry(QtCore.QRect(23, 21, 51, 41))
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit.setGeometry(QtCore.QRect(80, 30, 111, 41))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_2.setGeometry(QtCore.QRect(80, 100, 111, 41))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setGeometry(QtCore.QRect(23, 91, 51, 41))
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self.groupBox)
        self.pushButton.setGeometry(QtCore.QRect(80, 160, 91, 41))
        self.pushButton.setObjectName("pushButton")
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(400, 120, 231, 211))
        self.groupBox_2.setObjectName("groupBox_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox_2)
        self.lineEdit_3.setGeometry(QtCore.QRect(70, 30, 111, 41))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        self.label_3.setGeometry(QtCore.QRect(13, 21, 51, 41))
        self.label_3.setObjectName("label_3")
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox_2)
        self.pushButton_2.setGeometry(QtCore.QRect(70, 100, 91, 41))
        self.pushButton_2.setObjectName("pushButton_2")
        self.tableWidget = QtWidgets.QTableWidget(Form)
        self.tableWidget.setGeometry(QtCore.QRect(40, 390, 1000, 431))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(370, 350, 141, 31))
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")

        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setFont(font)
        self.label_5.setGeometry(QtCore.QRect(210, 30, 400, 61))
        self.label_5.setObjectName("label_5")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "每次登统计都很麻烦，做个小东西看看能不能比较快捷的完成信息的统计。^_^!"))
        self.groupBox.setTitle(_translate("Form", "单表"))
        self.label.setText(_translate("Form", "样表名称"))
        self.label_2.setText(_translate("Form", "填表人"))
        self.pushButton.setText(_translate("Form", "确定"))
        self.groupBox_2.setTitle(_translate("Form", "统表"))
        self.label_3.setText(_translate("Form", "样表名称"))
        self.pushButton_2.setText(_translate("Form", "确定"))
        self.label_4.setText(_translate("Form", "填表结果核对"))
        self.label_5.setText(_translate("Form", "自动填表Demo"))

        self.pushButton.clicked.connect(self.soloexcel)

    def soloexcel(self):
        targetexcel=self.lineEdit.text()
        someone=self.lineEdit_2.text()
        excelcool.openxl(targetexcel,someone)
        df = read_excel("{}.xlsx".format(someone))
        rows = len(df)
        columns = len(df.columns)
        someone_data = DataFrame.from_records(df)
        someone_data = np.array(someone_data).tolist()
        self.tableWidget.setRowCount(rows)
        self.tableWidget.setColumnCount(columns)
        self.tableWidget.setAlternatingRowColors(True)
        for i in range(rows):
            for j in range(self.tableWidget.columnCount()):
                data = QtWidgets.QTableWidgetItem(str(someone_data[i][j]))
                self.tableWidget.setItem(i, j, data)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_Form()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())