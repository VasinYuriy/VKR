# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'languageChoiceWidget.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_languageChoiceWidget(object):
    def setupUi(self, languageChoiceWidget):
        languageChoiceWidget.setObjectName("languageChoiceWidget")
        languageChoiceWidget.resize(810, 136)
        self.horizontalLayout = QtWidgets.QHBoxLayout(languageChoiceWidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.studentName = QtWidgets.QLabel(languageChoiceWidget)
        self.studentName.setObjectName("studentName")
        self.horizontalLayout.addWidget(self.studentName)
        self.languageChoice = QtWidgets.QComboBox(languageChoiceWidget)
        self.languageChoice.setObjectName("languageChoice")
        self.languageChoice.addItem("")
        self.languageChoice.addItem("")
        self.languageChoice.addItem("")
        self.horizontalLayout.addWidget(self.languageChoice)
        self.languageChoiceDiploma = QtWidgets.QComboBox(languageChoiceWidget)
        self.languageChoiceDiploma.setObjectName("languageChoiceDiploma")
        self.languageChoiceDiploma.addItem("")
        self.languageChoiceDiploma.addItem("")
        self.languageChoiceDiploma.addItem("")
        self.languageChoiceDiploma.addItem("")
        self.horizontalLayout.addWidget(self.languageChoiceDiploma)

        self.retranslateUi(languageChoiceWidget)
        QtCore.QMetaObject.connectSlotsByName(languageChoiceWidget)

    def retranslateUi(self, languageChoiceWidget):
        _translate = QtCore.QCoreApplication.translate
        languageChoiceWidget.setWindowTitle(_translate("languageChoiceWidget", "Form"))
        self.studentName.setText(_translate("languageChoiceWidget", "TextLabel"))
        self.languageChoice.setItemText(0, _translate("languageChoiceWidget", "английский"))
        self.languageChoice.setItemText(1, _translate("languageChoiceWidget", "немецкий"))
        self.languageChoice.setItemText(2, _translate("languageChoiceWidget", "французский"))
        self.languageChoiceDiploma.setItemText(0, _translate("languageChoiceWidget", "русский"))
        self.languageChoiceDiploma.setItemText(1, _translate("languageChoiceWidget", "английский"))
        self.languageChoiceDiploma.setItemText(2, _translate("languageChoiceWidget", "немецкий"))
        self.languageChoiceDiploma.setItemText(3, _translate("languageChoiceWidget", "французский"))
