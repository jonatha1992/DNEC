from datetime import datetime
from PyQt6.QtCore import QMetaObject, QRect, QSize, Qt, QCoreApplication 
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QCheckBox, QDateEdit, QLabel, QMainWindow, 
    QProgressBar, QPushButton, QTableView, 
    QVBoxLayout, QWidget
)
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(750, 335)
        MainWindow.setProperty(u"fixedSize", QSize(720, 400))
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.status_label = QLabel(self.centralwidget)
        self.status_label.setObjectName(u"status_label")
        self.status_label.setGeometry(QRect(330, 40, 181, 16))
        font = QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setUnderline(True)
        font.setStrikeOut(False)
        self.status_label.setFont(font)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setGeometry(QRect(20, 300, 591, 23))
        self.progressBar.setValue(24)
        self.progressBar.setTextVisible(True)
        self.btn_delete = QPushButton(self.centralwidget)
        self.btn_delete.setObjectName(u"btn_delete")
        self.btn_delete.setGeometry(QRect(640, 240, 75, 24))
        self.status_label_2 = QLabel(self.centralwidget)
        self.status_label_2.setObjectName(u"status_label_2")
        self.status_label_2.setGeometry(QRect(20, 40, 181, 16))
        font1 = QFont()
        font1.setPointSize(10)
        font1.setBold(True)
        font1.setUnderline(True)
        self.status_label_2.setFont(font1)
        self.status_label_2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_status_process = QLabel(self.centralwidget)
        self.label_status_process.setObjectName(u"label_status_process")
        self.label_status_process.setGeometry(QRect(190, 280, 231, 16))
        self.label_status_process.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.btn_select_files = QPushButton(self.centralwidget)
        self.btn_select_files.setObjectName(u"btn_select_files")
        self.btn_select_files.setGeometry(QRect(220, 240, 75, 24))
        self.layoutWidget = QWidget(self.centralwidget)
        self.layoutWidget.setObjectName(u"layoutWidget")
        self.layoutWidget.setGeometry(QRect(20, 60, 181, 170))
        self.verticalLayout = QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.label_start = QLabel(self.layoutWidget)
        self.label_start.setObjectName(u"label_start")

        self.verticalLayout.addWidget(self.label_start)

        self.date_start = QDateEdit(self.layoutWidget)
        self.date_start.setObjectName(u"date_start")
        self.date_start.setCalendarPopup(True)

        self.verticalLayout.addWidget(self.date_start)

        self.label_end = QLabel(self.layoutWidget)
        self.label_end.setObjectName(u"label_end")

        self.verticalLayout.addWidget(self.label_end)

        self.date_end = QDateEdit(self.layoutWidget)
        self.date_end.setObjectName(u"date_end")
        self.date_end.setCalendarPopup(True)

        self.verticalLayout.addWidget(self.date_end)



        self.chk_verificar = QCheckBox(self.layoutWidget)
        self.chk_verificar.setObjectName(u"chk_verificar")
        self.chk_verificar.setChecked(True)

        self.verticalLayout.addWidget(self.chk_verificar)

        self.btn_process = QPushButton(self.centralwidget)
        self.btn_process.setObjectName(u"btn_process")
        self.btn_process.setGeometry(QRect(20, 240, 75, 24))
        self.btn_save = QPushButton(self.centralwidget)
        self.btn_save.setObjectName(u"btn_save")
        self.btn_save.setGeometry(QRect(130, 240, 75, 24))
        self.tableView = QTableView(self.centralwidget)
        self.tableView.setObjectName(u"tableView")
        self.tableView.setGeometry(QRect(220, 60, 500, 171))
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
        
        self.progressBar.setVisible(False)
        current_date = datetime.now()
        first_day_of_month = current_date.replace(day=1)

        self.date_start.setDate(first_day_of_month)
        self.date_end.setDate(current_date)
        
        
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"Consolidado Semanal", None))
        self.status_label.setText(QCoreApplication.translate("MainWindow", u"Cantidad de Archivos", None))
        self.btn_delete.setText(QCoreApplication.translate("MainWindow", u"Borrar", None))
        self.status_label_2.setText(QCoreApplication.translate("MainWindow", u"Parametros", None))
        self.label_status_process.setText("")
        self.btn_select_files.setText(QCoreApplication.translate("MainWindow", u"Seleccionar", None))
        self.label_start.setText(QCoreApplication.translate("MainWindow", u"Fecha inicial:", None))
        self.label_end.setText(QCoreApplication.translate("MainWindow", u"Fecha final:", None))
        self.chk_verificar.setText(QCoreApplication.translate("MainWindow", u"Verificar con base existente", None))
        self.btn_process.setText(QCoreApplication.translate("MainWindow", u"Procesar", None))
        self.btn_save.setText(QCoreApplication.translate("MainWindow", u"Guardar", None))
    # retranslateUi
