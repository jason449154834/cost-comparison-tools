from PySide2.QtWidgets import QApplication,QMainWindow, QPushButton,QPlainTextEdit,QMessageBox,QTableWidgetItem,QFileDialog,QInputDialog,QLineEdit
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Qt
from PySide2.QtGui import QColor
from window_control.BOQ_compare import *
from database_process import *

class main_window():
    def __init__(self):
        self.ui = QUiLoader().load('UI/cost_compare.ui')
        ####################################################################################
        #导入基准项目
        self.ui.pushButton_upload_base.clicked.connect(self.pushButton_upload_base_clicked)
        #导入对比项目
        self.ui.pushButton_upload_compare.clicked.connect(self.pushButton_upload_compare_clicked)
        #取值方法
        self.ui.comboBox_compare_type.currentIndexChanged.connect(self.comboBox_compare_type_currentIndexChanged)
        #分析
        self.ui.pushButton_compare_begin.clicked.connect(self.pushButton_compare_begin_clicked)
        
        ####################################################################################
        #页面初始化
        self.main_window_show()
        
        #页面初始化
    def main_window_show(self):
        self.file_Path_base,self.file_Path_compare = '',''
        self.compare_type = 0
        
        ####################################################################################
        #导入基准项目
    def pushButton_upload_base_clicked(self):
        self.file_Path_base = upload_excel(self.ui)
        self.ui.label_base_add.setText(self.file_Path_base)
        print('导入基准项目ok!')

        #导入对比项目
    def pushButton_upload_compare_clicked(self):
        self.file_Path_compare = upload_excel(self.ui)
        self.ui.label_compare_add.setText(self.file_Path_compare)
        print('导入对比项目ok!')

        #分析
    def pushButton_compare_begin_clicked(self):
        if len(self.file_Path_base) > 0 and len(self.file_Path_compare) > 0 :
            self.ui.label_state.setText('对比开始')
            compare_BOQ_beign(self.file_Path_base,self.file_Path_compare,self.ui,self.compare_type)
        else:
            QMessageBox.information(self.ui,'提示','未选择项目！')

    
        #取值方法
    def comboBox_compare_type_currentIndexChanged(self,index):
        self.compare_type = index
        print('取值方法:',index)


        
        
            
       