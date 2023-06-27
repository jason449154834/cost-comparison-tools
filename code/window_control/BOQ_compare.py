from database_process import *
from PySide2.QtWidgets import QApplication,QMainWindow, QPushButton,QPlainTextEdit,QMessageBox,QTableWidgetItem,QFileDialog,QInputDialog,QLineEdit
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Qt
from PySide2.QtGui import QColor
import xlrd
import xlwt
import xlsxwriter
import openpyxl
import numpy as np
import pandas as pd
from transformers import AutoTokenizer
from collections import Counter
from sklearn.metrics.pairwise import cosine_similarity

#Excel通用模型
class excel_model():
    def __init__(self,save_path):
        #Excel输出
        self.wb_output = xlsxwriter.Workbook(save_path, options={
            # 全局设置
            'nan_inf_to_errors': True,
            'strings_to_numbers': False,  # str 类型数字转换为 int 数字
            'strings_to_urls': False,  # 自动识别超链接
            'constant_memory': False,  # 连续内存模式 (True 适用于大数据量输出)
            'default_format_properties': {                                           
                'font_name': '宋体',  # 字体. 默认值 "Arial"
                'font_size': 11,  # 字号. 默认值 11
                # 'bold': False,  # 字体加粗
                # #'num_format': '#,##0.00', #数字格式
                # #'border': 6,  # 单元格边框宽度. 默认值 0
                'align': 'left',  # 对齐方式
                'valign': 'vcenter',  # 垂直对齐方式
                'text_wrap': True,  # 单元格内是否自动换行
                },
            }
                                            )  
        #设置居右千分位数字格式
        self.cell_format = self.wb_output.add_format({'locked': False})
        self.cell_format.set_bottom(1)
        self.cell_format.set_top(1)
        self.cell_format.set_left(1)
        self.cell_format.set_right(1)
        self.cell_format.set_border_color("black")
        self.cell_format.set_align("right")
        self.cell_format.set_num_format('#,##0.00')

        #设置居中整数数字格式
        self.cell_format2 = self.wb_output.add_format({'locked': False})
        self.cell_format2.set_bottom(1)
        self.cell_format2.set_top(1)
        self.cell_format2.set_left(1)
        self.cell_format2.set_right(1)
        self.cell_format2.set_align("center")
        self.cell_format2.set_border_color("black")

        #设置居中千分位数字格式
        self.cell_format3 = self.wb_output.add_format({'locked': False})
        self.cell_format3.set_bottom(1)
        self.cell_format3.set_top(1)
        self.cell_format3.set_left(1)
        self.cell_format3.set_right(1)
        self.cell_format3.set_align("center")
        self.cell_format3.set_border_color("black")
        self.cell_format3.set_num_format('#,##0.00')

        #设置居中日期格式
        self.cell_format4 = self.wb_output.add_format({'locked': False})
        self.cell_format4.set_bottom(1)
        self.cell_format4.set_top(1)
        self.cell_format4.set_left(1)
        self.cell_format4.set_right(1)
        self.cell_format4.set_align("center")
        self.cell_format4.set_border_color("black")
        self.cell_format4.set_num_format('yyyy-mm-dd')
            
        #设置居右百分比数字格式
        self.cell_format5 = self.wb_output.add_format({'locked': False})
        self.cell_format5.set_bottom(1)
        self.cell_format5.set_top(1)
        self.cell_format5.set_left(1)
        self.cell_format5.set_right(1)
        self.cell_format5.set_border_color("black")
        self.cell_format5.set_align("right")
        self.cell_format5.set_num_format('0.00%')

        #设置居左整数数字格式
        self.cell_format6 = self.wb_output.add_format({'locked': False})
        self.cell_format6.set_bottom(1)
        self.cell_format6.set_top(1)
        self.cell_format6.set_left(1)
        self.cell_format6.set_right(1)
        self.cell_format6.set_align("left")
        self.cell_format6.set_border_color("black")
        
        # 设置部分单元格为不可编辑
        self.locked_format = self.wb_output.add_format({'locked': True})  # 创建一个锁定格式
        self.locked_format.set_bottom(1)
        self.locked_format.set_top(1)
        self.locked_format.set_left(1)
        self.locked_format.set_right(1)
        self.locked_format.set_align("center")
        self.locked_format.set_border_color("black")
        
    def save(self):
        self.wb_output.close() 
       
####################################################################################
#导入项目
def upload_excel(ui):
    #选择训练集
    file_Path,_  = QFileDialog.getOpenFileName(ui,"选择文件",r"./train_data/","(*.xls *.xlsx)" )
    return file_Path

#导入数据模型
class compare_date_model():
    def __init__(self, file_Path ,ui ,date_type):
        self.file_Path = file_Path
        self.initialize(file_Path,ui,date_type)
        #读取单位转化
        db = db_process()
        db.db_get()
        sql = 'select info,uuid from measurement_change order by id'
        db.db_load_selected(sql)
        self.Measurement_in = []
        self.Measurement_out = []
        for i in db.data_loaded:
            self.Measurement_in.append(i[0])
            self.Measurement_out.append(i[1])
        self.Measurement_tag_change()
        self.text_to_token("./bert-base-chinese")
        self.group_by_Measurement()
        
    #初始化
    def initialize(self,file_Path,ui,date_type):
        #将所有数据表合并成一个
        base_data_sheets_dict = pd.read_excel(file_Path, sheet_name=None)
        # 遍历所有工作表，获取项目名称并存储为列表
        base_data_sheet_name = [sheet_data.iloc[0, 0] for sheet_data in base_data_sheets_dict.values()]
        a = 0
        for sheet_name, df in base_data_sheets_dict.items():
            # 添加新列并赋值
            df['项目名称'] = base_data_sheet_name[a]
            a += 1
        base_data = pd.concat(base_data_sheets_dict.values(), ignore_index=True)
        base_data = base_data.iloc[int(ui.spinBox_base_line.value()):]
        self.all_data_list_0=base_data['项目名称'].iloc[:].tolist()
        if date_type == 0:
            # 筛选第一列类型为数值的行
            base_data = base_data[pd.to_numeric(base_data.iloc[:, int(ui.spinBox_base_1.value())], errors='coerce').notna()]
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_base_1.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_base_2.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_base_3.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_base_5.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_base_7.value())]])
            # 加载并处理数据
            self.all_data_list_1=base_data.iloc[:, int(ui.spinBox_base_1.value())].tolist()
            self.all_data_list_2=base_data.iloc[:, int(ui.spinBox_base_2.value())].tolist()
            self.all_data_list_3=base_data.iloc[:, int(ui.spinBox_base_3.value())]
            self.all_data_list_4=base_data.iloc[:, int(ui.spinBox_base_4.value())]
            self.all_data_list_5=base_data.iloc[:, int(ui.spinBox_base_5.value())].tolist()
            self.all_data_list_6=base_data.iloc[:, int(ui.spinBox_base_6.value())].tolist()
            self.all_data_list_7=base_data.iloc[:, int(ui.spinBox_base_7.value())].tolist()
        elif date_type == 1:
            base_data = base_data[pd.to_numeric(base_data.iloc[:, int(ui.spinBox_compare_1.value())], errors='coerce').notna()]
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_compare_1.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_compare_2.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_compare_3.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_compare_5.value())]])
            base_data = base_data.dropna(subset=[base_data.columns[int(ui.spinBox_compare_7.value())]])
            # 加载并处理数据
            self.all_data_list_1=base_data.iloc[:, int(ui.spinBox_compare_1.value())].tolist()
            self.all_data_list_2=base_data.iloc[:, int(ui.spinBox_compare_2.value())].tolist()
            self.all_data_list_3=base_data.iloc[:, int(ui.spinBox_compare_3.value())]
            self.all_data_list_4=base_data.iloc[:, int(ui.spinBox_compare_4.value())]
            self.all_data_list_5=base_data.iloc[:, int(ui.spinBox_compare_5.value())].tolist()
            self.all_data_list_6=base_data.iloc[:, int(ui.spinBox_compare_6.value())].tolist()
            self.all_data_list_7=base_data.iloc[:, int(ui.spinBox_compare_7.value())].tolist()

        self.all_data = base_data
        self.sheet_name = base_data_sheet_name
        #print(self.all_data_list_1.tolist())
        
    #分词转换
    def text_to_token(self,tokenizer):
        self.tokenizer = AutoTokenizer.from_pretrained(tokenizer) #用于对文本进行分词的分词器。
        self.vocab_size = self.tokenizer.vocab_size #分词器长度
        print('读取模型字典：'+tokenizer+'，字典大小：'+str(self.vocab_size))
        texts=self.all_data_list_3.fillna(' ').astype(str)+self.all_data_list_4.fillna(' ').astype(str)
        self.all_data_list_3 = self.all_data_list_3.tolist()
        self.all_data_list_4 = self.all_data_list_4.tolist()
        #print(texts)
        self.text_token = [] #转换后分词向量
        # 统计分词频率
        for text in texts:
            # 使用Counter统计元素出现次数
            element_count = Counter(self.tokenizer(text[:511]).input_ids[1:-1]) #文字过长则限制至512字
            sorted_element_count = [0 for i in range(self.vocab_size)]
            for element, count in element_count.items():
                sorted_element_count[element]=count
            self.text_token.append(sorted_element_count)
        print('分词转换完成')

    #转化单位
    def Measurement_tag_change(self):
        self.Measurement = []
        for tag in self.all_data_list_5:
            if tag in self.Measurement_in:
                i=self.Measurement_in.index(tag)#搜寻位置
                self.Measurement.append(self.Measurement_out[i])
            else:
                #找不到位置，返回末尾值
                self.Measurement.append(6)
        print('单位转换完成')
        return self.Measurement
         
    #按单位分组
    def group_by_Measurement(self):
        #记录单位分组ID
        self.Measurement_group = []
        for i in range(6):
            Measurement_group_list = []
            a = 0
            for j in self.Measurement:
                if j == i+1: Measurement_group_list.append(a)#记录位置
                a += 1
            self.Measurement_group.append(Measurement_group_list)
        #print(self.Measurement_group)
        #print(self.Measurement)
        
        #记录分词分组
        self.text_token_group = []
        for i in self.Measurement_group:
            text_token_group_list = []
            for j in i:
                text_token_group_list.append(self.text_token[j])
            self.text_token_group.append(text_token_group_list)
        #print(self.text_token_group)

#对比模型
class compare_result():
    def __init__(self, base_data ,compare_data ,ui):
        self.base_data = base_data
        self.compare_data = compare_data
        self.ui = ui
        self.base_id = [] #记录基准清单位置ID
        self.compare_id = [] #记录对比ID
        self.similarity_value = [] #记录相似概率
        
    #计算余弦相似度
    def count_Cosine_similarity(self):
        #计算最高前n项所在位置及其值
        def find_largest_elements(lst, n):
            sorted_list = sorted(enumerate(lst), key=lambda x: x[1], reverse=True)
            largest_elements = sorted_list[:n]
            return [element[0] for element in largest_elements],sorted(lst, reverse=True)[:n]
        #按单位类别遍历
        for i in range(6):
            if len(self.base_data.text_token_group[i])>1 and len(self.compare_data.text_token_group[i])>1 :
                similarity = cosine_similarity(self.base_data.text_token_group[i], self.compare_data.text_token_group[i])
                #print(similarity)
                a = 0
                for j in similarity:
                    #最高前n项在对比清单所在位置及其值
                    similarity_lagest,similarity_lagest_value = find_largest_elements(j, int(self.ui.spinBox_compare_BOQ.value()))
                    #print(similarity_lagest,similarity_lagest_value)
                    #返回在原始数据中ID
                    compare_id_list = []
                    for k in similarity_lagest:
                        compare_id_list.append(self.compare_data.Measurement_group[i][k])
                    #print(compare_id_list)
                    self.compare_id.append(compare_id_list)
                    self.base_id.append(self.base_data.Measurement_group[i][a])
                    self.similarity_value.append(similarity_lagest_value)
                    a += 1
        #print('base_id:',self.base_id)
        #print('compare_id:',self.compare_id)
        #print('similarity_value:',self.similarity_value)
        
    #单价对比
    def compare_price(self):
        #对比清单单价按基准单价原始排序
        #基准清单单价序列
        base_BOQ_price = [float(value) if isinstance(value, (int, float)) else 0 for value in self.base_data.all_data_list_7]
        #对比清单单价序列
        compare_BOQ_price = [float(value) if isinstance(value, (int, float)) else 0 for value in self.compare_data.all_data_list_7]
        #基准清单与对比清单按原始顺序ID
        self.compare_id_original = []
        #按基准清单原始顺序对应对比清单价格，记录相似概率
        self.compare_BOQ_price_original = [] 
        self.similarity_value_original = [] 
        for i in range(len(self.base_id)):
            compare_price_id = self.base_id.index(i)
            self.compare_id_original.append(self.compare_id[compare_price_id])
            BOQ_price_original_list = []
            for j in self.compare_id[compare_price_id]:
                BOQ_price_original_list.append(compare_BOQ_price[j])
            self.compare_BOQ_price_original.append(BOQ_price_original_list)
            self.similarity_value_original.append(self.similarity_value[compare_price_id])

        #计算对比单价
        self.compare_result = [[row[0], sum(row) / len(row)] for row in self.compare_BOQ_price_original]
        #print('compare_id:',self.compare_id)
        #print('compare_id_original:',self.compare_id_original)
        #print('compare_BOQ_price_original:',self.compare_BOQ_price_original)
        #print('base_BOQ_price:',base_BOQ_price)
        #print('compare_result:',self.compare_result)
        #print('similarity_value_original:',self.similarity_value_original)

#对比开始
def compare_BOQ_beign(file_Path_base,file_Path_compare,ui,compare_type=0):
    # 读取Excel文件
    base_data = compare_date_model(file_Path_base,ui,0)
    compare_data = compare_date_model(file_Path_compare ,ui,1)
    ui.label_state.setText('数据预处理完成！')
    #print(base_data.all_data,compare_data.all_data)
    #print(base_data.all_data_list_1,base_data.all_data_list_2,base_data.all_data_list_3,base_data.all_data_list_4,base_data.all_data_list_5,base_data.all_data_list_6,base_data.all_data_list_7)
    #对比数据
    compare_result_data = compare_result(base_data ,compare_data ,ui)
    compare_result_data.count_Cosine_similarity()
    compare_result_data.compare_price()
    ui.label_state.setText('数据对比完成！')
    
    #excel输出地址
    #save_path = './Compare_Result.xlsx'
    save_path,_ = QFileDialog.getSaveFileName(ui, "选择存储路径",r"./分析结果","(*.xlsx)")
    #excel输出
    def result_to_excel(base_data,compare_data,compare_result_data,save_path,compare_type,not_compare):
        #输出excel
        wb = excel_model(save_path)
        sheet_output = wb.wb_output.add_worksheet('分析结果')
        #表头
        title = ['基准项目单位工程','序号','项目编码','项目名称','项目特征','单位','工程量','基准项目单价','对比项目单价','基准项目合价','对比项目合价','单价差','合价差','价差百分比','相近概率']
        #设置列宽
        head_width = [12, 6, 15, 15, 45, 6, 12, 10, 10, 15, 15, 10, 12, 10, 8]
        a = 0
        for i in title:
            sheet_output.write(0,a,i,wb.cell_format2)
            sheet_output.set_column(a,a,head_width[a])
            a += 1
        for i in range(len(base_data.all_data_list_1)):
            sheet_output.write(i+1,0,base_data.all_data_list_0[i],wb.cell_format2)
            sheet_output.write(i+1,1,base_data.all_data_list_1[i],wb.cell_format2)
            sheet_output.write(i+1,2,base_data.all_data_list_2[i],wb.cell_format2)
            sheet_output.write(i+1,3,base_data.all_data_list_3[i],wb.cell_format2)
            sheet_output.write(i+1,4,base_data.all_data_list_4[i],wb.cell_format6)
            sheet_output.write(i+1,5,base_data.all_data_list_5[i],wb.cell_format2)
            try:Quantity = float(base_data.all_data_list_6[i])
            except:Quantity = 0
            else:
                if compare_result_data.similarity_value_original[i][0]*100 < not_compare:
                    Quantity = 0
            try:price1 = float(base_data.all_data_list_7[i])
            except:price1 = 0
            try:price2 = float(compare_result_data.compare_result[i][compare_type])
            except:price2 = 0
            sheet_output.write(i+1,6,Quantity,wb.cell_format)
            sheet_output.write(i+1,7,price1,wb.cell_format)
            sheet_output.write(i+1,8,price2,wb.cell_format)
            total1 = Quantity*price1
            total2 = Quantity*price2
            sheet_output.write(i+1,9,total1,wb.cell_format)
            sheet_output.write(i+1,10,total2,wb.cell_format)
            sheet_output.write(i+1,11,price1-price2,wb.cell_format)
            sheet_output.write(i+1,12,total1-total2,wb.cell_format)
            if total1 != 0:
                sheet_output.write(i+1,13,(total1-total2)/total1,wb.cell_format5)
            else:
                sheet_output.write(i+1,13,0,wb.cell_format5)
            sheet_output.write(i+1,14,compare_result_data.similarity_value_original[i][0],wb.cell_format5)
        
        sheet_output1 = wb.wb_output.add_worksheet('清单对比结果')
        #表头
        title = ['基准项目单位工程','序号','项目编码','项目名称','项目特征','单位','基准项目单价', \
            '相近概率','对比项目单位工程','序号','项目编码','项目名称','项目特征','单位','对比项目单价']
        #设置列宽
        head_width = [12, 6, 15, 15, 45, 6, 10, \
            8, 12, 6, 15, 15, 45, 6, 10]
        a = 0
        for i in title:
            sheet_output1.write(0,a,i,wb.cell_format2)
            sheet_output1.set_column(a,a,head_width[a])
            a += 1
        a,b = 0,0
        for i in base_data.all_data_list_1:
            sheet_output1.write(b+1,0,base_data.all_data_list_0[a],wb.cell_format2)
            sheet_output1.write(b+1,1,i,wb.cell_format2)
            sheet_output1.write(b+1,2,base_data.all_data_list_2[a],wb.cell_format2)
            sheet_output1.write(b+1,3,base_data.all_data_list_3[a],wb.cell_format2)
            sheet_output1.write(b+1,4,base_data.all_data_list_4[a],wb.cell_format6)
            sheet_output1.write(b+1,5,base_data.all_data_list_5[a],wb.cell_format2)
            sheet_output1.write(b+1,6,base_data.all_data_list_7[a],wb.cell_format)
            c = 0
            for j in compare_result_data.compare_id_original[a]:
                sheet_output1.write(b+1,7,compare_result_data.similarity_value_original[a][c],wb.cell_format5)
                sheet_output1.write(b+1,8,compare_data.all_data_list_0[j],wb.cell_format2)
                sheet_output1.write(b+1,9,compare_data.all_data_list_1[j],wb.cell_format2)
                sheet_output1.write(b+1,10,compare_data.all_data_list_2[j],wb.cell_format2)
                sheet_output1.write(b+1,11,compare_data.all_data_list_3[j],wb.cell_format2)
                sheet_output1.write(b+1,12,compare_data.all_data_list_4[j],wb.cell_format6)
                sheet_output1.write(b+1,13,compare_data.all_data_list_5[j],wb.cell_format2)
                sheet_output1.write(b+1,14,compare_data.all_data_list_7[j],wb.cell_format)
                b += 1
                c += 1
            a += 1

        wb.save()
    
        
    if len(save_path)>0:
        result_to_excel(base_data,compare_data,compare_result_data,save_path,compare_type,float(ui.spinBox_not_compare.value()))
        ui.label_state.setText('对比完成！结果输出至:'+save_path)
        
        QMessageBox.information(ui,'提示','对比完成！')
    else:
        QMessageBox.information(ui,'提示','未选择储存路径！')