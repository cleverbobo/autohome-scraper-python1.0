import xlrd
import urllib.parse
import urllib.request
from bs4 import BeautifulSoup
from selenium import webdriver
# 导入界面设计文件
from tkinter.constants import FALSE, TRUE
from PySide2.QtWidgets import QApplication, QMessageBox
from PySide2.QtUiTools import QUiLoader
from PySide2.QtGui import  QIcon
# 导入线程库
from threading import Thread
# 文件操作库
from tkinter.filedialog import askopenfilename
# excel表格操作
import openpyxl
class status:
    def __init__(self):
        self.dirname=''
        self.driver=''
        self.ui = QUiLoader().load('get_car_data.ui')
        self.ui.select_driver.clicked.connect(self.select_driver)
        self.ui.select_file.clicked.connect(self.get_dirname)
        self.ui.start.clicked.connect(self.start)

    def start(self):
        QMessageBox.information(self.ui,'操作成功','请耐心等待！期间不要关闭任何弹窗，最小化即可！')
        self.ui.start.setEnabled(False)
        # 获取汽车模型具体名称,当出现‘全新’，‘新’字时将其删除
        self.model_name_list = self.read_data(self.dirname)
        # 每个模型的搜索结果
        self.model_number=[]
        # 将汽车模型进行urlencode
        search_url_list=self.decode(self.model_name_list)
        # 将模型名称作为关键字进行搜索
        self.invalid_models_list=[]
        key_numbers=self.get_url_number(search_url_list)
        # print(self.invalid_models_list)
        # 线程个数
        self.thread_num = 1
        # 获取模型配置页的url
        config_url_list = self.fund(self.config_url(key_numbers),self.thread_num)
        # 抓取配置页中的关键数据，并存入列表
        model_index = [0]
        for i,number in enumerate(self.model_number):
            model_index.append(model_index[i]+number)
        model_index.pop(0)
        # 生成为找到模型的txt
        file=open('./result/not_found_list.txt','w',encoding='gbk')
        for name in self.invalid_models_list:
            file.write(name+'\n')
        file.close()
        # 调用多个线程同时抓取
        keys=[[] for i in range(self.thread_num)]
        values=[[] for i in range(self.thread_num)]
        threads = [None] * self.thread_num
        for i in range(self.thread_num):
                threads[i] = Thread(target=self.get_data, args=(config_url_list[i],keys,values,i))
                threads[i].start() # 开始线程
        for i in range(self.thread_num):
            threads[i].join() 
        

        # 数据清洗
        require_list=('车型名称','厂商指导价(元)','能源类型','环保标准','排量(L)',\
                      '变速箱类型','最大马力(Ps)','长度(mm)','宽度(mm)','车门数(个)',\
                      '工信部综合油耗(L/100km)','轴距(mm)','整备质量(kg)','气缸数(个)',\
                      '驱动方式','助力类型','空调温度控制方式','上市时间','高度(mm)')
        real_values=[]
        # 匹配数据，进行清洗
        for k,v in zip(keys,values):
            for num in range(len(k)):
                tool=[]
                for key in require_list:  
                    for i,data in zip(k[num],v[num]):
                        if i==key:
                            tool.append(data)
                            break
                real_values.append(tool)
        con = 0
        for i in range(len(real_values)):
            if(i<model_index[con]):
                real_values[i].insert(0,self.model_name_list[con])
            else:
                con += 1
                real_values[i].insert(0,self.model_name_list[con])
        #写入excel文档
        result = openpyxl.load_workbook('result.xlsx')
        sheet1 = result['data']
        for row in real_values:
            # 添加到下一行的数据
            sheet1.append(row)
        result.save('./result/result_test.xlsx')
        QMessageBox.information(self.ui,'抓取完成','请查看您的文件！')
        self.ui.start.setEnabled(True)

    def select_driver(self):
        self.driver = askopenfilename()

    # 分割url_list，用于放到不同的线程中
    def fund(self,listTemp, n):
        results = []
        length = int(len(listTemp)/n)
        for i in range(n):
            temp = listTemp[i:i + length]
            results.append(temp)
        return results
    
    # 获取文件路径
    def get_dirname(self):
        self.dirname = askopenfilename()
        self.ui.dir_name.append(self.dirname+'\n')
    
    # 读取文件，获取模型名称
    def read_data(self,dirname):
        car_data = xlrd.open_workbook(dirname)
        data = car_data.sheet_by_index(0)
        # 找到具体的名字
        car_model = data.col_values(colx=1)
        car_model.pop(0)
        for i in range(len(car_model)):
            if '全新' in car_model[i]:
                car_model[i]=car_model[i].replace('全新','')
            elif '新' in car_model[i]:
                car_model[i]=car_model[i].replace('新','')
            else:
                continue
        return car_model  
    
    # 对模型名称进行编码
    def decode(self,name_list):
        decode_name_list = []
        for name in name_list:
            encodedUrl = name.encode('gb2312')
            decodedUrl = urllib.parse.quote(encodedUrl)
            decode_name_list.append('https://www.che168.com/china/list/?kw='+decodedUrl)
        return decode_name_list
   
    # 获取url中的关键数字
    def get_url_number(self,url_list):
        all_key_number=[]
        key = '/dealer/'
        for url in url_list:
            key_url = []
            key_url_number = []
            all_url = [] 
            html = urllib.request.urlopen(url).read().decode("gbk")
            soup = BeautifulSoup(html, features='html.parser')
            tags = soup.find_all('a')
            for tag in tags:
                all_url.append(str(tag.get('href')).strip())
            for url in all_url:
                if key in url:
                    key_url.append(url)
            if len(key_url) != 0:
                for number in key_url:
                    key_url_number.append(int(number.split('/')[3][0:8]))
                all_key_number.append(key_url_number)
            elif(len(key_url)) == 0: 
                all_key_number.append([0,])
        # 做一些数据处理
        # 配置无效的模型列表
        index=[i for i,key in enumerate(all_key_number) if key == [0,]]
        for i in index:
            self.invalid_models_list.append(self.model_name_list[i])
        for name in self.invalid_models_list:
            self.model_name_list.remove(name)
        # 获取每个模型的有效搜索数量
        for key_number in all_key_number:
            if key_number != [0,]:
                self.model_number.append(len(key_number))

        return all_key_number
   
    # 生成配置页url链接
    def config_url(self,url_number):
        keyword = 'https://www.che168.com/CarConfig/CarConfig.html?infoid='
        config_url = []
        for i in url_number:
            if i !=[0]: 
                for j in i:
                    config_url.append(keyword+str(j))
        return config_url
   
    # 获取数据
    def get_data(self,config_url_list,all_keys,all_values,index):
        # driver = webdriver.Chrome(self.driver)
        driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')
        driver.set_page_load_timeout(10)
        for url in config_url_list:
            driver.get(url)
            keys = driver.find_elements_by_class_name('table-left') 
            values = driver.find_elements_by_class_name('table-right') 
            tool_keys=[]
            tool_values=[]
            for value in values:
                tool_values.append(value.get_attribute('innerHTML'))
            for key in keys:
                tool_keys.append(key.get_attribute('innerHTML'))
            all_keys[index].append(tool_keys)
            all_values[index].append(tool_values)               
        # return all_keys,all_values


app = QApplication([])
app.setWindowIcon(QIcon('logo.png'))
test = status()
test.ui.show()
app.exec_()
test.update_flag=False


