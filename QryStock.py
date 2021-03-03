from time import sleep
from selenium import webdriver
import selenium
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import csv
import pathlib
from pandas import DataFrame
import PySimpleGUI as sg

class PyGui:
    TableList=''
    TableListHeading=''
    def __init__(self) -> None:
        pass
    def open_Table(self,Qry):
        tableDF = DataFrame()
        tableDF = Qry.crawlDataDF
        self.TableList = tableDF.values.tolist()
        self.TableListHeading = list(tableDF.head())
        return self.set_Table_Window()
        pass
    def set_Table_Window(self):
        Table_Window_Layout=[
            [sg.Table(self.TableList,headings=self.TableListHeading,num_rows=min(30,len(self.TableList)),k='Table', select_mode="extended", def_col_width=15, vertical_scroll_only=False, auto_size_columns=False)],
            [sg.Button('匯出'),sg.Button('關閉')]
        ]
        return sg.Window('結果清單',Table_Window_Layout,margins=(5,5),finalize=True,resizable=True)
    def set_StartUp_Window(self,Qry):
        startUp_Window_Layout=[
            [sg.Text('選擇週：'),sg.Combo(Qry.dateList,default_value=Qry.dateList[0],readonly=True,k='-Date-',size=(15,1))],
            [sg.Button('確定'),sg.Button('取消')]
        ]
        return sg.Window('選擇時間點',startUp_Window_Layout,margins=(20,5), finalize=True)
class QryStock:
    date_Element =''
    inputCoid_Element = ''
    sub_Coid_Element =''
    current_Date_Index =''
    current_Date =''
    driver=''
    coidList=[]
    current_coid=''
    dateList=[]
    crawlDataDF=[]
    current_Process=0
    total=0
    exist=0
    import_csv='.\上市&上櫃.csv'
    coidList_Dict={"股號":[],"股名":[],"千張持股變化":[],"本周千張持股":[],"上周千張持股":[]}
    coidList_Dict_Type={"股號":'string',"股名":'string',"千張持股變化":'double',"本周千張持股":'double',"上周千張持股":'double'}

    def export(self):
        filename = sg.popup_get_file('選擇儲存路徑','匯出表格',default_path=f'{self.current_Date} - 集保戶股權分散表 － 匯出',save_as=True,file_types=(("CSV 檔","*.csv"),("Excel 檔","*.xlsx")),no_window=True)
        if(pathlib.Path(filename).suffix==".csv"):
            self.crawlDataDF.to_csv(filename,encoding='utf-8', index=False)
        if(pathlib.Path(filename).suffix==".xlsx"):
            self.crawlDataDF.to_excel(filename,encoding='utf-8', index=False)
        pass
    def auto_Mode(self):  # 自動模式
        with open(self.import_csv, newline='', encoding="utf-8") as csvfile_Lc:  # 讀入CSV檔案
            rows = csv.DictReader(csvfile_Lc)
            for row in rows:
                self.total += 1
                if((len(row['代號']) == 4) and row['代號'].isnumeric()):  # 檢查股號是否為純號碼以及是否為4位數
                    self.exist+=1
                    Co_id = row['代號']
                    name = row['名稱']
                    self.coidList.append([Co_id,name])

    def __init__(self) -> None:
        self.crawlDataDF = DataFrame(self.coidList_Dict)
        url = 'https://www.tdcc.com.tw/smWeb/QryStock.jsp'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.get(url)
        wait = ui.WebDriverWait(self.driver,10)
        wait.until(lambda driver: driver.find_element_by_id(id_='scaDates'))
        self.date_Element = Select(self.driver.find_element_by_id(id_='scaDates'))
        self.inputCoid_Element = self.driver.find_element_by_id(id_='StockNo')
        pass
    
    def get_Date(self):
        self.date_Element = Select(self.driver.find_element_by_id(id_='scaDates'))
        for date in self.date_Element.options:
            self.dateList.append(date.text)
        
    
    def set_COID(self,coidString):
        sleep(.5)
        self.inputCoid_Element = self.driver.find_element_by_id(id_='StockNo')
        self.inputCoid_Element.click()
        try:
            self.inputCoid_Element.send_keys(coidString)
        except selenium.common.exceptions.UnexpectedAlertPresentException:
            print('跳出警告')
            #self.driver.switch_to.alert.accept()
            self.inputCoid_Element = self.driver.find_element_by_id(id_='StockNo')
            self.inputCoid_Element.click()
            self.inputCoid_Element.send_keys(coidString)
        self.inputCoid_Element.send_keys(Keys.ENTER)
        wait = ui.WebDriverWait(self.driver,4)
        try:
            wait.until(lambda driver: driver.find_element_by_name('radioStockNo'))
        except selenium.common.exceptions.TimeoutException:
            return False
        self.sub_Coid_Element = self.driver.find_element_by_name('radioStockNo')
        if self.sub_Coid_Element.is_selected():
            #print(f'已選擇股號：{self.sub_Coid_Element.get_attribute("VALUE")}')
            return True
        else:
            print(f'找不到{coidString}')
            return False

    def submit(self,isCurrentWeek):
        submit_btn = self.driver.find_element_by_name('sub')
        submit_btn.click()
        wait = ui.WebDriverWait(self.driver,3)
        try:
            wait.until(lambda driver: driver.find_element_by_xpath('//td[contains(text(),"1,000,001以上")]/following-sibling::td[3]'))
        except selenium.common.exceptions.TimeoutException:
            self.driver.refresh()
            self.set_COID(self.current_coid[0])
            if(isCurrentWeek):
                self.submitGetThisWeek()
            else:
                self.submitGetLastWeek()
        table_element = self.driver.find_element_by_xpath('//td[contains(text(),"1,000,001以上")]/following-sibling::td[3]')
        num = float(table_element.text)
        return round(num,2)

    def submitGetLastWeek(self):
        wait = ui.WebDriverWait(self.driver,10)
        wait.until(lambda driver: driver.find_element_by_id(id_='scaDates'))
        self.date_Element = Select(self.driver.find_element_by_id(id_='scaDates'))
        try:
            self.date_Element.select_by_index(self.current_Date_Index+1)
        except selenium.common.exceptions.NoSuchElementException:
            self.driver.refresh()
            self.set_COID(self.current_coid[0])
            self.submitGetLastWeek()
        return self.submit(False)

    def submitGetThisWeek(self):
        wait = ui.WebDriverWait(self.driver,10)
        wait.until(lambda driver: driver.find_element_by_id(id_='scaDates'))
        self.date_Element = Select(self.driver.find_element_by_id(id_='scaDates'))
        try:
            self.date_Element.select_by_index(self.current_Date_Index)
        except selenium.common.exceptions.NoSuchElementException:
            self.driver.refresh()
            self.set_COID(self.current_coid[0])
            self.submitGetThisWeek()
        return self.submit(True)

    def q_Sumbit(self,date):
        self.current_Date_Index=self.dateList.index(date)
        self.current_Date = date
        for coid in self.coidList:
            self.current_coid=coid
            sg.one_line_progress_meter('爬取資料',self.current_Process,self.exist-1,key='Process',orientation='h')
            if(self.set_COID(self.current_coid[0])):
                print(f'正在抓取股號：{coid}的資料')
                currentWeek = self.submitGetThisWeek()
                if(self.set_COID(coid[0])):
                    lastWeek = self.submitGetLastWeek()
                else:
                    self.current_Process+=1
                    continue
                numChange=currentWeek-lastWeek
                numChange=round(numChange,2)
                dict_add={"股號":str(self.coidList[self.coidList.index(coid)][0]),"股名":str(self.coidList[self.coidList.index(coid)][1]),"抓取週":self.current_Date,"千張持股變化":numChange,"本周千張持股":currentWeek,"上周千張持股":lastWeek}
                print(dict_add)
                cols=["股號","股名","抓取週","千張持股變化","本周千張持股","上周千張持股"]
                self.crawlDataDF = self.crawlDataDF.append(dict_add, ignore_index=True)
                self.crawlDataDF = self.crawlDataDF[cols]
                self.current_Process+=1
            else:
                self.current_Process+=1
                continue
        
        pass
Qry = QryStock()
Qry.get_Date()
Pygui = PyGui()
main_Window = Pygui.set_StartUp_Window(Qry)
table_Window = None
while True:
    window,event,values = sg.read_all_windows()
    if window == main_Window:
        if event == '確定':
            print('開始抓取',values['-Date-'],'的資料')
            #print(len(Qry.dateList),':',Qry.dateList.index(values['-Date-']))
            if(len(Qry.dateList)-1!=Qry.dateList.index(values['-Date-'])):
                Qry.auto_Mode()
                Qry.q_Sumbit(values['-Date-'])
                main_Window.close()
                table_Window = Pygui.open_Table(Qry)
            else:
                sg.popup_error('選擇週為最早週，請選擇其他週！','範圍超過')
        if event in ('取消',sg.WIN_CLOSED):
            break
    if window == table_Window:
        if event == '匯出':
            Qry.export()
        if event in ('關閉',sg.WIN_CLOSED):
            table_Window.close()
            break

main_Window.close()
#Qry.auto_Mode()
#Qry.q_Sumbit()
print(Qry.crawlDataDF)