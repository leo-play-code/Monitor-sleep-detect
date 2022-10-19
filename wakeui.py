import sys
from gspread.models import Worksheet
import pandas as pd
import os
import keyboard  
import csv
import time
import pyautogui
from pathlib import Path
import subprocess
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from functools import partial
import gspread
from google.oauth2.service_account import Credentials
import datetime as dt
from datetime import date as dt3, timedelta
from datetime import datetime as dt2
python_path = os.getcwd()
def get_sheet_list():
    try:
        auth_json_path = python_path+'/res/ddm-test-answer-2021-ebe1485ce81b.json'
        gss_scopes = ['https://spreadsheets.google.com/feeds']
        #連線
        credentials = Credentials.from_service_account_file(auth_json_path,scopes=gss_scopes)
        gss_client = gspread.authorize(credentials)
        #開啟 Google Sheet 資料表
        spreadsheet_key = '12RTFxb0UQ45VwPdiljt1fdc9fanzCW50dwFJRuk2fWA' 
        sheet = gss_client.open_by_key(spreadsheet_key)
        worksheet_list = sheet.worksheets()
        work=f'{worksheet_list}'
        work=list(work.split("'"))
        i=1
        worklist=[]
        while i<len(work):
            worklist.append(work[i])
            i+=2
        return worklist
    except:
        print('沒網路,請連網！！')
def get_monitor_list():
    result = subprocess.run(['system_profiler', 'SPDisplaysDataType'], stdout=subprocess.PIPE)
    data=result.stdout
    print(data)
    list_data=str(data).replace("b'","").replace("'","").split("\\n")
    print(' ')
    print(list_data)
    new_data=[]
    for item in list_data:
        item = item.replace("  ","")
        if item!='':
            new_data.append(item) 
    print(' ')
    print(new_data)
    i=0
    find_display=[]
    while i<len(new_data):
        if 'Displays:' in new_data[i] :
            print(new_data[i+1])
            monitor=new_data[i+1].replace(":","")
            find_display.append(monitor)
        i+=1
    return find_display
class wakeui(QMainWindow):
    def __init__(self):
        self.x=800
        self.y=800
        super(wakeui, self).__init__()
        self.setFixedSize(self.x,self.y)
        # 設置窗口標題
        self.setWindowTitle('LEO-auto-sleep')
        #應用的初始調色板
        self.origPalette = QApplication.palette()
        #外部輸入
        self.worklist=get_sheet_list()
        self.test_mode=['close screan','sleep/wake up']
        # 初始設定
        self.label_title_setup = QLabel("基本資料設定：",self)
        self.label_title_setup.setGeometry(self.x*0.1, self.y*0.1,self.x*0.15, 30)
        self.start_ornot=False
        #密碼輸入
        self.label_password = QLabel("輸入密碼:",self)
        self.label_password.setGeometry(self.x*0.03, self.y*0.15,self.x*0.08, 30)
        self.set_password = QLineEdit('',self)
        self.set_password.setPlaceholderText("密碼輸入：")
        self.set_password.setGeometry(self.x*0.13, self.y*0.15, self.x*0.15, 30)
        self.set_password.textChanged.connect(lambda: self.save_password())
        #選取要測試的螢幕
        self.label_monitor = QLabel("選擇螢幕:",self)
        self.label_monitor.setGeometry(self.x*0.03, self.y*0.20, self.x*0.08, 30)
        self.monitor = QComboBox(self)
        self.monitor.show()
        self.monitor.addItems(get_monitor_list())
        self.monitor.currentIndexChanged.connect(self.save_monitor)
        self.monitor.setGeometry(self.x*0.14, self.y*0.20, self.x*0.15, 30)
        #選取要測試的模式
        self.label_mode = QLabel("選擇模式:",self)
        self.label_mode.setGeometry(self.x*0.03, self.y*0.25, self.x*0.08, 30)
        self.mode = QComboBox(self)
        self.mode.show()
        self.mode.addItems(self.test_mode)
        self.mode.currentIndexChanged.connect(self.save_mode)
        self.mode.setGeometry(self.x*0.14, self.y*0.25, self.x*0.15, 30)
        # 參數設定
        self.label_state_setup = QLabel("參數設定:",self)
        self.label_state_setup.setGeometry(self.x*0.1, self.y*0.35,self.x*0.08, 30)
        #設定時間
        self.label_times = QLabel("輸入時間:",self)
        self.label_times.setGeometry(self.x*0.03, self.y*0.40,self.x*0.08, 30)
        self.set_times = QLineEdit('',self)
        self.set_times.setPlaceholderText("時間輸入：")
        self.set_times.setGeometry(self.x*0.13, self.y*0.40, self.x*0.15, 30)
        self.set_times.textChanged.connect(lambda: self.save_times())
        #設定次數
        self.label_count = QLabel("輸入次數:",self)
        self.label_count.setGeometry(self.x*0.03, self.y*0.45,self.x*0.08, 30)
        self.set_count = QLineEdit('',self)
        self.set_count.setPlaceholderText("次數輸入：")
        self.set_count.setGeometry(self.x*0.13, self.y*0.45, self.x*0.15, 30)
        self.set_count.textChanged.connect(lambda: self.save_count())
        # sheet設定
        self.label_sheet_setup = QLabel("sheet設定:",self)
        self.label_sheet_setup.setGeometry(self.x*0.1, self.y*0.55,self.x*0.08, 30)
        #選擇要跑的sheet
        self.label_sheet = QLabel("選擇sheet:",self)
        self.label_sheet.setGeometry(self.x*0.03, self.y*0.60, self.x*0.08, 30)
        self.sheet = QComboBox(self)
        self.sheet.show()
        self.sheet.addItems(get_sheet_list())
        self.sheet.currentIndexChanged.connect(self.save_sheet)
        self.sheet.setGeometry(self.x*0.14, self.y*0.60, self.x*0.15, 30)
        #創造新的sheet
        self.label_sheet = QLabel("建立新sheet:",self)
        self.label_sheet.setGeometry(self.x*0.03, self.y*0.65, self.x*0.1, 30)
        self.set_new_sheet = QLineEdit('',self)
        self.set_new_sheet.setPlaceholderText("輸入新sheet：")
        self.set_new_sheet.setGeometry(self.x*0.14, self.y*0.65, self.x*0.15, 30)
        self.set_new_sheet.textChanged.connect(lambda: self.new_sheet_name())
        self.create_sheet=QPushButton('建立', self)
        self.create_sheet.setGeometry(self.x*0.30,self.y*0.65,self.x*0.08, 30)
        self.create_sheet.clicked.connect(self.create_new_sheet)
        # 搜尋
        self.savebutton=QPushButton('執行', self)
        self.savebutton.setGeometry(self.x*0.05,self.y*0.75, 200, 80)
        self.savebutton.clicked.connect(self.search)  
        # 暫停button
        self.stopbutton=QPushButton('暫停', self)
        self.stopbutton.setStyleSheet("background-color: rgb(255,106,106)")
        self.stopbutton.setGeometry(self.x*0.05,self.y*0.75, 200, 80)
        self.stopbutton.clicked.connect(self.stop)  
        self.change_mode()
        # 刷新狀態
        self.refreshbutton=QPushButton('刷新', self)
        self.refreshbutton.setStyleSheet("background-color: rgb(255,215,0)")
        self.refreshbutton.setGeometry(self.x*0.085,self.y*0.03,100, 40)
        self.refreshbutton.clicked.connect(self.refresh)  
        self.change_mode()
        # 顯示
        col_lst = ['時間','次數','螢幕是否喚醒','測試模式']
        self.MyTable = QTableWidget(100,4,self)
        #設定字型、表頭   
        self.MyTable.setHorizontalHeaderLabels(col_lst)
        #設定豎直方向表頭不可見
        self.MyTable.verticalHeader().setVisible(False)
        self.MyTable.setFrameShape(QFrame.NoFrame)
        self.MyTable.setGeometry(self.x*0.38,self.y*0.03, self.x*0.6, self.y*0.9)
        self.MyTable.clearContents()
        self.MyTable.setStyleSheet("color: black;")
    def save_password(self):
        data=self.set_password.text()
        return data
    def save_monitor(self):
        data=self.monitor.currentText()
        return data
    def save_mode(self):
        data=self.mode.currentText()
        return data
    def save_times(self):
        data=self.set_times.text()
        return data
    def save_count(self):
        data=self.set_count.text()
        return data
    def save_sheet(self):
        data=self.sheet.currentText()
        return data
    def new_sheet_name(self):
        data=self.set_new_sheet.text()
        return data
    # def stop(self):
    #     self.search().stop()
    def change_mode(self):
        if self.start_ornot==False:
            self.savebutton.show()
            self.stopbutton.hide()
        else:
            self.savebutton.hide()
            self.stopbutton.show()
        QApplication.processEvents()
    def stop(self):
        self.start_ornot=False
    def refresh(self):
        # 刷新sheet
        self.worklist=get_sheet_list()
        self.label_sheet = QLabel("選擇sheet:",self)
        self.label_sheet.setGeometry(self.x*0.03, self.y*0.60, self.x*0.08, 30)
        self.sheet = QComboBox(self)
        self.sheet.show()
        self.sheet.addItems(get_sheet_list())
        self.sheet.currentIndexChanged.connect(self.save_sheet)
        self.sheet.setGeometry(self.x*0.14, self.y*0.60, self.x*0.15, 30)
        # 刷新螢幕
        self.label_monitor = QLabel("選擇螢幕:",self)
        self.label_monitor.setGeometry(self.x*0.03, self.y*0.20, self.x*0.08, 30)
        self.monitor = QComboBox(self)
        self.monitor.show()
        self.monitor.addItems(get_monitor_list())
        self.monitor.currentIndexChanged.connect(self.save_monitor)
        self.monitor.setGeometry(self.x*0.14, self.y*0.20, self.x*0.15, 30)
        QApplication.processEvents()
        
    def search(self):
        self.start_ornot=True
        self.change_mode()
        self.MyTable.clearContents()
        mode=self.save_mode()
        password=self.save_password()
        time_set=self.save_times()
        count=self.save_count()
        try:
            int(time_set)
        except:
            time_set=1
        try:
            int(count)
        except:
            count=1
        
        try:
            if password !='#-#':
                password=password
        except:
            password='#-#'
        break_while = False
        os.system('open -a Terminal')
        sheet=self.create_title()
        # try:
        self.i=0
        while self.i<int(count):
            
            if mode=='close screan':
                self.sleep_mon(int(time_set),str(password))
            elif mode=='sleep/wake up':
                self.close_com(int(time_set),str(password))
            if self.start_ornot==False:
                self.change_mode()
                break
            data_x = pd.read_csv(python_path+'/res/output.csv').columns.tolist()
            vol_1 = len(data_x)
            data_x= pd.read_csv(python_path+'/res/output.csv')
            row_4 = len(data_x)
            self.MyTable.clearContents()
            #查詢到的更新帶表格當中
            if sheet==False:
                pass
            else:
                self.input_googlesheet(sheet,data_x.時間[row_4-1],data_x.次數[row_4-1],data_x.螢幕是否喚醒[row_4-1],mode)
            for i_x in range(row_4):
                for j_y in range(vol_1):
                    print(j_y,vol_1)
                    if j_y==0:
                        temp_data_1 = data_x.時間[i_x]  # 臨時記錄，不能直接插入表格
                        print(temp_data_1)
                        data_1 = QTableWidgetItem(str(temp_data_1))  # 轉換後可插入表格
                        self.MyTable.setItem(i_x, j_y, data_1)
                    elif j_y==1:
                        temp_data_1 = data_x.次數[i_x]  # 臨時記錄，不能直接插入表格
                        print(temp_data_1)
                        data_1 = QTableWidgetItem(str(temp_data_1))  # 轉換後可插入表格
                        self.MyTable.setItem(i_x, j_y, data_1)
                    elif j_y==2:
                        temp_data_1 = data_x.螢幕是否喚醒[i_x]  # 臨時記錄，不能直接插入表格
                        print(temp_data_1)
                        data_1 = QTableWidgetItem(str(temp_data_1))  # 轉換後可插入表格
                        self.MyTable.setItem(i_x, j_y, data_1)
                        if '成功' in data_x.螢幕是否喚醒[i_x]:
                            self.MyTable.item(i_x, 2).setBackground(QColor(0 ,255 ,127))
                        elif '未連接' in data_x.螢幕是否喚醒[i_x]:
                            self.MyTable.item(i_x, 2).setBackground(QColor(255,236,139))
                            break_while = True
                        else:
                            self.MyTable.item(i_x, 2).setBackground(QColor(238,99,99))
                            break_while = True
                    elif j_y==3:
                        temp_data_1 = mode  # 臨時記錄，不能直接插入表格
                        print(temp_data_1)
                        data_1 = QTableWidgetItem(str(temp_data_1))  # 轉換後可插入表格
                        self.MyTable.setItem(i_x, j_y, data_1)
                    self.MyTable.resizeColumnsToContents()
                    QApplication.processEvents()
                self.MyTable.resizeColumnsToContents()
                QApplication.processEvents()
            if break_while == True:
                break_while = False
                break
            self.i+=1
            # except KeyboardInterrupt:
            #     pass
    def create_new_sheet(self):
        # 儲存就可以直接搜尋ㄌ
        if self.new_sheet_name() not in self.worklist:
            self.sheet = QComboBox(self)
            self.sheet.show()
            new_list=self.worklist
            print(new_list)
            print(self.new_sheet_name())
            new_list.insert(0,self.new_sheet_name())
            print(new_list)
            self.sheet.addItems(new_list)
            self.sheet.currentIndexChanged.connect(self.save_sheet)
            self.sheet.setGeometry(self.x*0.14, self.y*0.60, self.x*0.15, 30)
            QApplication.processEvents()
    # 建立sheet title
    def create_title(self):
        sheetcanbecreate=0
        # try:
        auth_json_path =python_path+'/res/ddm-test-answer-2021-ebe1485ce81b.json'
        gss_scopes = ['https://spreadsheets.google.com/feeds']
        #連線
        #credentials = ServiceAccountCredentials.from_json_keyfile_name(auth_json_path,gss_scopes)
        credentials = Credentials.from_service_account_file(auth_json_path,scopes=gss_scopes)
        gss_client = gspread.authorize(credentials)
        #開啟 Google Sheet 資料表
        spreadsheet_key = '12RTFxb0UQ45VwPdiljt1fdc9fanzCW50dwFJRuk2fWA' 
        #建立工作表1
        #sheet = gss_client.open_by_key(spreadsheet_key).sheet1
        #自定義工作表名稱
        ddtimestart=f"{dt2.now().strftime('%Y%m%d%H%M%S')}"
        sheet=self.save_sheet()
        try:
            googlename=sheet
        except:
            googlename=ddtimestart
            sheet = gss_client.open_by_key(spreadsheet_key).add_worksheet(googlename, 1000, 4, index=None)
            values =['時間','次數','螢幕是否喚醒','測試模式']
            sheet.insert_row(values, 1) #插入values到第1列
            sheetcanbecreate+=1
        if sheetcanbecreate==0:
            try:
                sheet = gss_client.open_by_key(spreadsheet_key).add_worksheet(googlename, 1000, 4, index=None)
                values =['時間','次數','螢幕是否喚醒','測試模式']
                sheet.insert_row(values, 1) #插入values到第1列
            except:
                sheet = gss_client.open_by_key(spreadsheet_key).worksheet(googlename)
                values =['時間','次數','螢幕是否喚醒','測試模式']
                sheet.insert_row(values, len(sheet.get_all_values())+1) #插入values到第1列
        return sheet
        # # except:
            # print('網路掛了,建立sheet失敗')
            # return False
    # 測試流程
    def typing(self,item):
        time.sleep(2)
        keyboard.write(item)

    def check_monitor(self,time_now_dt):
        result = subprocess.run(['system_profiler', 'SPDisplaysDataType'], stdout=subprocess.PIPE)
        if self.save_monitor() in str(result.stdout):
            if 'Display Asleep: Yes' in str(result.stdout):
                with open(python_path+'/res'+"/output.csv",'a') as fd:
                    writer = csv.writer(fd)
                    time_now_dt_temp=dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    writer.writerow([time_now_dt_temp,self.i+1,'失敗',self.save_mode()])
                # print('monitor close')
            else:
                with open(python_path+'/res'+"/output.csv",'a') as fd:
                    writer = csv.writer(fd)
                    time_now_dt_temp=dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    writer.writerow([time_now_dt_temp,self.i+1,'成功喚醒',self.save_mode()])
        else:
            with open(python_path+'/res'+"/output.csv",'a') as fd:
                writer = csv.writer(fd)
                time_now_dt_temp=dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                writer.writerow([time_now_dt_temp,self.i+1,'未連接到此螢幕',self.save_mode()])
        while dt.datetime.now()<time_now_dt:
            pass
    def sleep_mon(self,time_set,password):
        self.typing('sudo -s')
        keyboard.press_and_release('enter')
        if password !='#-#':
            self.typing(password)
            keyboard.press_and_release('enter')
        time_now_dt=dt.datetime.now()+timedelta(minutes=time_set)
        time_now=time_now_dt.strftime('%m/%d/%Y %H:%M:%S')
        print(time_now)
        self.typing('pmset schedule wake '+'"'+time_now+'"')
        keyboard.press_and_release('enter')
        self.typing('pmset displaysleepnow')
        keyboard.press_and_release('enter')
        time.sleep(5)
        # check_monitor(time_now_dt,i,python_path)
        while dt.datetime.now()<time_now_dt:
            if self.start_ornot==False:
                self.change_mode()
                QApplication.processEvents()
                return
            QApplication.processEvents()
        time.sleep(5)
        pyautogui.click(button = 'left')
        if password !='#-#':
            self.typing(password)
            keyboard.press_and_release('enter')
        time.sleep(5)
        self.check_monitor(time_now_dt)
    def close_com(self,time_set,password):
        self.typing('sudo -s')
        keyboard.press_and_release('enter')
        if password !='#-#':
            self.typing(password)
            keyboard.press_and_release('enter')
        time_now_dt=dt.datetime.now()+timedelta(minutes=time_set)
        time_sleep=dt.datetime.now()+timedelta(seconds=10)
        time_sleep=time_sleep.strftime('%m/%d/%Y %H:%M:%S')
        time_now=time_now_dt.strftime('%m/%d/%Y %H:%M:%S')
        print(time_now)
        self.typing('pmset schedule wake '+'"'+time_now+'"')
        keyboard.press_and_release('enter')
        self.typing('pmset sleepnow')
        keyboard.press_and_release('enter')
        time.sleep(5)
        # check_monitor(time_now_dt,i,python_path)
        while dt.datetime.now()<time_now_dt:
            if self.start_ornot==False:
                self.change_mode()
                QApplication.processEvents()
                return
            QApplication.processEvents()
        time.sleep(5)
        pyautogui.click(button = 'left')
        if password !='#-#':
            self.typing(password)
            keyboard.press_and_release('enter')
        time.sleep(5)
        self.check_monitor(time_now_dt)
    def input_googlesheet(self,sheet,time_set,count,state,mode):
        try:
            values = [str(time_set),str(count),str(state),str(mode)]
            sheet.insert_row(values, len(sheet.get_all_values())+1)
        except:
            print('網路掛了') 
    # except:
    #     print('無法輸入到google sheet')
            
if __name__ == '__main__':
    file = Path(python_path+'/res'+"/output.csv")
    if file.exists():
        os.remove(python_path+'/res'+"/output.csv")
    with open(python_path+'/res'+"/output.csv",'a') as fd:
        writer = csv.writer(fd)
        writer.writerow(['時間','次數','螢幕是否喚醒','模式'])
    app = QApplication(sys.argv)
    window = wakeui()
    #設置樣式表
    # app.setStyleSheet(qdarkstyle.load_stylesheet())
    window.show()
    sys.exit(app.exec())