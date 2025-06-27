from PyQt6 import QtWidgets,QtGui
from PyQt6.QtCore import Qt
import sys
import ico
from ExcelViewerTool import ExcelViewer
from CustomWidget import ExpandableItem,CalendarItem
import datetime as dt
import calendar

#主視窗
class MyWidget(QtWidgets.QWidget):
    #初始化設定
    def __init__(self):
        super().__init__()
        self.mainLayout=QtWidgets.QVBoxLayout(self)
        self.mainLayout.setContentsMargins(0,0,0,0)
        self.mainLayout.setSpacing(10)
        
        self.ER=ExcelViewer()
        
        self.setBasicWindow()   #設定基本視窗資訊
        self.setMenuBar()       #設定視窗選單
        self.setFileInputUI()   #設定檔案輸入的UI
        self.setFileResultUI()  #設定結果輸出UI
        
    #設定基本視窗資訊
    def setBasicWindow(self):
        self.setWindowTitle("AccessExcelViewer")
        self.resize(800,600)

        pixmap=QtGui.QPixmap()
        pixmap.loadFromData(ico.getIcon_bin())
        icon=QtGui.QIcon(pixmap)
        self.setWindowIcon(icon)
    
    #設定視窗選單
    def setMenuBar(self):
        self.MenuBar=QtWidgets.QMenuBar(self)

        self.menu_File=QtWidgets.QMenu('檔案')
        self.menu_content=QtWidgets.QMenu('內容')

        self.action_open=QtGui.QAction('開啟檔案')
        self.action_open.triggered.connect(self.openFile_Dialog)
        self.action_exit=QtGui.QAction('離開')
        self.action_exit.triggered.connect(sys.exit)
        self.menu_File.addAction(self.action_open)
        self.menu_File.addAction(self.action_exit)
        
        self.action_reset=QtGui.QAction('重置')
        self.action_reset.triggered.connect(self.reset)
        self.menu_content.addAction(self.action_reset)

        self.MenuBar.addMenu(self.menu_File)
        self.MenuBar.addMenu(self.menu_content)

        self.mainLayout.addWidget(self.MenuBar)

    #設定檔案輸入的UI
    def setFileInputUI(self):        
        girdBox=QtWidgets.QWidget(self)
        gird=QtWidgets.QGridLayout(girdBox)
        gird.setContentsMargins(10, 0, 10,0)
        gird.setSpacing(0)

        path_lab=QtWidgets.QLabel(parent=girdBox)       # 建立'檔案路徑'標籤
        path_lab.setText("檔案路徑:")
        path_lab.setSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum,QtWidgets.QSizePolicy.Policy.Preferred)
        gird.addWidget(path_lab, 0, 0)

        filename_lab=QtWidgets.QLabel(parent=girdBox)   # 建立'搜尋名稱'標籤
        filename_lab.setText("搜尋名稱:")
        filename_lab.setSizePolicy(QtWidgets.QSizePolicy.Policy.Maximum,QtWidgets.QSizePolicy.Policy.Preferred)
        gird.addWidget(filename_lab, 1, 0)

        self.path_input=QtWidgets.QLineEdit(parent=girdBox) # 建立'檔案路徑'單行輸入框
        self.path_input.returnPressed.connect(self.openFile)
        gird.addWidget(self.path_input,0,1)

        self.searchName_cmb = QtWidgets.QComboBox(parent=girdBox) # 建立'下拉選單'搜尋名稱
        self.searchName_cmb.setEnabled(False)
        self.searchName_cmb.activated.connect(self.setDetailPage)
        self.searchName_cmb.activated.connect(self.setSheetPage)
        gird.addWidget(self.searchName_cmb,1,1)

        selectPath_btn = QtWidgets.QPushButton(parent=girdBox)   # 建立'選擇檔案'按鈕
        selectPath_btn.setText('選擇檔案')
        selectPath_btn.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred,QtWidgets.QSizePolicy.Policy.Preferred)
        selectPath_btn.clicked.connect(self.openFile_Dialog)
        gird.addWidget(selectPath_btn, 0, 2,2,1)  
        
        gird.setRowMinimumHeight(2,int(path_lab.height()/2))#排版的空格

        self.door_ckb=QtWidgets.QCheckBox(parent=girdBox)   # 建立'大門紀錄'的勾選按鈕
        self.door_ckb.setText('大門紀錄')
        self.door_ckb.clicked.connect(self.setDetailPage)
        self.door_ckb.clicked.connect(self.setSheetPage)
        self.door_ckb.setChecked(True)
        self.door_ckb.setEnabled(False)
        gird.addWidget(self.door_ckb,3,0)
        self.mainLayout.addWidget(girdBox)

        self.Stockroom_ckb=QtWidgets.QCheckBox(parent=girdBox)  # 建立'庫房紀錄'的勾選按鈕
        self.Stockroom_ckb.setText('庫房紀錄 ')
        self.Stockroom_ckb.clicked.connect(self.setDetailPage)
        self.Stockroom_ckb.clicked.connect(self.setSheetPage)
        self.Stockroom_ckb.setEnabled(False)
        gird.addWidget(self.Stockroom_ckb,3,1)
        self.mainLayout.addWidget(girdBox)
        
    #設定結果輸出的UI
    def setFileResultUI(self):
        
        #建立輸出格式
        self.result_tab=QtWidgets.QTabWidget()

        #建立詳細資料頁面
        self.page_detail_box=QtWidgets.QListWidget()
        page_detail_layout=QtWidgets.QVBoxLayout()
        page_detail_layout.addWidget(self.page_detail_box)

        #建立表格資料頁面
        self.page_sheet_layout=QtWidgets.QGridLayout()
        self.page_sheet_box=QtWidgets.QWidget()
        self.page_sheet_box.setLayout(self.page_sheet_layout)


        self.result_tab.addTab(self.page_detail_box,"詳細資料")
        self.result_tab.addTab(self.page_sheet_box,"表格資料")
        self.mainLayout.addWidget(self.result_tab)

#========================================================================
#自訂意涵式
    #開啟檔案_有檔案選擇視窗
    def openFile_Dialog(self):
        filePath = QtWidgets.QFileDialog.getOpenFileName(filter="xls (*xls);;xlsx (*xlsx);;ALL (*)")
        self.path_input.setText(filePath[0])
        self.openFile()

    #開啟檔案
    def openFile(self):
        if(self.path_input.text()==''):
            return 
        
        try:
            self.ER.readFile(self.path_input.text())
            result=self.ER.loadName()
            self.searchName_cmb.addItem('-')
            self.searchName_cmb.addItems(result)
            self.searchName_cmb.setEnabled(True)
            self.door_ckb.setEnabled(True)
            self.Stockroom_ckb.setEnabled(True)
        except:
            self.showMessageBOx("檔案輸入錯誤",'請確認輸入檔案的正確性')
            

    #對話框
    def showMessageBOx(self,title:str,info:str):
        mbox=QtWidgets.QMessageBox(self)
        mbox.information(self,title,info)

    #設定詳細資料頁面
    def setDetailPage(self):
        #透過filter篩選站號
        #0:無 1:大門 2:庫房 3:大門+庫房
        station_filter=0
        if(self.door_ckb.isChecked()):
            station_filter+=1
        if(self.Stockroom_ckb.isChecked()):
            station_filter+=2

        self.page_detail_box.clear()
        all_data=self.ER.searchFile(self.searchName_cmb.currentText())
        attendanceRecord,detailRecord=self.ER.sortData(all_data,station_filter)
        for index in range(len(attendanceRecord)):
            self.addDetailItem(attendanceRecord[index],detailRecord[index],self.page_detail_box)

    #新增詳細資料的自定義物件
    def addDetailItem(self,title:str,detail:str,parentList:QtWidgets.QListWidget):
        list_item=QtWidgets.QListWidgetItem(parentList)
        item_widget=ExpandableItem(title,detail,list_item)
        list_item.setSizeHint(item_widget.sizeHint())
        parentList.addItem(list_item)
        parentList.setItemWidget(list_item,item_widget)

    #設定表格資料頁面
    def setSheetPage(self):
        self.clearSheetPage()
        if(self.searchName_cmb.currentText()=='-'):
            return
        all_data=self.ER.searchFile(self.searchName_cmb.currentText())
        # [0]:日期str [1]:上班bool [2]:下班bool
        decideResult=self.ER.decideResult(all_data)
        totaldays=calendar.monthrange(int(self.ER.year),int(self.ER.month))
        firstweekday=dt.date(int(self.ER.year),int(self.ER.month),1).isoweekday()
        
        titlelabel=QtWidgets.QLabel(parent=self.page_sheet_box)
        titlelabel.setText(f'{self.ER.year}-{self.ER.month}月')
        titlelabel.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred,QtWidgets.QSizePolicy.Policy.Maximum)
        titlelabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        titlelabel.setStyleSheet("background-color: gray;color:white;")
        self.page_sheet_layout.addWidget(titlelabel,0,0,1,7)
        
        weeklist=["星期日","星期一","星期二","星期三","星期四","星期五","星期六"]
        for i in range(len(weeklist)):
            tempLabel=QtWidgets.QLabel(text=weeklist[i],parent=self.page_sheet_box)
            tempLabel.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred,QtWidgets.QSizePolicy.Policy.Maximum)
            tempLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
            tempLabel.setStyleSheet("background-color: gray;color:white;")
            self.page_sheet_layout.addWidget(tempLabel,1,i)

        for day in range(totaldays[1]):
            x=dt.date(int(self.ER.year),int(self.ER.month),(day+1)).isoweekday()
            x=x % 7
            y=int(((firstweekday%7)+day)/7)+2
            
            if(len(decideResult) and f'{self.ER.year}-{str(int(self.ER.month))}-{str(day+1).zfill(2)}' in decideResult[0][0]):
                self.addSheetItem((day+1),decideResult[0][1],decideResult[0][2],x,y)
                decideResult.pop(0)
            else:
                self.addSheetItem((day+1),None,None,x,y)
    
    #新增表格資料的自定義物件
    def addSheetItem(self,day,on,left,x,y):
        if(on==None and left==None):
            info='<font color="black">無資料</font>'
        elif(on and left):
            info='<font color="green">正常</font>'
        elif(not on and not left):
            info='<font color="red">遲到/早退</font>'
        elif(not on and left):
            info='<font color="blue">遲到</font>'
        elif(on and not left):
            info='<font color="orange">早退</font>'
        self.page_sheet_layout.addWidget(CalendarItem(str(day).zfill(2),info),y,x)

    #重置畫面
    def reset(self):
        self.path_input.clear()
        self.searchName_cmb.clear()
        self.searchName_cmb.setEnabled(False)
        self.page_detail_box.clear()
        self.clearSheetPage()
        self.door_ckb.setEnabled(False)
        self.door_ckb.setChecked(True)
        self.Stockroom_ckb.setEnabled(False)
        self.Stockroom_ckb.setChecked(False)
        self.ER=ExcelViewer()

    def clearSheetPage(self):
        while self.page_sheet_layout.count():
            item=self.page_sheet_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()  # 安全刪除控件
            del item  # 釋放項目

#========================================================================

if __name__=='__main__':
    app=QtWidgets.QApplication(sys.argv)

    Form=MyWidget()
    Form.show()

    sys.exit(app.exec())