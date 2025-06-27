import pandas as pd
import os.path as op
from datetime import datetime
import datetime as dt


class ExcelViewer:
    #初始化
    def __init__(self,path=''):
        self.weeklist=["星期一","星期二","星期三","星期四","星期五","星期六","星期日"]
        if(path!=''):
            self.readFile(path)

    # 讀取檔案
    def readFile(self,path):
        self.path=path
        extension=op.splitext(path)[1]

        if(extension=='.xls'):
            self.rawdata=pd.read_excel(path,engine='xlrd')
            date=pd.ExcelFile(path).sheet_names[0].split('_')[0]
        elif(extension=='.xlsx'):
            self.rawdata=pd.read_excel(path,engine='openpyxl')
            date=pd.ExcelFile(path).sheet_names[0].split('_')[0]
        else:
            raise Exception('File read error: only .xls and .xlsx files are supported. Please check your input file.')
        
        self.year=f"{date[0]}{date[1]}{date[2]}{date[3]}"
        self.month=f"{date[4]}{date[5]}"

    #取得檔案中全部姓名
    def loadName(self):
        result=[]
        
        for index,row in self.rawdata.iterrows():
            if(pd.notna(row['名稱']) and str(row['名稱']) not in result and pd.notna(row['站號'])): #取得所有不重複名稱，並過濾空值與異常值
                result.append(row['名稱'])

        return result
    
    #搜尋檔案
    def searchFile(self,name=''):
        result=[]
        
        for index, row in self.rawdata.iterrows():
            temp=[]
            if(name in str(row['名稱'])):
                temp.append(row['時間'])
                temp.append(row['站號'])
                temp.append(row['名稱'])
                result.append(temp)
            elif('/' in str(row['時間'])):
                temp.append(row['時間'])
                result.append(temp)

        return result
    
    #整理檔案
    def sortData(self,searchResult,station_filter=1):
        attendanceRecord=[]
        detailRecord=[]
        temp=''
        detailTemp=''
        clockin=None
        clockout=None
        date=''

        #[0]:時間,[1]站號,[2]:名稱
        for row in searchResult:
            if('/' in row[0]):
                if(clockin !=None and clockout != None):
                    temp=self.getColorStr(date,clockin,clockout)
                    attendanceRecord.append(temp)
                    detailRecord.append(detailTemp)
                    temp=''
                    detailTemp=''
                    date=self.getDate(row[0])
                    clockin=None
                    clockout=None
                else:
                    date=self.getDate(row[0])
                    detailTemp=''
            else:
                if(row[1]=='大門'):
                    if(clockin==None and clockout==None):
                        clockin=datetime.strptime(row[0],'%H:%M:%S')
                        clockout=datetime.strptime(row[0],'%H:%M:%S')
                    elif(clockin>datetime.strptime(row[0],'%H:%M:%S')):
                        clockin=datetime.strptime(row[0],'%H:%M:%S')
                    elif(clockout<datetime.strptime(row[0],'%H:%M:%S')):
                        clockout=datetime.strptime(row[0],'%H:%M:%S')
                #依照filter篩選顯示站號(0:無 1:大門 2:庫房 3:大門+庫房)
                if((station_filter==1 or station_filter==3) and row[1]=='大門'):
                    detailTemp+=row[2]+' '+datetime.strptime(row[0],'%H:%M:%S').strftime('%H:%M:%S')+' '+row[1]+'\n'
                elif((station_filter==2 or station_filter==3) and row[1]=='庫房'):
                    detailTemp+=row[2]+' '+datetime.strptime(row[0],'%H:%M:%S').strftime('%H:%M:%S')+' '+row[1]+'\n'
        if(clockin !=None and clockout != None):              
            temp=self.getColorStr(date,clockin,clockout)
            attendanceRecord.append(temp)
            detailRecord.append(detailTemp)

        return attendanceRecord,detailRecord

    def decideResult(self,searchResult):
        Result=[]
        clockin=None
        clockout=None
        date=''

        #[0]:時間,[1]站號,[2]:名稱
        for row in searchResult:
            if('/' in row[0]):
                if(clockin !=None and clockout != None):
                    #判斷上班時間是否遲到 正常時間為00:00:00~08:00:59
                    if(clockin>=datetime.strptime("08:01:00", "%H:%M:%S")):
                        on=False
                    else:
                        on=True
                    #判斷下班時間是否提早 正常時間為17:00:00~23:59:59
                    if(clockout<datetime.strptime("17:30:00", "%H:%M:%S")):
                        left=False
                    else:
                        left=True
                    Result.append([date,on,left])
                    date=self.getDate(row[0])
                    clockin=None
                    clockout=None
                else:
                    date=self.getDate(row[0])
            else:
                if(row[1]=='大門'):
                    if(clockin==None and clockout==None):
                        clockin=datetime.strptime(row[0],'%H:%M:%S')
                        clockout=datetime.strptime(row[0],'%H:%M:%S')
                    elif(clockin>datetime.strptime(row[0],'%H:%M:%S')):
                        clockin=datetime.strptime(row[0],'%H:%M:%S')
                    elif(clockout<datetime.strptime(row[0],'%H:%M:%S')):
                        clockout=datetime.strptime(row[0],'%H:%M:%S')
        if(clockin !=None and clockout != None):              
            #判斷上班時間是否遲到 正常時間為00:00:00~08:00:59
            if(clockin>=datetime.strptime("08:01:00", "%H:%M:%S")):
                on=False
            else:
                on=True
            #判斷下班時間是否提早 正常時間為17:00:00~23:59:59
            if(clockout<datetime.strptime("17:30:00", "%H:%M:%S")):
                left=False
            else:
                left=True
            Result.append([date,on,left])

        return Result

    #取得日期(年-月-日-星期)
    def getDate(self,date:str):
        year_int=int(self.year)
        month_int=int(self.month)
        day_int=int(date.split('/')[1])
        weekday=self.weeklist[dt.date(year_int,month_int,day_int).weekday()]
        
        result=self.year+'-'+str(month_int)+'-'+str(day_int).zfill(2)+" "+weekday

        return result

    #取得排序後並上色的文字
    def getColorStr(self,date,clockin,clockout):
        #判斷上班時間是否遲到 正常時間為00:00:00~08:00:59
        if(clockin>=datetime.strptime("08:01:00", "%H:%M:%S")):
            strtime=clockin.strftime('%H:%M:%S')
            result=f'{date} <font color="red">上班時間:{strtime}</font> '
        else:
            strtime=clockin.strftime('%H:%M:%S')
            result=f'{date} <font color="green">上班時間:{strtime}</font> '
        result+='-----'
        #判斷下班時間是否提早 正常時間為17:00:00~23:59:59
        if(clockout<datetime.strptime("17:30:00", "%H:%M:%S")):
            strtime=clockout.strftime('%H:%M:%S')
            result+=f'<font color="red">下班時間:{strtime}</font>'
        else:
            strtime=clockout.strftime('%H:%M:%S')
            result+=f'<font color="green">下班時間:{strtime}</font>'

        return result