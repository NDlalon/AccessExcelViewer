# 門禁紀錄excel閱讀器
透過PyQt6開發的excel閱讀器，針對以大門門禁紀錄為出勤紀錄的情境，能夠快速從眾多資料中找到自己當月的出缺勤紀錄。  
:warning:目前對於excel文件的格式限制較多。

# 格式限制
在目前的版本中對於excel的格式限制較多，若格式上存在差異可能導致程式錯誤以下為重要限制

## 1.工作表
工作表的名稱需如下圖  
![image](https://github.com/user-attachments/assets/2006efad-3ed6-4ecb-86ba-3238b29cbec3)  
不須完全一樣但至少要有年分與月份如:202505  
並且要讀取的工作列必須放置於最前面，否則程式無法讀取
## 2.欄位
在欄位中須包含時間、站號、名稱  
並且在時間欄位中，需要以特定格式來區分紀錄的日期如下圖  
![image](https://github.com/user-attachments/assets/2fafd8b2-e01e-436f-b2b0-f72773d6de45)  
在時間中若沒有日期作為分隔，則所有資料室為同一天  
其他未提到的欄位則可忽略

# ico.py
在ico.py中可自行設定視窗icon，只需將圖片的二進位輸入進icon=b''即可  
  
```python
icon=b''

def getIcon_bin():
    return icon
```
