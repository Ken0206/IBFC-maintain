
### IBFC-maintain.py
```
國票 text file 匯入 excel file︰
將 IBFC-maintain.py 放在 txt 目錄
直接執行 IBFC-maintain.py 就可產出各 excel file
```
可正確處理檔名︰
```1-computer_name.txt```

---
會錯誤的檔名︰
```IB3-computer_name.txt```

處理方法(例)︰
```
IB3-computer_name.txt 改名 為 I-computer_name.txt
執行 IBFC-maintain.py 產生 *.xlsx
修改 I.xlsx 為 IB3.xlsx
開 Excel 編輯 IB3.xlsx 內 活頁頁籤名稱 computer_name
```
