參考: https://docs.python.org/zh-tw/3.13/library/venv.html

先取得目前路徑

```
pwd
```

因為需要安裝包 為了不全域安裝 需要在目前路徑建立虛擬環境

```
python3 -m venv 剛剛得到的路徑
```

進入虛擬環境 取得 activate 路徑 大概在專案資料夾的這邊

```
source 剛剛得到的路徑加上/bin/activate
```

安裝套件 可以使用 pip 或 pip3 盡量都使用 pip3 因為我使用 3.13.2
不要一個用 pip 裝 另外一個用 pip3

```
pip3 install openpyxl
```

執行

```
python3 app.py
```