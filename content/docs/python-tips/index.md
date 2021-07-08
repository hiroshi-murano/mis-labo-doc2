---
title: 'Python Tips'
date: 2019-02-11T19:27:37+10:00
draft: false
weight: 1
---

<!-- TOC -->

- [1. 基本テンプレート](#1-%E5%9F%BA%E6%9C%AC%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88)
- [2. json](#2-json)
    - [2.1. 書き込み](#21-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF)
    - [2.2. 読み込み](#22-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF)
- [3. テキストファイル](#3-%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB)
    - [3.1. 書き込み](#31-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF)
    - [3.2. 読み込み](#32-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF)
- [4. ファイル名・path](#4-%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E5%90%8D%E3%83%BBpath)
- [5. openpyxl](#5-openpyxl)
    - [5.1. 読み込み](#51-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF)
    - [5.2. 書き込み](#52-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF)
    - [5.3. 書き込み書式付き](#53-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF%E6%9B%B8%E5%BC%8F%E4%BB%98%E3%81%8D)

<!-- /TOC -->

# 1. 基本テンプレート
<a id="markdown-%E5%9F%BA%E6%9C%AC%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88" name="%E5%9F%BA%E6%9C%AC%E3%83%86%E3%83%B3%E3%83%97%E3%83%AC%E3%83%BC%E3%83%88"></a>





```
import datetime


def func1():
    """
    """

    today = datetime.datetime.now()
    print(today.strftime("%Y/%m/%d %H:%M:%S"))


if __name__ == '__main__':

    func1()
```


# 2. json
<a id="markdown-json" name="json"></a>

## 2.1. 書き込み
<a id="markdown-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF" name="%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF"></a>

```
import json
import codecs

dictData = {}
dictData['名前'] = '斉藤'
dictData['年齢'] = 25
dictData['体重'] = 54.3
dictData['入社日'] = '1995-09-15'

json_text = json.dumps(dictData, ensure_ascii=False, indent=2)

fp = codecs.open('sample.json', 'w', 'utf-8')
fp.write(json_text)
fp.close()
```


## 2.2. 読み込み
<a id="markdown-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF" name="%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF"></a>

```
import json
import codecs

fp = codecs.open('sample.json', 'r', 'utf-8')
json_text = fp.read()
fp.close()

dictData = json.loads(json_text)
print(dictData)
```



# 3. テキストファイル
<a id="markdown-%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB" name="%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB"></a>

## 3.1. 書き込み
<a id="markdown-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF" name="%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF"></a>

```
import codecs

text = 'ああああああ\nいいいいいいい\nうううううう'

fp = codecs.open('sample.txt', 'w', 'utf-8')
fp.write(text)
fp.close()
```


## 3.2. 読み込み
<a id="markdown-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF" name="%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF"></a>

```
import codecs

fp = codecs.open('sample.txt', 'r', 'utf-8')
text = fp.read()
fp.close()

print(text)
```


# 4. ファイル名・path
<a id="markdown-%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E5%90%8D%E3%83%BBpath" name="%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E5%90%8D%E3%83%BBpath"></a>


```
import os

fname = r'c:\temp\abc.xlsx'

dirname = os.path.dirname(fname)
print(dirname)

basename = os.path.basename(fname)
print(basename)

root, ext = os.path.splitext(basename)
print(root)
print(ext)
```


# 5. openpyxl
<a id="markdown-openpyxl" name="openpyxl"></a>


## 5.1. 読み込み
<a id="markdown-%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF" name="%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%81%BF"></a>

```
import openpyxl

# セルへのアクセス
wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Sheet1']

for row in range(1,  ws.max_row + 1):
    for col in range(1,  ws.max_column + 1):
        data = ws.cell(row=row, column=col).value
        print(data)

wb.close()


# iter_rowsを使って行単位でアクセス
wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Sheet1']

for row in ws.iter_rows():
    print(row[0].value)
    print(row[1].value)
    print(row[2].value)
    print(row[3].value)

wb.close()

# valuesを使ってシートを一気に読み込む
wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Sheet1']

for row in ws.values:
    print(row)

wb.close()
```



## 5.2. 書き込み
<a id="markdown-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF" name="%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF"></a>



```
import openpyxl
import datetime

# 既存ファイルへの書き込み
wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Sheet1']

ws.cell(row=2, column=1).value = '伊集院'
ws.cell(row=2, column=2).value = 43
ws.cell(row=2, column=3).value = 59.4
ws.cell(row=2, column=4).value = datetime.datetime.today()

wb.save('test2.xlsx')


# 新規ファイル作成
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'サンプルデータ'

ws.cell(row=2, column=1).value = '伊集院'
ws.cell(row=2, column=2).value = 43
ws.cell(row=2, column=3).value = 59.4
ws.cell(row=2, column=4).value = datetime.datetime.today()

wb.save('test3.xlsx')
```

## 5.3. 書き込み(書式付き)
<a id="markdown-%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF%E6%9B%B8%E5%BC%8F%E4%BB%98%E3%81%8D" name="%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%81%BF%E6%9B%B8%E5%BC%8F%E4%BB%98%E3%81%8D"></a>

```
import datetime
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill
from openpyxl.styles import Font


def writeData(cell, txt, color):
    """
    フォント・背景色、罫線を指定してテキストを書き込む
    """

    side = Side(style='thin', color='000000')
    border = Border(top=side, bottom=side, left=side, right=side)
    fill = PatternFill(patternType='solid', fgColor=color)
    font = Font(name='Meiryo UI')

    cell.value = txt
    cell.border = border
    cell.fill = fill
    cell.font = font


def writeText(cell, txt):
    """
    フォントを指定してテキストを書き込む
    """

    font = Font(name='Meiryo UI')
    cell.value = txt
    cell.font = font


def func1():
    """
    """

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'サンプルデータ'

    writeData(ws.cell(row=2, column=1), '伊集院', 'ddFFFF')
    writeData(ws.cell(row=2, column=2), 43, 'FFFFee')
    writeText(ws.cell(row=2, column=3), 54.3)
    writeText(ws.cell(row=2, column=4), datetime.datetime.today())

    wb.save('test4.xlsx')


if __name__ == '__main__':

    func1()
```