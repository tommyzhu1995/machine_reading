import pandas as pd
from sklearn import preprocessing
import openpyxl

# Title行均不计入索引中
excel_path = 'original.xlsx'
# d = pd.read_excel(excel_path, sheet_name=0, usecols=[8,9,11,12,13])
d = pd.read_excel(excel_path, sheet_name=0)
totalRowNumber = d.shape[0]
keyword8 = ['other', '登录问题', '物流问题', '下单问题']
keyword9 = ['other', '不了解任何渠道', '根据iNet上的操作手册自己解决', '联系ITS', '联系OPS']
keyword11 = ['德勤', '京东&苏宁', '线下合作']
keyword12 = ['德勤', '京东&苏宁', '线下合作', 'other']
keyword13 = ['other', '爱车', '宠物', '餐饮', '购物', '丽人', '学习培训', '医疗健康', '运动健身', '亲子', '休闲娱乐']


def specification(keywordList, columnIndex):
    zeroList = [0 for index in range(len(keywordList))]
    oneHotList = pd.DataFrame(columns=keywordList)
    for rowNumber in range(totalRowNumber):
        cellValue = str(d.iloc[rowNumber, columnIndex])
        # print(rowNumber, cellValue, sep=': ')
        lineValue = []
        if cellValue == 'nan':
            lineValue = zeroList
        else:
            for i, word in enumerate(keywordList):
                score = cellValue.count(word)
                if score > 0:
                    score = 1
                lineValue.append(score)
        oneHotList.loc[rowNumber] = lineValue
    print(oneHotList)
    return oneHotList


def oneHotEncoder(oneHotList):
    enc = preprocessing.OneHotEncoder()
    encoding = enc.fit_transform(oneHotList)
    result = encoding.toarray()
    return result


def writeNewColumn(columnIndex, result):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.worksheets[0]
    ws.insert_cols(columnIndex+2)
    for index, row in enumerate(ws.rows):
        if index == 0:
            row[columnIndex+1].value = 'onehot'
        else:
            row[columnIndex+1].value = str(result[index-1])
    wb.save('original.xlsx')
    wb.close()


def wholeProcess(keywordList, columnIndex, offset):
    realColumnIndex = columnIndex+offset
    hotList = specification(keywordList, columnIndex)
    newColumn = oneHotEncoder(hotList)
    writeNewColumn(realColumnIndex, newColumn)


if __name__ == "__main__":
    wholeProcess(keyword8, 8, 0)
    wholeProcess(keyword9, 9, 1)
    wholeProcess(keyword11, 11, 2)
    wholeProcess(keyword12, 12, 3)
    wholeProcess(keyword13, 13, 4)
