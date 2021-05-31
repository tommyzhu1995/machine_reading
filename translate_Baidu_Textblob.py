import httplib2
import urllib
import random
import json
from hashlib import md5
import pandas as pd
import time
from textblob import TextBlob


class baiduTranslate:
    appid = '20210408000768352'  # 你的appid
    secretKey = 'xKxDrPV3oYQRuOe83IFP'  # 你的密钥

    httpClient = None
    # myurl = 'https://fanyi-api.baidu.com/api/trans/vip/translate'
    myurl = 'http://api.fanyi.baidu.com/api/trans/vip/translate'

    fromLang = 'en'  # 翻译源语言
    toLang = 'zn'  # 译文语言
    salt = random.randint(32768, 65536)

    def translate(self, question):
        # 签名
        sign = self.appid + question + str(self.salt) + self.secretKey
        m1 = md5()
        m1.update(sign.encode(encoding='utf-8'))
        sign = m1.hexdigest()
        myUrl = self.myurl + '?q=' + urllib.parse.quote(
            question) + '&from=' + self.fromLang + '&to=' + self.toLang + '&appid=' + self.appid + '&salt=' + str(
            self.salt) + '&sign=' + sign

        try:
            h = httplib2.Http('.cache')
            response, content = h.request(myUrl)
            if response.status == 200:
                response = json.loads(content.decode('utf-8'))  # loads将json数据加载为dict格式
        except httplib2.ServerNotFoundError:
            print("Site is Down!")

        return response["trans_result"][0]['dst']


def senti(columnIndex):
    excel_path = 'DCoinSurveyETL.xlsx'
    d = pd.read_excel(excel_path, sheet_name=0)
    totalRowNumber = d.shape[0]
    newDataFrame = pd.DataFrame(columns=['翻译前', '翻译后'])
    Trans = baiduTranslate()

    for rowNumber in range(totalRowNumber):
        print("Translating column {0} row {1}".format(columnIndex, rowNumber))
        cellValue = str(d.iloc[rowNumber, columnIndex])
        lineValue = [cellValue]
        if len(cellValue) <= 3:
            continue
        elif rowNumber == 730:
            continue
        else:
            translation = Trans.translate(cellValue)
            lineValue.append(translation)
        newDataFrame.loc[rowNumber] = lineValue
        time.sleep(1)

    newDataFrame['sentiment'] = ''
    dataFrameRowNumber = newDataFrame.shape[0]

    for singleRow in range(dataFrameRowNumber):
        sentence = newDataFrame.iloc[singleRow, 1]
        testimonial = TextBlob(sentence)
        single_polarity = testimonial.sentiment.polarity
        newDataFrame.iloc[singleRow, 2] = single_polarity

    return newDataFrame


if __name__ == "__main__":
    column5 = senti(5)
    column14 = senti(14)
    with pd.ExcelWriter('情感分析.xlsx') as writer:
        column5.to_excel(writer, encoding='utf_8_sig', sheet_name='未消费原因分析')
        column14.to_excel(writer, encoding='utf_8_sig', sheet_name='功能建议分析')
