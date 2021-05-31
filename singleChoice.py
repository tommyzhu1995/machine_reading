import pandas as pd
import re
import jieba.analyse as analyse
import jieba
import pareto
from collections import Counter

# Title行均不计入索引中
excel_path = 'DCoinSurveyETL.xlsx'
# d = pd.read_excel(excel_path, sheet_name=0, usecols=[8,9,11,12,13])
d = pd.read_excel(excel_path, sheet_name=0)
totalRowNumber = d.shape[0]


def cut_sent(para):
    para = re.sub('([。\;；！？\?])([^”’])', r"\1\n\2", para)  # 单字符断句符
    para = re.sub('(\.{6})([^”’])', r"\1\n\2", para)  # 英文省略号
    para = re.sub('(\…{2})([^”’])', r"\1\n\2", para)  # 中文省略号
    para = re.sub('([。；！？\?][”’])([^，。！？\?])', r'\1\n\2', para)
    # 如果双引号前有终止符，那么双引号才是句子的终点，把分句符\n放到双引号后，注意前面的几句都小心保留了双引号
    para = para.rstrip()  # 段尾如果有多余的\n就去掉它
    # 很多规则中会考虑分号;，但是这里我把它忽略不计，破折号、英文双引号等同样忽略，需要的再做些简单调整即可。
    return para.split("\n")


# def get_para_ques(question):
#     topk = 2
#     tf_result = analyse.extract_tags(question, topK=topk, withWeight=False)
#     return tf_result


def seg_word(text):
    text_temp = list(jieba.cut(text))
    text = Counter(text_temp)
    text_temp = [k for k in text.keys()]
    return text_temp


def sentenceCount_n_keywordCount(columnIndex):
    keywordList = []
    counter = 0
    for rowNumber in range(totalRowNumber):
        cellValue = str(d.iloc[rowNumber, columnIndex])
        # cellValue = cellValue.replace('/ Login problem', '')
        # cellValue = cellValue.replace('/ Logistics problem', '')
        # cellValue = cellValue.replace('/ Purchase problem', '')
        # cellValue = cellValue.replace('other', '')
        # cellValue = cellValue.replace('/', ';')
        # cellValue = cellValue.replace('问题', '')
        if cellValue == 'nan' or cellValue == ' ':
            continue
        else:
            template = cut_sent(cellValue)
            for sentence in template:
                for i in '，。；,.;':
                    sentence = sentence.replace(i, '')
                keyword = seg_word(sentence)
                keywordList = keywordList + keyword
                counter = counter + 1
        # print(sentence)
    answer = Counter(keywordList)
    print(counter)
    print(answer)
    newFrame = pd.DataFrame.from_dict(answer, orient='index')
    bookName = "Sheet"+str(columnIndex)
    newFrame.to_excel(excel_writer=writer, sheet_name=bookName)


if __name__ == "__main__":
    writer = pd.ExcelWriter('多选题分析v2.0.xlsx')
    sentenceCount_n_keywordCount(8)
    sentenceCount_n_keywordCount(9)
    sentenceCount_n_keywordCount(11)
    sentenceCount_n_keywordCount(12)
    sentenceCount_n_keywordCount(13)
    writer.save()
    writer.close()
