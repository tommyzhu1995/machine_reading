import pandas as pd
import re
import jieba.analyse as analyse
import jieba
from collections import Counter

# Title行均不计入索引中
excel_path = 'DCoinSurveyETL.xlsx'
# d = pd.read_excel(excel_path, sheet_name=0, usecols=[8,9,11,12,13])
d = pd.read_excel(excel_path, sheet_name=0)
totalRowNumber = d.shape[0]
keywordList14 = [['商户', '合作', '门店', '商家', '用场'], ['种类', '品类', '品种'], ['页面', '界面', '购物车', '系统', '速度'],
                 ['功能', '提现', '小程序', '红包', '公共交通', '游戏', '充值', '银行'], ['人民币', '现金', '支付宝', '话费', '兑换', '余额'],
                 ['渠道', '途径'], ['折扣', '优惠', '价格'], ['客服', '售后', '退换', '物流'], ['评论']]
keywordList5 = [['合适', '心仪', '种类', '品种', '市面', '有理', '用场', '合作', '场景', '线下', '性价比', '兑换', '吸引力', '选好'
                    , '理想', '品类', '购买'], ['优惠', '价格', '网络', '市面', '活动'], ['地址'], ['麻烦', '费时', '手机', '系统'],
                ['需求'], ['差价', '金额', '现金'], ['余额', '数量', '数额', '不足', 'not enough', '足够', '账户', '金额', '积分'],
                ['未来', '大件', '攒钱'], ['刚刚', '习惯', '时间', '淘宝', '机会', '购物'], ['周边']]
keywordClassify5 = [0, 0, 1, 1, 2, 3, 2, 2, 2, 0]
keywordClassify14 = [0, 0, 1, 3, 2, 2, 0, 3, 3]
finalScore5 = [0, 0, 0, 0]
finalScore14 = [0, 0, 0, 0]


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
    return text


def sentenceCount_n_keywordCount(columnIndex, keywordList):
    counter = 0
    score = [0 for index in range(len(keywordList))]
    for rowNumber in range(totalRowNumber):
        cellValue = str(d.iloc[rowNumber, columnIndex])
        if cellValue == 'nan' or cellValue == ' ':
            continue
        else:
            template = cut_sent(cellValue)
            for sentence in template:
                for i in '，。；,.;':
                    sentence = sentence.replace(i, '')
                counter = counter + 1
                keyword = seg_word(sentence)
                for index, singleList in enumerate(keywordList):
                    common = keyword & Counter(singleList)
                    if sum(common.values()) > 0:
                        score[index] = score[index] + 1
    print(counter, score, sep=': ')
    return score


if __name__ == "__main__":
    score5 = sentenceCount_n_keywordCount(5, keywordList5)
    score14 = sentenceCount_n_keywordCount(14, keywordList14)
    for i in range(len(score5)):
        finalScore5[keywordClassify5[i]] += score5[i]
    print(finalScore5)
    for i in range(len(score14)):
        finalScore14[keywordClassify14[i]] += score14[i]
    print(finalScore14)
