import pandas as pd
import re
import jieba
import wordcloud
from scipy.misc import imread
from string import punctuation
import matplotlib.pyplot as plt

excel_path = '情感分析.xlsx'
mask = imread('cloud.png')
d = pd.read_excel(excel_path, sheet_name=1)
totalRowNumber = d.shape[0]

with open('remove_words.txt', 'r', encoding="utf-8") as fp:
    remove_words = fp.read().split()

add_punc = '0123456789'
all_punc = punctuation + add_punc

object_list = ''
for rowNumber in range(totalRowNumber):
    cellValue = str(d.iloc[rowNumber, 1])
    cellValue = re.sub('[a-zA-Z]', '', cellValue)
    cellValue = cellValue.replace('\n', '')
    seg_list_exact = jieba.cut(cellValue, cut_all=False)
    for word in seg_list_exact:
        # 看每个分词是否在常用词表中或结果是否为空或\xa0不间断空白符，如果不是再追加
        if word not in remove_words and word != ' ' and word != '\xa0':
            object_list = ' '.join([object_list, word])

print(object_list)

w = wordcloud.WordCloud(background_color="white",  # 设置背景颜色
                        mask=mask,  # 设置背景图片
                        width=1000,
                        height=860,
                        max_words=150,  # 设置最大现实的字数
                        stopwords=wordcloud.STOPWORDS,  # 设置停用词
                        max_font_size=150,  # 设置字体最大值
                        random_state=30,
                        collocations=False)
w.generate(object_list)
file_name = '功能改进_词云.png'
w.to_file(file_name)

img_plt = plt.imread(file_name)
plt.imshow(img_plt)
plt.show()
