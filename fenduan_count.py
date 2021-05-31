import jieba
import jieba.analyse as analyse
import docx
# 安装的包为python-docx


def wordtotxt(path):
    path = 'test2.docx'
    doc = docx.Document(path)
    docText = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    # print(docText)

    newPathLoc = path.find(".doc")
    newPath = path[:newPathLoc] + '.txt'

    with open(newPath, "w", encoding='utf-8') as f:
        f.write(docText)

    return newPath


def seg_word(currentText):
    text_temp = list(jieba.cut(currentText))
    return text_temp


def get_para_ques(question):
    topk = 0
    if len(question) <= 10:
        topk = 3
    elif len(question) < 20:
        topk = 5
    elif len(question) >= 20:
        topk = 7
    # 若答案不准确，可尝试更改关键词数量topK
    tf_result = analyse.extract_tags(question, topK=topk, withWeight=True)
    print(tf_result)
    return tf_result


def calculation_para(para, question, score):
    num_counter = 0
    num_same = 0
    for i, word in enumerate(question):
        num = para.count(word) * (score[i]+1)
        num_same = num_same + num
        if para.count(word) > 0:
            num_counter = num_counter + 1
    return [num_counter, num_same]


class para_text:
    def __init__(self, path):
        self.path = wordtotxt(path)
        text_list = []
        # 打开文件
        # 添加转码function
        try:
            with open(self.path, 'r', encoding='gb18030') as f:
                lines = f.readlines()
                text_full = ''.join(lines)
                f.close()
        except UnicodeDecodeError:
            with open(self.path, 'r', encoding='utf8') as f:
                lines = f.readlines()
                text_full = ''.join(lines)
                f.close()
        # 去除无用字符
        text_full = text_full.replace('\n', '')
        text_full = text_full.replace('\u3000', '')
        text_full = text_full.replace('…', '')
        text_full = text_full.replace('\u0020', '')
        text_full = text_full.replace(' ', '')
        # 截取文字
        while len(text_full) > 450:
            append_text = text_full[0:449]
            # while u'\u4e00' <= append_text[-1] <= u'\u9fff':
            while append_text[-1] != u'\u3002' and append_text[-1] != u'\u002e' and append_text[-1] != u'\uff1b' \
                    and append_text[-1] != u'\u003b':
                append_text = append_text[:-1]
            text_list.append(append_text)
            text_full = text_full.replace(text_full[0:449], '')
        # 将整段文章做成list
        self.seg_list = [seg_word(textlist) for textlist in text_list]

    def find_most_related(self, ques_text):
        question_paraphrased = []
        question_paraphrased_score = []
        sample = self.seg_list
        question_paraphrased_temp = get_para_ques(ques_text)
        for word, weight in question_paraphrased_temp:
            question_paraphrased.append(word)
            question_paraphrased_score.append(weight)
        # 求解出每个文档中与答案最相关的段落
        most_related_para = -1
        max_related_counter = 0
        max_related_score = 0
        for p_idx, para_tokens in enumerate(sample):
            # 计算相同词出现次数
            if len(sample) > 0:
                [related_counter, related_score] = calculation_para(para_tokens, question_paraphrased,
                                                                    question_paraphrased_score)
                print(p_idx, related_counter, related_score, sep=', ')
            else:
                continue
            # 判断，如果计算得到的分数大于最大分数时
            # 则修改最相关段落为当前段落，最大值为当前值
            if related_counter > max_related_counter:
                most_related_para = p_idx
                max_related_counter = related_counter
                max_related_score = related_score
            elif related_counter == max_related_counter:
                if related_score > max_related_score:
                    most_related_para = p_idx
                    max_related_counter = related_counter
                    max_related_score = related_score
        # 最终保存下，每个文档与答案最相关的段落编号
        choice = most_related_para
        answer = ''.join(sample[choice])
        # 输出该段落
        print(choice, answer, sep='. ')
        return answer


if __name__ == "__main__":
    # text是输入的问题x
    text_seglist = para_text(r'C:\Users\tozhu\PycharmProjects\pythonProject\test2.docx')
    # 例题
    input_data = "年度休假如何申请"
    paragraph_wanted = text_seglist.find_most_related(input_data)
