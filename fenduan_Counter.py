import jieba
import jieba.analyse as analyse
from collections import Counter


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
    tf_result = analyse.extract_tags(question, topK=topk, withWeight=False)
    print(tf_result)
    return tf_result


def calculation_para(para, question):
    common = Counter(para) & Counter(question)
    num_same = sum(common.values())
    return num_same


class para_text:
    def __init__(self, path):
        self.path = path
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
        sample = self.seg_list
        question_paraphrased = get_para_ques(ques_text)
        # 求解出每个文档中与答案最相关的段落
        most_related_para = -1
        max_related_score = 0
        for p_idx, para_tokens in enumerate(sample):
            # 计算相同词出现次数
            if len(sample) > 0:
                related_score = calculation_para(para_tokens, question_paraphrased)
                print(related_score)
            else:
                continue
            # 判断，如果计算得到的分数大于最大分数时
            # 则修改最相关段落为当前段落，最大值为当前值
            if related_score > max_related_score:
                most_related_para = p_idx
                max_related_score = related_score
        # 最终保存下，每个文档与答案最相关的段落编号
        choice = most_related_para
        answer = ''.join(sample[choice])
        # 输出该段落
        print(sample[choice])
        print(choice, answer, sep=',')
        return answer


if __name__ == "__main__":
    # text是输入的问题
    text_seglist = para_text('shoucefiles/yuangongshouce.txt')
    input_data = "报道需要准备什么文件？"
    paragraph_wanted = text_seglist.find_most_related(input_data)