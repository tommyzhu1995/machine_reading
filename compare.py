import abc
import re
import difflib
import pandas as pd
from fuzzywuzzy import fuzz


class CompareBase(metaclass = abc.ABCMeta):
    def __init__(self):
        pass

    def _preprocess_space(self, text):
        #先删除掉两边的空白字符
        p_text = text.strip()

        ## 删除掉不是在英文字符之间空格
        p_text = re.sub(r'\s+', ' ', p_text)
        pt1 = r'(?<![a-zA-Z]) (?![a-zA-Z])'
        pt2 = r'(?<=[a-zA-Z]) (?![a-zA-Z])'
        # pt = '{}|{}'.format(pt1, pt2)
        pt = pt1
        p_text = re.sub(pt, '', p_text)
        return p_text

    def _process_text(self, text):
        ## 全角转半角
        char_string = '''ａ a ｂ b ｃ c ｄ d ｅ e ｆ f ｇ g ｈ h ｉ i ｊ j ｋ k ｌ l ｍ m ｎ n ｏ o ｐ p ｑ q ｒ r ｓ s ｔ t ｕ u ｖ v ｗ w ｘ x ｙ y ｚ z Ａ A Ｂ B Ｃ C Ｄ D Ｅ E Ｆ F Ｇ G Ｈ H Ｉ I Ｊ J Ｋ K Ｌ L Ｍ M Ｎ N Ｏ O Ｐ P Ｑ Q Ｒ R Ｓ S Ｔ T Ｕ U Ｖ V Ｗ W Ｘ X Ｙ Y Ｚ Z １ 1 ２ 2 ３ 3 ４ 4 ５ 5 ６ 6 ７ 7 ８ 8 ９ 9 ０ 0 ｀ ` ” " ’ ‘ “ " ‘ ‘ ＿ _ － - ～ ~ ＝ = ＋ + ＼ \ ｜ | ／ / （ ( ） ) ［ [ ］ ] 【 [ 】 ] ｛ { ｝ } ＜ < ＞ > ． . ， , ； ; ： : ！ ! ＾ ^ ％ % ＃ # ＠ @ ＄ $ ＆ & ？ ? ＊ * 。 .'''
        char_list = char_string.split()
        full_angle_list = char_list[::2]
        semi_angle_list = char_list[1::2]
        full2semi = dict(list(zip(full_angle_list, semi_angle_list)))

        p_text = ''
        for char in text:
            p_char = full2semi.get(char, char)
            p_char = re.sub(r'\u200b', '', p_char)
            p_text += p_char
        
        p_text = re.sub(r"[_]{2,}", '', p_text)
        ## 处理空白字符
        p_text = self._preprocess_space(p_text)
        return p_text

    def preprocess(self, texts):
        p_texts = [self._process_text(text) for text in texts]
        return p_texts

    def _join_digit_texts(self, digit_texts):
        digit_full_text = ''
        digit_endlines = []
        endline_index = 0
        for text in digit_texts:
            if digit_full_text == '':
                digit_full_text += text
                endline_index += len(text)
            else:
                text_len = len(text)
                temp_text = digit_full_text[-1] + ' ' + text
                p_temp_text = self._preprocess_space(temp_text)
                if p_temp_text == temp_text:
                    digit_full_text += (' ' + text)
                    endline_index += (text_len + 1)
                # elif p_temp_text == digit_full_text[-1] + text:
                #     digit_full_text += text
                #     endline_index += text_len
                # else:
                #     raise Exception('The space chars in "{}" should be preprocessed before this method is called.'.format(text))
                else:
                    digit_full_text += text
                    endline_index += text_len

            digit_endlines.append(endline_index)

        return digit_full_text, digit_endlines

    def _join_scanned_texts(self, scanned_texts):
        scanned_full_text = ' '.join(scanned_texts)
        scanned_full_text = self._preprocess_space(scanned_full_text)
        return scanned_full_text

    @abc.abstractmethod
    def _compare(self, digit_full_text, scanned_full_text):
        pass

    @abc.abstractmethod
    def _split_diff_res(self, diff_res, digit_endlines):
        pass

    def __call__(self, digit_texts, scanned_texts):
        p_digit_texts = self.preprocess(digit_texts)
        p_scanned_texts = self.preprocess(scanned_texts)
        if p_digit_texts == []:
            _p_digit_texts = [''] * len(scanned_texts)
            scores = [0] * len(scanned_texts)
            tags = ['ADD'] * len(scanned_texts)
            final_res = pd.DataFrame(list(zip(_p_digit_texts, p_scanned_texts,
                                       scores, tags
                                       )),
                              columns = ['unsigned', 'signed',
                                         'score', 'tag'
                                        ])
        else:
            digit_full_text, digit_endlines = self._join_digit_texts(p_digit_texts)
            scanned_full_text = self._join_scanned_texts(p_scanned_texts)
            diff_res = self._compare(digit_full_text, scanned_full_text)
            final_res = self._split_diff_res(diff_res, digit_endlines)
        return final_res

    def compare_all(self, df):
        """
        根据前面比对结果, 计算整体比对结果...
        """
        doc_data = "".join(list(df["unsigned"]))
        pdf_data = "".join(list(df["signed"]))
        score = fuzz.ratio(doc_data, pdf_data)
        return score


class Compare_Difflib(CompareBase):
    def _compare(self, digit_full_text, scanned_full_text):
        differ = difflib.Differ()
        diff_res = list(differ.compare(digit_full_text, scanned_full_text))
        return diff_res

    def _split_diff_res(self, diff_res, digit_endlines):
        digit_endlines_set = set(digit_endlines)
        index = 0
        digit_full_text = ''
        scanned_full_text = ''
        for char_diff in diff_res:
            if char_diff.startswith('+ '):
                scanned_full_text += char_diff[-1]
            else:
                if index in digit_endlines_set:
                    digit_full_text += '\n'
                    scanned_full_text += '\n'
                
                index += 1
                digit_full_text += char_diff[-1]
                if char_diff.startswith('  '):
                    scanned_full_text += char_diff[-1]

        digit_texts = digit_full_text.split('\n')
        scanned_texts = scanned_full_text.split('\n')

        scores = []
        tags = []
        for digit_text, scanned_text in zip(digit_texts, scanned_texts):
            _digit_text = digit_text.strip()
            _scanned_text = scanned_text.strip()
            score = fuzz.ratio(_digit_text, _scanned_text)
            scores.append(score)
            if score == 100:
                tag = 'PASS'
            else:
                if digit_text == '':
                    tag = 'ADD'
                elif scanned_text == '':
                    tag = 'SUB'
                else:
                    tag = 'MODIFY'
            tags.append(tag)

        res_df = pd.DataFrame(list(zip(digit_texts, scanned_texts,
                                       scores, tags
                                       )),
                              columns = ['unsigned', 'signed',
                                         'score', 'tag'
                                        ])
                    
        return res_df


class Compare_Difflib2(Compare_Difflib):
    def _zh_en(self, texts):
        zh_cnt = 0
        en_cnt = 0
        th = 0.3
        eps = 1e-5
        for text in texts:
            for char in text:
                char_code = ord(char)
                if '\u4e00' <= char <= '\u9fef':
                    zh_cnt += 1
                elif 65 <= char_code <= 90 or 97 <= char_code <= 122:
                    en_cnt += 1

        zh_en_ratio = zh_cnt * 1.0 / (zh_cnt + en_cnt + eps)
        if zh_en_ratio >= th:
            return 'zh'
        else:
            return 'en'

    def _join_digit_texts(self, digit_texts):
        lang = self._zh_en(digit_texts)
        if lang == 'en':
            _digit_texts = [re.split(r'\s+', text) for text in digit_texts]
            digit_full_text = []
            digit_endlines = []
            endline_index = 0
            for word_list in _digit_texts:
                digit_full_text += word_list
                endline_index += len(word_list)
                digit_endlines.append(endline_index)
        else:
            digit_full_text, digit_endlines = super()._join_digit_texts(digit_texts)
            digit_full_text = list(digit_full_text)

        return digit_full_text, digit_endlines, lang

    def _join_scanned_texts(self, scanned_texts):
        lang = self._zh_en(scanned_texts)
        if lang == 'en':
            _scanned_texts = [re.split(r'\s+', text) for text in scanned_texts]
            scanned_full_text = []
            for word_list in _scanned_texts:
                scanned_full_text += word_list
        else:
            scanned_full_text = super()._join_scanned_texts(scanned_texts)
            scanned_full_text = list(scanned_full_text)

        return scanned_full_text, lang

    def _compare(self, digit_full_text, scanned_full_text):
        differ = difflib.Differ()
        diff_res = list(differ.compare(digit_full_text, scanned_full_text))
        seq = difflib.SequenceMatcher(None, digit_full_text, scanned_full_text)
        score = round(seq.ratio() * 100)
        return diff_res, score

    def _split_diff_res(self, diff_res, digit_endlines, digit_lang, scanned_lang):
        digit_endlines_set = set(digit_endlines)
        index = 0
        digit_full_text = []
        scanned_full_text = []
        _digit_texts = []
        _scanned_texts = []
        for token_diff in diff_res:
            if token_diff.startswith('+ '):
                _scanned_texts += [token_diff[2:]]
            elif not token_diff.startswith('? '):
                if index in digit_endlines_set:
                    digit_full_text.append(_digit_texts)
                    scanned_full_text.append(_scanned_texts)
                    _digit_texts = []
                    _scanned_texts = []

                index += 1
                _digit_texts += [token_diff[2:]]
                if token_diff.startswith('  '):
                    _scanned_texts += [token_diff[2:]]

        if _digit_texts != []:
            digit_full_text.append(_digit_texts)

        if _scanned_texts != []:
            scanned_full_text.append(_scanned_texts)

        if digit_lang == 'en':
            digit_texts = [' '.join(tlist) for tlist in digit_full_text]
        else:
            digit_texts = [''.join(tlist) for tlist in digit_full_text]

        if scanned_lang == 'en':
            scanned_texts = [' '.join(tlist) for tlist in scanned_full_text]
        else:
            scanned_texts = [''.join(tlist) for tlist in scanned_full_text]

        scores = []
        tags = []
        for digit_text, scanned_text in zip(digit_texts, scanned_texts):
            _digit_text = digit_text.strip()
            _scanned_text = scanned_text.strip()
            score_1 = fuzz.ratio(_digit_text, _scanned_text)
            score_2 = fuzz.ratio(_digit_text.lower(), _scanned_text.lower())
            score = max(score_1, score_2)
            scores.append(score)
            if score == 100:
                tag = 'PASS'
            else:
                if digit_text == '':
                    tag = 'ADD'
                elif scanned_text == '':
                    tag = 'SUB'
                else:
                    tag = 'MODIFY'
            tags.append(tag)

        res_df = pd.DataFrame(list(zip(digit_texts, scanned_texts,
                                       scores, tags
                                       )),
                              columns = ['unsigned', 'signed',
                                         'score', 'tag'
                                        ])
                    
        return res_df

    def __call__(self, digit_texts, scanned_texts):
        p_digit_texts = self.preprocess(digit_texts)
        p_scanned_texts = self.preprocess(scanned_texts)
        if p_digit_texts == []:
            _p_digit_texts = [''] * len(scanned_texts)
            scores = [0] * len(scanned_texts)
            tags = ['ADD'] * len(scanned_texts)
            final_res = pd.DataFrame(list(zip(_p_digit_texts, p_scanned_texts,
                                       scores, tags
                                       )),
                              columns = ['unsigned', 'signed',
                                         'score', 'tag'
                                        ])
            score = 0
        else:
            digit_full_text, digit_endlines, digit_lang = self._join_digit_texts(p_digit_texts)
            scanned_full_text, scanned_lang = self._join_scanned_texts(p_scanned_texts)
            diff_res, score = self._compare(digit_full_text, scanned_full_text)
            final_res = self._split_diff_res(diff_res, digit_endlines, digit_lang, scanned_lang)
        return final_res, score


class Wrapped_Compare_Difflib2(Compare_Difflib2):
    def __call__(self, digit_texts, scanned_texts):
        final_res, score = super().__call__(digit_texts, scanned_texts)
        return final_res


class Compare_Difflib3(Compare_Difflib2):
    def _join_digit_texts(self, digit_texts):
        lang = self._zh_en(digit_texts)
        if lang == 'en':
            _digit_texts = [re.split(r'\s+', text) for text in digit_texts]
            digit_full_text = []
            digit_endlines = []
            endline_index = 0
            for word_list in _digit_texts:
                digit_full_text += word_list
                endline_index += len(word_list)
                digit_endlines.append(endline_index)
        else:
            # digit_full_text, digit_endlines = super()._join_digit_texts(digit_texts)
            # digit_full_text = list(digit_full_text)

            pt = r'([^\u4E00-\u9FEF]+)'
            _digit_texts = []
            for text in digit_texts:
                tokens = []
                slist = re.split(pt, text)
                for s in slist:
                    mobj = re.search(pt, s)
                    if mobj is None:
                        tokens.extend(list(s))
                    else:
                        tokens.append(mobj.group())
                _digit_texts.append(tokens)

            digit_full_text = []
            digit_endlines = []
            endline_index = 0
            for word_list in _digit_texts:
                digit_full_text += word_list
                endline_index += len(word_list)
                digit_endlines.append(endline_index)

        return digit_full_text, digit_endlines, lang

    def _join_scanned_texts(self, scanned_texts):
        lang = self._zh_en(scanned_texts)
        if lang == 'en':
            _scanned_texts = [re.split(r'\s+', text) for text in scanned_texts]
            scanned_full_text = []
            for word_list in _scanned_texts:
                scanned_full_text += word_list
        else:
            # scanned_full_text = super()._join_scanned_texts(scanned_texts)
            # scanned_full_text = list(scanned_full_text)

            pt = r'([^\u4E00-\u9FEF]+)'
            _scanned_texts = []
            for text in scanned_texts:
                tokens = []
                slist = re.split(pt, text)
                for s in slist:
                    mobj = re.search(pt, s)
                    if mobj is None:
                        tokens.extend(list(s))
                    else:
                        tokens.append(mobj.group())

                _scanned_texts.append(tokens)

            scanned_full_text = []
            for word_list in _scanned_texts:
                scanned_full_text += word_list

        return scanned_full_text, lang


class Wrapped_Compare_Difflib3(Compare_Difflib3):
    def __call__(self, digit_texts, scanned_texts):
        final_res, score = super().__call__(digit_texts, scanned_texts)
        return final_res


if __name__ == '__main__':
    digit_texts = '''合同编号：F20FTSNCWLJ0346-01 
技术服务合同  
项目名称： 23个样品高通量信息采集及分析 
委托方（甲方）：北京华益健康药物研究中心 
受托方（乙方）：武汉华大医学检验所有限公司 
签订时间：  2020 年  4 月  3 日 
签订地点： 北京 
有效期限： 2020 年  4 月  3 日 至  2021 年  11 月  12 日 
Page 1 of 13 '''.split('\n')

    scanned_texts = '''华大基因O|30合同编号：F20 FTSNCWLJO3460个
技术服务合同
VRAT项目名称：23个样品高通量信息采集及分析
委托方(甲方)：北京华益健康药物研究中心
受托方(乙方)：武汉华大医学检验所有限公司
签订时间：2020年4月3日
签订地点：北京
有效期限：2020年4月3日至2021年11月12日口
a1f13'''.split('\n')

    comp = Compare_Difflib()
    df = comp(digit_texts, scanned_texts)
    print(df)