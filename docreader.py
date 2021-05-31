import os
import pdfplumber
import subprocess
import re
import xml.etree.ElementTree as ET
from tempfile import TemporaryDirectory
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from ..utils import timeout_wrap
from .config import LIBRE_TIMEOUT

class DocReader(object):
    """
    读取Docx与PDF文档
    top: 70
    docx返回数据[p]
    pdf返回数据[{head: [], body: [], foot: []}]
    """

    def __init__(self, path, top=0):
        self.path = path
        self.top = top
        if path is None:
            return []
        path2 = path.lower()
        if path2.endswith(".docx"):
            self.type = 1
            self.data = self.read_docx()
            self.strip()
        elif path2.endswith(".doc"):
            self.type = 2
            ##
            self.data = self.read_doc3()

            ##
            # self.data = self.read_doc4()
            # self.strip()
        elif path2.endswith(".pdf"):
            self.type = 3
            if self.top > 0:
                self.data = self.read_pdf()
            else:
                self.data = self.read_pdf2()
        elif path2.endswith(".txt"):
            self.type = 4
            self.data = self.read_txt()
            self.strip()
        else:
            self.data = []
        pass

    def read_docx(self):
        doc = Document(self.path)
        full_text = []
        for block in self.iter_block_items(doc):
            if isinstance(block, Paragraph):
                full_text.append(block.text)
            elif isinstance(block, Table):
                try:
                    for row in block.rows:
                        for cell in self.iter_unique_cells(row):
                            for p in cell.paragraphs:
                                full_text.append(p.text)
                except:
                    print("error table...")
            else:
                print("undefined type: ", type(block))
        return full_text

    def read_doc(self):
        try:
            output = subprocess.check_output(
                ["/usr/local/bin/antiword", "-w 0 ", self.path]).decode("utf-8")
            arr = ['']
            for char in output:
                if char == '\n':
                    arr.append('')
                else:
                    arr[len(arr)-1] = arr[len(arr)-1]+char
        except Exception as e:
            print(str(e))
            arr = []
        return arr

    def read_doc2(self):
        try:
            output = subprocess.check_output(
                ['antiword', '-x', 'db', self.path]).decode('utf-8')

            root = ET.fromstring(output)
            bookinfo_list = root.findall('bookinfo')
            for bi in bookinfo_list:
                root.remove(bi)

            with TemporaryDirectory() as tmp_dir:
                docbook_file = os.path.join(tmp_dir, 'anti.xml')
                root_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
                with open(docbook_file, 'w') as f:
                    f.write(root_str)

                docx_file = os.path.join(tmp_dir, 'pandoc.docx')

                subprocess.check_output(
                    ['pandoc', docbook_file, '-f', 'docbook', '-t', 'docx', '-o', docx_file]
                )
                _doc_reader = DocReader(docx_file)

        except Exception as e:
            print(str(e))
            return []

        return _doc_reader.data

    @timeout_wrap(LIBRE_TIMEOUT)
    def read_doc3(self):
        try:
            with TemporaryDirectory() as tmp_dir:
                fn_ext = os.path.basename(self.path)
                fn, ext = os.path.splitext(fn_ext)
                docx_file = os.path.join(tmp_dir, '{}.docx'.format(fn))

                subprocess.check_output(
                    ['libreoffice', '--headless', '--invisible', '--convert-to', 'docx', self.path, '--outdir', tmp_dir]
                )
                _doc_reader = DocReader(docx_file)

        except Exception as e:
            print(str(e))
            return []

        return _doc_reader.data

    def read_doc4(self):
        try:
            output = subprocess.check_output(
                ["catdoc", "-w", self.path]).decode("utf-8")
            arr = ['']
            for char in output:
                if char == '\n':
                    arr.append('')
                else:
                    arr[len(arr)-1] = arr[len(arr)-1]+char
        except Exception as e:
            print(str(e))
            arr = []
        return arr

    def read_pdf(self):
        full_text = []
        with pdfplumber.open(self.path) as pdf:
            for page in pdf.pages:
                # texts = page.extract_text()
                # 标准格式下 去掉页眉 页脚
                # texts = page.crop((0, 70, 595, 800)).extract_text()
                text1 = page.crop((0, 0, page.width, self.top)).extract_text()
                text2 = page.crop((0, self.top, page.width, page.height)).extract_text()
                head = []
                content = []
                foot = []
                if text1:
                    head = text1.split("\n")
                if text2:
                    content = text2.split("\n")
                    last = content[-1]
                    if valid_foot(last):
                        content = content[:len(content) - 1]
                        foot.append(last)
                # 页眉大于30个字符，则认为是正文
                h_len = len("".join(head))
                if h_len > 30:
                    content = head + content
                    head = []
                content2 = []
                for line in content:
                    line = line.strip()
                    if line != "":
                        content2.append(line)
                page_data = {"head": head, "body": content2, "foot": foot}
                full_text.append(page_data)
        return full_text

    def read_pdf2(self):
        full_text = []
        with pdfplumber.open(self.path) as pdf:
            for page in pdf.pages:
                # texts = page.extract_text()
                texts = page.extract_text(y_tolerance=5)
                texts = texts.split("\n")
                full_text += texts
        return full_text

    def read_txt(self):
        full_text = []
        with open(self.path, encoding="utf-8") as f:
            for line in f.readlines():
                full_text.append(line)
        return full_text

    def strip(self):
        result = []
        for text in self.data:
            text = text.strip()
            if text != "":
                result.append(text)
        self.data = result

    @staticmethod
    def iter_block_items(parent):
        """
        Generate a reference to each paragraph and table child within *parent*,
        in document order. Each returned value is an instance of either Table or
        Paragraph. *parent* would most commonly be a reference to a main
        Document object, but also works for a _Cell object, which itself can
        contain paragraphs and tables.
        """
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
            # print(parent_elm.xml)
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    @staticmethod
    def iter_unique_cells(row):
        """Generate cells in *row* skipping empty grid cells.
          https://stackoverflow.com/questions/48090922/python-docx-row-cells-return-a-merged-cell-multiple-times
          合并单元格去重复
        """
        prior_tc = None
        for cell in row.cells:
            this_tc = cell._tc
            if this_tc is prior_tc:
                continue
            prior_tc = this_tc
            yield cell


# 校验是否是页码
foot_ptn = re.compile(r"(\d+|Page\d+of\d+|\d+/\d+)$")
def valid_foot(content):
    content = re.sub(r"\s", "", content)
    if re.match(foot_ptn, content):
        return True
    else: 
        return False

if __name__ == '__main__':
    # doc = DocReader("C:/Users/dxiang/Desktop/OCR测试模板/F20CQSSCSLJ1469-三方直销带修订标记合同.docx")
    # for p in doc.data:
    #     print(p)
    doc = DocReader(r"/root/bbb.doc")
    for p in doc.data:
        print(p)

