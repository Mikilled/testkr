import spacy
import re
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph


def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def find_name(text):
    doc = nlp(text)
    names = [ent.text for ent in doc.ents if ent.label_ == "PER"]
    if names:
        print(re.split('\n|\t', ' '.join(names)))
        return True
    return False

def find_group(text):
    pattern = r"\b[А-Я]{2,3}-\d{2,3}\b"
    match = re.search(pattern, text)
    if match:
        print(match.group(0))
        return True
    return False
def find_lab(text):
    pattern = r'((?:[Пп]о\s)?[Пп]рактик[ае]\s№\d+|(?:по\s)?[Лл]абораторн[ао][яй]\sработ[ае]\s№?\d+)'
    match = re.search(pattern, text)
    if match:
        print(match.group(0))
        return True
    return False
def find_varkaf(text):
    pattern = r'(Вариант\s\d+)|(Кафедра|Факультет\s)|((?:По\s)?[Дд]исциплин[ае]:?\s*(.*))|((?:По\s)?[Тт]ем[ае]:?\s*(.*))'
    match = re.search(pattern, text)
    if match:
        print(text)
        # print(match.group(0))
        return True
    return False
def find_group(text):
    pattern = r'[\'"«»,](.*?)[\'"«»,]'
    match = re.search(pattern, text)
    if match:
        print(match.group(0))
        return True
    return False


def pars_info(text):
    if find_name(text):
        return
    if find_group(text):
        return
    if find_lab(text):
        return
    if find_varkaf(text):
        return


import docx

nlp = spacy.load("ru_core_news_sm")
doc = docx.Document(r"test.docx")
count = 0
text = ''
for i, block in enumerate(iter_block_items(doc)):
    if count == 30:
        break
    count += 1
    if isinstance(block,Paragraph):
        pars_info(block.text)

        text +=block.text

    else:
        for row in block.rows:
            for cell in row.cells:
                pars_info(cell.text)
                text += cell.text


