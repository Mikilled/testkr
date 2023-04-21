import spacy
import re
import docx
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
        names_tit.append(re.split('\n|\t', ' '.join(names)))
        #print(re.split('\n|\t', ' '.join(names)))
        return True
    return False

def find_group(text):
    pattern = r"\b[А-Я]{2,3}-\d{2,3}\b"
    match = re.search(pattern, text)
    if match:
        print(match.group(0))
        info['группа'] = match.group(0)
        return True
    return False
def find_lab(text):
    pattern = r'((?:[Пп]о\s)?[Пп]рактик[ае]\s№\d+|(?:по\s)?[Лл]абораторн[ао][яй]\sработ[ае]\s№?\d+)'
    match = re.search(pattern, text)
    if match:
        info['тип'] = match.group(0)
        return True
    return False
def find_varkaf(text):
    pattern = r'(Вариант\s\d+)|(Кафедра|Факультет\s)|((?:По\s)?[Дд]исциплин[ае]:?\s*(.*))|((?:По\s)?[Тт]ем[ае]:?\s*(.*))'
    match = re.search(r'(Вариант\s\d+)', text)
    if match:
        info['вариант'] = text
        return True
    match = re.search(r'[Кк]афедра|[Фф]акультет\s.*', text)
    if match:
        if "\n" in text:
            line = text.split("\n")
            for i in line:
                if "афедра" in i:
                    info['кафедра'] = i
                if "акультет" in i:
                    info['факультет'] = i
        else:
            if "афедра" in text:
                info['кафедра'] = text
            if "акультет" in text:
                info['факультет'] = text
        return True
    match = re.search(r'(?:[Пп]о\s)?[Дд]исциплин[ае]:?\s*(.*)', text)
    if match:
        info['дисциплина'] = text
        return True
    match = re.search(r'[\'"«»,](?!.*_)(.*?)[\'"«»,]', text)
    if match:
        info['тема'] = text
        return True
    return False
def find_group(text):
    pattern = r"\b[А-Я]{2,3}-\d{2,3}\b"
    match = re.findall(pattern, text)
    if match:
        info['группа'] = ''.join(match)
        return True
    return False


def pars_info(text):
    find_name(text)
    find_group(text)


    find_group(text)

    find_lab(text)

    find_varkaf(text)






info = {
    'факультет': "",
    "кафедра": "",
    "тип": "",
    "тема": "",
    "дисциплина": "",
    "вариант": "",
    'группа':"",
    "выполнил": "",
    "проверил": "",
}




names_tit = []
nlp = spacy.load("ru_core_news_sm")
doc = docx.Document(r"testwrite.docx")
count = 0
text = ''
for i, block in enumerate(iter_block_items(doc)):
    if count == 30:
        break
    count += 1
    if isinstance(block,Paragraph):
        c = block.text
        pars_info(c)
        text +=block.text

    else:
        for row in block.rows:
            for cell in row.cells:
                pars_info(cell.text)
                text += cell.text

info['проверил'] = names_tit[-1]
info['выполнил'] = names_tit[0]


print(info)