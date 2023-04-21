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
    match = re.search(r'([Кк]афедра|[Фф]акультет\s)', text)
    if match:
        if "афедра" in text:
            info['кафедра'] = text
        if "акультет" in text:
            info['факультет'] = text
        return True
    match = re.search(r'(?:По\s)?[Дд]исциплин[ае]:?\s*(.*)', text)
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
    match = re.search(pattern, text)
    if match:
        info['группа'] = text
        return True
    return False


def pars_info(text):
    # match = re.search(r'([Фф]акультет\s)', text)
    # if match:
    #     info['факультет'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Кк]афедра\s)', text)
    # if match:
    #     info['кафедра'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Оо]тчет по:)', text)
    # if match:
    #     info['тип'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Тт]ема:\s)', text)
    # if match:
    #     info['тема'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Дд]исциплина:\s)', text)
    # if match:
    #     info['дисциплина'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Вв]ариант:\s)', text)
    # if match:
    #     info['вариант'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Гг]руппа:\s)', text)
    # if match:
    #     info['группа'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Вв]ыполнил:\s)', text)
    # if match:
    #     info['выполнил'] = text.split(':',1)[1].strip()
    # match = re.search(r'([Пп]роверил:\s)', text)
    # if match:
    #     info['проверил'] = text.split(':',1)[1].strip()
    fields = [
        (r'[Фф]акультет\s', 'факультет'),
        (r'[Кк]афедра\s', 'кафедра'),
        (r'[Оо]тчет по:', 'тип'),
        (r'[Тт]ема:\s', 'тема'),
        (r'[Дд]исциплина:\s', 'дисциплина'),
        (r'[Вв]ариант:\s', 'вариант'),
        (r'[Гг]руппа:\s', 'группа'),
        (r'[Вв]ыполнил:\s', 'выполнил'),
        (r'[Пп]роверил:\s', 'проверил'),
        (r'[Цц]ель работы:\s', 'цель'),
        (r'[Зз]адание:\s', 'задание')
    ]

    for pattern, key in fields:
        match = re.search(pattern, text)
        if match:
            info[key] = text.split(':', 1)[1].strip()

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
    "цель": "",
    "задание": "",
}




names_tit = []
nlp = spacy.load("ru_core_news_sm")
doc = docx.Document(r"form.docx")
count = 0
text = ''
for i, block in enumerate(iter_block_items(doc)):
    if count == 60:
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




print(info)