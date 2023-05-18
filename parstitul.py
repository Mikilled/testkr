from functools import reduce

import spacy
import re
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import operator


def pars_titul(doc_name):
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

    def find_name(text,table):
        doc = nlp(text)
        names = [ent.text for ent in doc.ents if ent.label_ == "PER"]
        j = []
        if names:
            for i in names:
                j.append(re.split("|".join("\n\t"), i))
            for i in j:
                if table:
                    names_tit.append( i[0])
                    continue
                if len(i)==1:
                    names_tit.insert(0,i[0])
                else:
                    names_tit.insert(0,i[0])
                    names_tit.append(i[1])
                    break

            #print(re.split('\n|\t', ' '.join(names)))
            return True
        return False


    def find_varkaf(text):
        pattern = r'(Вариант\s\d+)|(Кафедра|Факультет\s)|((?:По\s)?[Дд]исциплин[ае]:?\s*(.*))|((?:По\s)?[Тт]ем[ае]:?\s*(.*))'
        match = re.search(r'(Вариант\s\d+)', text)
        if match:
            info['вариант'] = text
            return True
        match = re.search(r'(?:[Пп]о\s)?[Дд]исциплин[ае]:?\s*(.*)', text)
        if match:
            info['дисциплина'] = match.group(1)
            return True
        match = re.findall(r"\b[А-Я]{2,3}-\d{2,3}\b", text)
        if match:
            #print(match)
            info['группа'] = ', '.join(match)
            return True
        pattern = r'((?:[Пп]о\s)?[Пп]рактик[ае]\s№\d+|(?:по\s)?[Лл]абораторн[ао][яй]\sработ[ае]\s№?\d+)'
        match = re.search('((?:[Пп]о\s)?[Пп]рактик[ае]\s№\d+|(?:[Пп]о\s)?[Лл]абораторн[ао][яй]\sработ[ае]\s№?\d+|(?:[Пп]о\s)?[Пп]рактическ[ао][яй]\sработ[ае]\s№?\d+)', text)
        if match:
            info['тип'] = match.group(0)
            return True
        if info['группа'] == "":
            match = re.search(r'([\'"«»,](?!.*_)(.*?)[\'"«»,])', text)
            if match:
                info['тема'] = text
                return True
        return False


    def pars_info(text,table):
        for i in text.split('\n'):
            find_varkaf(i)
            find_name(i,table)


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
    doc = docx.Document(doc_name)
    count = 0
    text = ''

    for block in iter_block_items(doc):

        if count == 30 or ((isinstance(block, Paragraph))and(('овосибирск' in block.text) or('202' in block.text))):
            # Ваш код дальше
            if '202' not in block.text:
                continue
            count += 1
            break
        if isinstance(block,Paragraph):
            c = block.text
            pars_info(c,False)
            text +=block.text

        else:
            availability_table = True
            for row in block.rows:
                for cell in row.cells:
                    c = cell.text
                    pars_info(c,True)
                    #print(f"{cell.text} ------1---1---1----1---1---")
                    text += cell.text
        count += 1




    info['выполнил'] = '\n'.join(names_tit[:-1])
    info['проверил'] = names_tit[-1]


    if info['группа'].lower().split('-')[0] == ('аб' or 'абс'):
        info['факультет'] = 'Факультет автоматики и вычислительной техники'
        info["кафедра"] = 'Кафедра защиты информации'

    #print(f"итог - {info['группа']}")
    return [count,info]


if __name__ == '__main__':
    pars_titul(r"C:\Users\admin\Downloads\Otchyot_po_P_1_Poduto_E_I__ABs-123.docx")

