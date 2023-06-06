import os
import zipfile
from telebot import *
from docxcompose.composer import Composer
from docx import Document as Document_compose
from parstitul import pars_titul
from writetitul import write_titul
from pocessinginfile import process_document
import time as tm
from datetime import date
token = '6295000781:AAHo4pljJ1GmkRXRier4-ckD16n8sD5rCjo'
bot = telebot.TeleBot(token)




spams = {}


def is_spam(user_id, spams, msgs, max, ban, msg):
    try:
        usr = spams[user_id]
        usr["messages"] += 1
    except:
        spams[user_id] = {"next_time": int(tm.time()) + max, "messages": 1, "banned": 0}
        usr = spams[user_id]
    if usr["banned"] >= int(tm.time()):
        print(u'' + str((datetime.today()).strftime("%H:%M ")) + f"спамит {user_id}")
        return False
    else:
        if usr["next_time"] >= int(tm.time()):
            if usr["messages"] >= msgs:
                spams[user_id]["banned"] = tm.time() + ban
                if not (msg == ''):
                    text = f"{msg}{int(ban)} секунд"
                    bot.send_message(user_id, text)
                return False
        else:
            spams[user_id]["messages"] = 1
            spams[user_id]["next_time"] = int(tm.time()) + max
    return True




@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    upl = types.KeyboardButton("Загрузить отчет")
    inf = types.KeyboardButton("о боте")
    markup.add(upl,inf)
    print(message.chat.username)
    text = """
    Бот не делает:
    -Пишет за вас отчет
    -Заполняет не достающими данными (название картинок и тд.)
    -Не проверяет работу
    Бат делает:
    -Правильно оформляет титульный лист
    -Выставляет правильные шрифты, отступы
    -Находит ошибки оформления
    -Оформляет документы в формате docx
    """
    bot.send_message(message.chat.id, text, reply_markup=markup)


@bot.message_handler(content_types=['text'])
def actions(message):
    if not is_spam(message.chat.id, spams, 10, 10, 10, "бан на  "):
        return
    text = """
        *ВНИМАНИЕ!!!*
        Мини гайд как заполнить отчет:
        - Заголовки должны быть выделены *жирным* шрифтом
        - Все картинки должны иметь подпись \n          (Рисунок 1 – ER-диаграмма)
        - Все таблицы должны иметь заголовок \n         (Таблица 1 - Определение сущностей)
        - Титульный лист должен содержать город \n          (Новосибирск 2023)
    """
    if message.text == "Загрузить отчет":
        bot.send_message(message.chat.id, text, parse_mode='Markdown')
    if message.text == "о боте":
        text1 = text + """
            Бот не делает:
            -Пишет за вас отчет
            -Заполняет не достающими данными (название картинок и тд.)
            -Не проверяет работу
            Бат делает:
            -Правильно оформляет титульный лист
            -Выставляет правильные шрифты, отступы
            -Находит ошибки оформления
            - и тд.
            """
        bot.send_message(message.chat.id, text1, parse_mode='Markdown')



@bot.message_handler(content_types=['document'])
def handle_document(message):
    if not is_spam(message.chat.id, spams, 2, 15, 10, "Отправлять можно раз в "):
        return
    file_id = message.document.file_id
    file_info = bot.get_file(file_id)
    file_path = file_info.file_path
    print(message.document.file_name)
    if message.document.file_name.split('.')[-1] != 'docx':
        bot.send_message(message.chat.id, 'это не docx документ')
        return

    downloaded_file = bot.download_file(file_path)
    print(message.chat.username)
    if not os.path.exists("docx"):
        os.makedirs("docx")
    with open(f'docx/{message.document.file_id}.docx', 'wb') as new_file:
        new_file.write(downloaded_file)

    try:
        count, info = pars_titul(f'docx/{message.document.file_id}.docx')
        print("спарсил титул")
    except Exception as e:
        print(e)
        bot.send_message(message.chat.id, 'Ошибка парсинга титульника')
        del_files(message.document.file_id)
        return
    try:
        outdoc = write_titul(f'docx/titul-{message.document.file_id}.docx', info)
        outdoc.save(f'docx/titul-{message.document.file_id}.docx')
        print("Записал титул")
    except Exception as e:
        print(e)
        bot.send_message(message.chat.id, 'Ошибка записи')
        del_files(message.document.file_id)
        return

    try:
        new_doc,errors = process_document(f'docx/{message.document.file_id}.docx', count)
        new_doc.save(f'docx/text-{message.document.file_id}.docx')
    except Exception as e:
        print(e)
        bot.send_message(message.chat.id, 'Ошибка обработки документа')
        del_files(message.document.file_id)
        return

    try:
        master = Document_compose(f'docx/titul-{message.document.file_id}.docx')
        composer = Composer(master)
        doc2 = Document_compose(f'docx/text-{message.document.file_id}.docx')
        composer.append(doc2)
        composer.save(f'docx/new{message.document.file_id}.docx')
    except Exception as e:
        print(e)
        bot.send_message(message.chat.id, 'Ошибка склейки документов')
        del_files(message.document.file_id)
        return
    mail = ''
    if len(errors) != 0:
        mail = '\n'.join(errors)
    with open(f'docx/new{message.document.file_id}.docx', 'rb') as new_file:
        bot.send_document(message.chat.id, new_file, caption=mail)

    if not os.path.exists("zip"):
        os.makedirs("zip")

    with zipfile.ZipFile(rf'zip/docum.zip', 'w') as myzip:
        myzip.write(rf'docx/{message.document.file_id}.docx', compress_type=zipfile.ZIP_DEFLATED, arcname=f"{message.from_user.username}")

    del_files(message.document.file_id)

def del_files(name):
    try:
        os.remove(f'docx/titul-{name}.docx')
    except:
        pass
    try:
        os.remove(f'docx/text-{name}.docx')
    except:
        pass
    try:
        os.remove(f'docx/new{name}.docx')
    except:
        pass
    try:
        os.remove(f'docx/{name}.docx')
    except:
        pass


bot.polling(none_stop=True)