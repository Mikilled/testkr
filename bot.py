from telebot import *
import functools
from parstitul import pars_titul
from writetitul import write_titul
from fileprocessing import process_document
token = '6295000781:AAHo4pljJ1GmkRXRier4-ckD16n8sD5rCjo'
bot = telebot.TeleBot(token)
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    upl = types.KeyboardButton("Загрузить отчет")
    inf = types.KeyboardButton("о боте")
    markup.add(upl,inf)
    print(message.chat.username)
    bot.send_message(message.chat.id, 'Welcome to the club, buddy', reply_markup=markup)


@bot.message_handler(content_types=['text'])
def actions(message):
    if "Загрузить отчет":
        bot.send_message(message.chat.id, 'Загрузите отчет')


@bot.message_handler(content_types=['document'])
def handle_document(message):
    # Получение информации о файле
    file_id = message.document.file_id
    file_info = bot.get_file(file_id)
    file_path = file_info.file_path

    # Загрузка файла с помощью метода download_file
    downloaded_file = bot.download_file(file_path)
    print(message.chat.username)
    # Сохранение файла на сервере
    with open(f'docx/{message.chat.username}.docx', 'wb') as new_file:
        new_file.write(downloaded_file)
    count,info = pars_titul(f'docx/{message.chat.username}.docx')
    print(count)
    outdoc = write_titul(f'docx/new{message.chat.username}.docx', info)
    print(2222)
    new_doc = process_document(f'docx/{message.chat.username}.docx', outdoc, count)
    print(3333)
    new_doc.save(f'docx/new{message.chat.username}.docx')

bot.polling(none_stop=True)