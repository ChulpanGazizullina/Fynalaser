import operator
from collections import defaultdict
from math import trunc
import openpyxl
import shutil
import telebot
from telebot import types

# Глобальные переменные
user_table = None
parameters = None

# Указать свой токен
token = '7223850119:AAHLOW42ogKqOjEHjdZU74CWzv-sLR0gf0M'
finalyser_bot = telebot.TeleBot(token)

# Словарь вопросов
questions = {
    'Сколько ты планируешь потратить на сайт с системой CRM?': {'answer': 0, 'cells': ['D15']},
    'Логотипы, баннеры, макеты?': {'answer': 0, 'cells': ['D16']},
    'Стартовая реклама?': {'answer': 0, 'cells': ['D17']},
    'Риски (в рублях)?': {'answer': 0, 'cells': ['D18']},
    'Себестоимость одной единицы продукта?': {'answer': 0, 'cells': ['D24']},
    'Количество продукта на закупке, шт?': {'answer': 0, 'cells': ['C24']},
    'Стоимость доставки продукта за шт?': {'answer': 0, 'cells': ['D32']},
    'Стоимость лида?': {'answer': 0, 'cells': ['D1']},
    'За сколько вы планируете продавать продукт? (Средний чек)': {'answer': 0, 'cells': ['D33']},
    'Сколько лидов планируете получить в 1-й месяц?': {'answer': 0, 'cells': ['C40', 'D40', 'E40', 'F40', 'G40', 'H40',
                                                                             'I40', 'J40', 'K40', 'L40', 'M40', 'N40']},
    'Сколько лидов планируете получить во 2-й месяц?': {'answer': 0, 'cells': ['C37']},
    'Сколько лидов планируете получить в 3-й месяц?': {'answer': 0, 'cells': ['E37']},
    'Сколько лидов планируете получить в 4-й месяц?': {'answer': 0, 'cells': ['F37']},
    'Сколько лидов планируете получить в 5-й месяц?': {'answer': 0, 'cells': ['G37']},
    'Сколько лидов планируете получить в 6-й месяц?': {'answer': 0, 'cells': ['H37']},
    'Сколько лидов планируете получить в 7-й месяц?': {'answer': 0, 'cells': ['I37']},
    'Сколько лидов планируете получить в 8-й месяц?': {'answer': 0, 'cells': ['J37']},
    'Сколько лидов планируете получить в 9-й месяц?': {'answer': 0, 'cells': ['K37']},
    'Сколько лидов планируете получить в 10-й месяц?': {'answer': 0, 'cells': ['L37']},
    'Сколько лидов планируете получить в 11-й месяц?': {'answer': 0, 'cells': ['M37']},
    'Сколько лидов планируете получить в 12-й месяц?': {'answer': 0, 'cells': ['N37']},
    'Какая планируется средняя конверсия в продажи? (Введите число от 0 до 100, например "15")': {
        'answer': 0,
        'cells': ['C38', 'D38', 'E38', 'F38', 'G38', 'H38', 'I38', 'J38', 'K38', 'L38', 'M38', 'N38']
    },
    'Какую сумму планируете вынимать ежемесячно из проекта? (Если 0, напишите "0")': {
        'answer': 0,
        'cells': ['C56', 'D56', 'E56', 'F56', 'G56', 'H56', 'I56', 'J56', 'K56', 'L56', 'M56', 'N56']
    },
    'Зарплата маркетолога в месяц? (Если 0, напишите "0")': {
        'answer': 0,
        'cells': ['C57', 'D57', 'E57', 'F57', 'G57', 'H57', 'I57', 'J57', 'K57', 'L57', 'M57', 'N57']
    },
    'Зарплата бухгалтера в месяц? (Если 0, напишите "0")': {
        'answer': 0,
        'cells': ['C58', 'D58', 'E58', 'F58', 'G58', 'H58', 'I58', 'J58', 'K58', 'L58', 'M58', 'N58']
    },
    'Зарплата на логистику в месяц? (Если 0, напишите "0")': {
        'answer': 0,
        'cells': ['C59', 'D59', 'E59', 'F59', 'G59', 'H59', 'I59', 'J59', 'K59', 'L59', 'M59', 'N59']
    },
    'Ежемесячный бюджет на рекламу?': {
        'answer': 0,
        'cells': ['C67', 'D67', 'E67', 'F67', 'G67', 'H67', 'I67', 'J67', 'K67', 'L67', 'M67', 'N67']
    },
    'Расходы на телефонию в месяц?': {
        'answer': 0,
        'cells': ['C68', 'D68', 'E68', 'F68', 'G68', 'H68', 'I68', 'J68', 'K68', 'L68', 'M68', 'N68']
    },
    'Расходы на интернет в месяц?': {
        'answer': 0,
        'cells': ['C69', 'D69', 'E69', 'F69', 'G69', 'H69', 'I69', 'J69', 'K69', 'L69', 'M69', 'N69']
    },
    'Расходы на дизайнера в месяц?': {
        'answer': 0,
        'cells': ['C70', 'D70', 'E70', 'F70', 'G70', 'H70', 'I70', 'J70', 'K70', 'L70', 'M70', 'N70']
    },
    'Аренда офиса в месяц?': {
        'answer': 0,
        'cells': ['C71', 'D71', 'E71', 'F71', 'G71', 'H71', 'I71', 'J71', 'K71', 'L71', 'M71', 'N71']
    }
}

# Состояния пользователей
user_state = defaultdict(int)


@finalyser_bot.message_handler(commands=['start'])
def send_welcome(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn = types.KeyboardButton("Начать!")
    markup.add(btn)
    finalyser_bot.send_message(message.chat.id, 'Привет! Я помогу тебе создать финансовую модель для стартапа.\n'
                                                'Жми «Начать!» и вводи параметры в рублях '
                                                '(если этот параметр отсутствует, пиши «0»).',
                                reply_markup=markup)


@finalyser_bot.message_handler(func=lambda message: message.text == 'Начать!')
def creating_a_table(message):
    global user_table, parameters

    file = 'original_template.xlsx'
    copy_file = 'finmodel.xlsx'
    shutil.copyfile(file, copy_file)
    user_table = openpyxl.load_workbook(copy_file)
    parameters = user_table['Параметры']

    user_state[message.chat.id] = 0
    first_question = list(questions.items())[0][0]
    finalyser_bot.send_message(message.chat.id, first_question)


@finalyser_bot.message_handler(func=lambda msg: True)
def ask_questions(message):
    global user_table, parameters

    chat_id = message.chat.id
    index = user_state[chat_id]
    quest_list = list(questions.items())

    if index >= len(quest_list):
        finalyser_bot.send_message(chat_id, "Все вопросы уже заданы.")
        return

    quest, info = quest_list[index]

    # Проверка числа
    try:
        user_input = int(message.text.replace(" ", ""))
    except ValueError:
        finalyser_bot.send_message(chat_id, "Введите корректное целое число.")
        return

    questions[quest]['answer'] = user_input
    for cell in info['cells']:
        parameters[cell] = user_input

    user_table.save("finmodel.xlsx")

    index += 1
    user_state[chat_id] = index

    if index < len(quest_list):
        next_question = quest_list[index][0]
        finalyser_bot.send_message(chat_id, next_question)
    else:
        finalyser_bot.send_message(chat_id, "Готово! Отправляю файл.")
        with open("finmodel.xlsx", "rb") as file:
            finalyser_bot.send_document(chat_id, file)


finalyser_bot.infinity_polling()
