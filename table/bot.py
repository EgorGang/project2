"""
Основной модуль Telegram-бота, который обрабатывает команды и взаимодействует с пользователем.
"""
import telebot
from telebot.types import ReplyKeyboardMarkup, KeyboardButton
from journal import Journal
from decorators import handle_exceptions
import os

# Инициализация бота
TOKEN = 'ваш токен'
bot = telebot.TeleBot(TOKEN)

# Экземпляр журнала
journal = Journal()

# Функция для создания клавиатуры
def main_menu():
    """
    Создаёт главное меню с доступными действиями для пользователя.
    :return: ReplyKeyboardMarkup объект.
    """
    markup = ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(KeyboardButton("Добавить ученика"))
    markup.add(KeyboardButton("Удалить ученика"))
    markup.add(KeyboardButton("Добавить оценку"))
    markup.add(KeyboardButton("Удалить оценку"))
    markup.add(KeyboardButton("Изменить оценку"))
    markup.add(KeyboardButton("Информация об ученике"))
    markup.add(KeyboardButton("Посмотреть журнал"))
    markup.add(KeyboardButton("Экспорт журнала"))
    return markup

# Обработчики команд
@bot.message_handler(commands=['start'])
def start(message):
    """
    Обработчик команды /start. Приветствует пользователя и отображает главное меню.
    :param message: Сообщение от пользователя.
    """
    bot.send_message(message.chat.id, "Добро пожаловать в классный журнал!\n"
                     "Данный бот предназначен для учета оценок учеников. По окончанию работы бота вы можете экспортировать журнал, чтобы сохранить данные об оценках.\n"
                                      "Выберите действие:", reply_markup=main_menu())

@bot.message_handler(func=lambda message: True)
@handle_exceptions
def handle_message(message):
    """
    Обработчик всех текстовых сообщений от пользователя.
    :param message: Сообщение от пользователя.
    """
    if message.text == "Добавить ученика":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, add_student)

    elif message.text == "Удалить ученика":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, remove_student)

    elif message.text == "Добавить оценку":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, add_mark_step1)

    elif message.text == "Удалить оценку":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, remove_mark_step1)

    elif message.text == "Изменить оценку":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, change_mark_step1)

    elif message.text == "Информация об ученике":
        msg = bot.send_message(message.chat.id, "Введите имя ученика:")
        bot.register_next_step_handler(msg, student_info)

    elif message.text == "Посмотреть журнал":
        bot.send_message(message.chat.id, journal.table.to_string(), reply_markup=main_menu())

    elif message.text == "Экспорт журнала":
        msg = bot.send_message(message.chat.id, "Введите имя файла (например, journal.xlsx):")
        bot.register_next_step_handler(msg, export_journal)

    else:
        bot.send_message(message.chat.id, "Неизвестная команда. Выберите действие:", reply_markup=main_menu())

# Функции обработки данных
@handle_exceptions
def add_student(message):
    """
    Добавляет нового ученика в журнал.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    result = journal.add_student(name)
    bot.send_message(message.chat.id, result, reply_markup=main_menu())

@handle_exceptions
def remove_student(message):
    """
    Удаляет ученика из журнала.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    result = journal.remove_student(name)
    bot.send_message(message.chat.id, result, reply_markup=main_menu())

@handle_exceptions
def add_mark_step1(message):
    """
    Первый шаг добавления оценки: ввод имени ученика.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    if name not in journal.table.index:
        bot.send_message(message.chat.id, "Такого ученика нет!", reply_markup=main_menu())
        return
    msg = bot.send_message(message.chat.id, "Введите число месяца (1-31):")
    bot.register_next_step_handler(msg, add_mark_step2, name)

@handle_exceptions
def add_mark_step2(message, name):
    """
    Второй шаг добавления оценки: ввод числа месяца.
    :param message: Сообщение от пользователя с числом месяца.
    :param name: Имя ученика.
    """
    try:
        day = int(message.text)
        if 1 <= day <= 31:
            msg = bot.send_message(message.chat.id, "Введите оценку:")
            bot.register_next_step_handler(msg, add_mark_step3, name, day)
        else:
            bot.send_message(message.chat.id, "Некорректное число месяца!", reply_markup=main_menu())
    except ValueError:
        bot.send_message(message.chat.id, "Некорректные данные!", reply_markup=main_menu())

@handle_exceptions
def add_mark_step3(message, name, day):
    """
    Третий шаг добавления оценки: ввод оценки.
    :param message: Сообщение от пользователя с оценкой.
    :param name: Имя ученика.
    :param day: Число месяца.
    """
    try:
        mark = float(message.text)
        result = journal.add_mark(name, day, mark)
        bot.send_message(message.chat.id, result, reply_markup=main_menu())
    except ValueError:
        bot.send_message(message.chat.id, "Некорректная оценка!", reply_markup=main_menu())

@handle_exceptions
def remove_mark_step1(message):
    """
    Первый шаг удаления оценки: ввод имени ученика.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    if name not in journal.table.index:
        bot.send_message(message.chat.id, "Такого ученика нет!", reply_markup=main_menu())
        return
    msg = bot.send_message(message.chat.id, "Введите число месяца (1-31):")
    bot.register_next_step_handler(msg, remove_mark_step2, name)

@handle_exceptions
def remove_mark_step2(message, name):
    """
    Второй шаг удаления оценки: ввод числа месяца.
    :param message: Сообщение от пользователя с числом месяца.
    :param name: Имя ученика.
    """
    try:
        day = int(message.text)
        if 1 <= day <= 31:
            result = journal.remove_mark(name, day)
            bot.send_message(message.chat.id, result, reply_markup=main_menu())
        else:
            bot.send_message(message.chat.id, "Некорректное число месяца!", reply_markup=main_menu())
    except ValueError:
        bot.send_message(message.chat.id, "Некорректные данные!", reply_markup=main_menu())

@handle_exceptions
def change_mark_step1(message):
    """
    Первый шаг изменения оценки: ввод имени ученика.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    if name not in journal.table.index:
        bot.send_message(message.chat.id, "Такого ученика нет!", reply_markup=main_menu())
        return
    msg = bot.send_message(message.chat.id, "Введите число месяца (1-31):")
    bot.register_next_step_handler(msg, change_mark_step2, name)

@handle_exceptions
def change_mark_step2(message, name):
    """
    Второй шаг изменения оценки: ввод числа месяца.
    :param message: Сообщение от пользователя с числом месяца.
    :param name: Имя ученика.
    """
    try:
        day = int(message.text)
        if 1 <= day <= 31:
            msg = bot.send_message(message.chat.id, "Введите новую оценку:")
            bot.register_next_step_handler(msg, change_mark_step3, name, day)
        else:
            bot.send_message(message.chat.id, "Некорректное число месяца!", reply_markup=main_menu())
    except ValueError:
        bot.send_message(message.chat.id, "Некорректные данные!", reply_markup=main_menu())

@handle_exceptions
def change_mark_step3(message, name, day):
    """
    Третий шаг изменения оценки: ввод новой оценки.
    :param message: Сообщение от пользователя с новой оценкой.
    :param name: Имя ученика.
    :param day: Число месяца.
    """
    try:    
        new_mark = float(message.text)
        result = journal.change_mark(name, day, new_mark)    
        bot.send_message(message.chat.id, result, reply_markup=main_menu())
    except ValueError:    
        bot.send_message(message.chat.id, "Некорректная оценка!", reply_markup=main_menu())

@handle_exceptions
def student_info(message):
    """
    Предоставляет информацию об ученике.
    :param message: Сообщение от пользователя с именем ученика.
    """
    name = message.text
    result = journal.student_info(name)
    bot.send_message(message.chat.id, result, reply_markup=main_menu())

@handle_exceptions
def export_journal(message):
    """
    Экспортирует журнал в файл и отправляет его пользователю.
    :param message: Сообщение от пользователя с именем файла.
    """
    filename = message.text
    if not filename.endswith('.xlsx'):
        bot.send_message(
            message.chat.id,
            "Некорректный формат файла. Пожалуйста, используйте расширение .xlsx",
            reply_markup=main_menu()
        )
        return

    try:
        # Экспорт журнала в файл
        filepath = journal.export_to_excel(filename)
        
        # Проверка, существует ли файл
        if os.path.exists(filepath):
            # Отправка файла пользователю
            with open(filepath, 'rb') as file:
                bot.send_document(message.chat.id, file)
            
            # Удаление файла после отправки
            os.remove(filepath)
        else:
            bot.send_message(
                message.chat.id,
                "Ошибка: файл не был создан.",
                reply_markup=main_menu()
            )
    except Exception as e:
        bot.send_message(
            message.chat.id,
            f"Ошибка экспорта: {e}",
            reply_markup=main_menu()
        )

# Запуск бота
def run_bot():
    bot.polling(none_stop=True)