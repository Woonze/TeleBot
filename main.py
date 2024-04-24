import telebot
from telebot import types
import datetime
import openpyxl

bot = telebot.TeleBot('YOUR_BOT');

EXCEL_FILE_PATH = '/Users/woonze/desktop/Antonio.xlsx'

@bot.message_handler(commands=["start"])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    but1 = types.KeyboardButton("📆Расписание📆")
    markup.add(but1)

    bot.send_message(message.chat.id, "Я сигма-бот Агент Кей, также известный как Райан Гослинг! У меня есть данные о твоем расписании, тычь кнопку внизу.", reply_markup=markup)
    bot.send_sticker(message.chat.id, sticker='CAACAgIAAxkBAAJFjWXngk1NDA11GxTpscX10Br64hYzAALuGwACQ18JSWTDRcWxjp0vNAQ')

@bot.message_handler(commands=["time"])
def get_time(message):
    week_number = datetime.datetime.now().isocalendar()[1]
    if week_number % 2 == 0:
        week_type = "чет"
    else:
        week_type = "нечет"

    day_of_week = datetime.datetime.now().strftime("%A")
    if day_of_week == "Monday":
        day_of_week = "понедельник"
    elif day_of_week == "Tuesday":
        day_of_week = "вторник"
    elif day_of_week == "Wednesday":
        day_of_week = "среда"
    elif day_of_week == "Thursday":
        day_of_week = "четверг"
    elif day_of_week == "Friday":
        day_of_week = "пятница"
    elif day_of_week == "Saturday":
        day_of_week = "суббота"
    elif day_of_week == "Sunday":
        day_of_week = "воскресенье"
    # bot.reply_to(message, f"Сегодня {day_of_week}")

    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb.active
    if week_type == "нечет":
        if day_of_week == "понедельник":
            data1 = sheet['C1'].value
            data2 = sheet['C2'].value
            data3 = sheet['C3'].value
            data = f"\n{data1}\n{data2}\n{data3}"
        elif day_of_week == "вторник":
            data1 = sheet['C5'].value
            data2 = sheet['C6'].value
            data3 = sheet['C7'].value
            data = f"\n{data1}\n{data2}\n{data3}"
        elif day_of_week == "среда":
            data1 = sheet['C9'].value
            data2 = sheet['C10'].value
            data = (f"\n{data1}\n{data2}")
        elif day_of_week == "четверг":
            data1 = sheet['C12'].value
            data2 = sheet['C13'].value
            data3 = sheet['C14'].value
            data4 = sheet['C15'].value
            data = f"\n{data1}\n{data2}\n{data3}\n{data4}"
        elif day_of_week == "пятница":
            data1 = sheet['C17'].value
            data = f"\n{data1}"
        elif day_of_week == "суббота":
            data1 = sheet['C19'].value
            data = f"\n{data1}"
        elif day_of_week == "воскресенье":
            data1 = sheet['C21'].value
            data = f"\n{data1}"
    else:
        if day_of_week == "понедельник":
            data1 = sheet['G1'].value
            data2 = sheet['G2'].value
            data3 = sheet['G3'].value
            data = f"\n{data1}\n{data2}\n{data3}"
        elif day_of_week == "вторник":
            data1 = sheet['G5'].value
            data2 = sheet['G6'].value
            data3 = sheet['G7'].value
            data4 = sheet['G8'].value
            data = f"\n{data1}\n{data2}\n{data3}\n{data4}"
        elif day_of_week == "среда":
            data1 = sheet['G10'].value
            data = (f"\n{data1}")
        elif day_of_week == "четверг":
            data1 = sheet['G12'].value
            data2 = sheet['G13'].value
            data3 = sheet['G14'].value
            data = f"\n{data1}\n{data2}\n{data3}"
        elif day_of_week == "пятница":
            data1 = sheet['G17'].value
            data = f"\n{data1}"
        elif day_of_week == "суббота":
            data1 = sheet['G19'].value
            data = f"\n{data1}"
        elif day_of_week == "воскресенье":
            data1 = sheet['G21'].value
            data = f"\n{data1}"
    bot.reply_to(message, f"Сегодня {day_of_week}, расписание: \n{data}")

    wb.close()

@bot.message_handler(content_types="text")
def bot_massage(message):
    if message.chat.type == "private":
        if message.text == "📆Расписание📆":
            week_number = datetime.datetime.now().isocalendar()[1]
            if week_number % 2 == 0:
                week_type = "нечет"
            else:
                week_type = "чет"

            day_of_week = datetime.datetime.now().strftime("%A")
            if day_of_week == "Monday":
                day_of_week = "понедельник"
            elif day_of_week == "Tuesday":
                day_of_week = "вторник"
            elif day_of_week == "Wednesday":
                day_of_week = "среда"
            elif day_of_week == "Thursday":
                day_of_week = "четверг"
            elif day_of_week == "Friday":
                day_of_week = "пятница"
            elif day_of_week == "Saturday":
                day_of_week = "суббота"
            elif day_of_week == "Sunday":
                day_of_week = "воскресенье"
            # bot.reply_to(message, f"Сегодня {day_of_week}")

            wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
            sheet = wb.active
            if week_type == "нечет":
                if day_of_week == "понедельник":
                    data1 = sheet['C1'].value
                    data2 = sheet['C2'].value
                    data3 = sheet['C3'].value
                    data = f"\n{data1}\n{data2}\n{data3}"
                elif day_of_week == "вторник":
                    data1 = sheet['C5'].value
                    data2 = sheet['C6'].value
                    data3 = sheet['C7'].value
                    data = f"\n{data1}\n{data2}\n{data3}"
                elif day_of_week == "среда":
                    data1 = sheet['C9'].value
                    data2 = sheet['C10'].value
                    data = (f"\n{data1}\n{data2}")
                elif day_of_week == "четверг":
                    data1 = sheet['C12'].value
                    data2 = sheet['C13'].value
                    data3 = sheet['C14'].value
                    data4 = sheet['C15'].value
                    data = f"\n{data1}\n{data2}\n{data3}\n{data4}"
                elif day_of_week == "пятница":
                    data1 = sheet['C17'].value
                    data = f"\n{data1}"
                elif day_of_week == "суббота":
                    data1 = sheet['C19'].value
                    data = f"\n{data1}"
                elif day_of_week == "воскресенье":
                    data1 = sheet['C21'].value
                    data = f"\n{data1}"
            else:
                if day_of_week == "понедельник":
                    data1 = sheet['G1'].value
                    data2 = sheet['G2'].value
                    data3 = sheet['G3'].value
                    data = f"\n{data1}\n{data2}\n{data3}"
                elif day_of_week == "вторник":
                    data1 = sheet['G5'].value
                    data2 = sheet['G6'].value
                    data3 = sheet['G7'].value
                    data4 = sheet['G8'].value
                    data = f"\n{data1}\n{data2}\n{data3}\n{data4}"
                elif day_of_week == "среда":
                    data1 = sheet['G10'].value
                    data = (f"\n{data1}")
                elif day_of_week == "четверг":
                    data1 = sheet['G12'].value
                    data2 = sheet['G13'].value
                    data3 = sheet['G14'].value
                    data = f"\n{data1}\n{data2}\n{data3}"
                elif day_of_week == "пятница":
                    data1 = sheet['G17'].value
                    data = f"\n{data1}"
                elif day_of_week == "суббота":
                    data1 = sheet['G19'].value
                    data = f"\n{data1}"
                elif day_of_week == "воскресенье":
                    data1 = sheet['G21'].value
                    data = f"\n{data1}"
            bot.reply_to(message, f"Сегодня {day_of_week}, расписание: \n{data}")

            wb.close()

bot.polling(none_stop=True, interval=0)
