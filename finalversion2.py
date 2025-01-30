import telebot
import pandas as pd
import os
import pathlib


TOKEN = "7919180675:AAGxtnxxKc5ozTnlkhWH6rRS3NEpn1kKcQk"


bot = telebot.TeleBot(TOKEN)


FILE_PATH = pathlib.Path('C:/Users/nekit/Downloads/s1.xlsx')
FILE_PATH2 = pathlib.Path('C:/Users/nekit/Downloads/s2.xlsx')

#df=pd.read_excel('C:/Users/nekit/Downloads/s2.xlsx')
#print (df.columns)
#value=df.iloc[0:5]
#df = pd.DataFrame({'Unnamed: 3' : [1,2,3,4], 'Unnamed: 5' : [1,2,3,4]})
#print (df)

@bot.message_handler(commands=['update_file'])
def update_file(message):
    if FILE_PATH.exists():
        FILE_PATH.unlink()  
        bot.reply_to(message, "Файл s1.xlsx успешно удалён, отправьте новый файл с таким же названием(s1.xlsx)")
    else:
        bot.reply_to(message, "Файл s1.xlsx не был загружен")

@bot.message_handler(commands=['update_file2'])
def update_file2(message):
    if FILE_PATH2.exists():
        FILE_PATH2.unlink()  
        bot.reply_to(message, "Файл s2.xlsx успешно удалён. Отправьте новый файл с таким же названием(s2.xlsx)")
    else:
        bot.reply_to(message, "Файл s2.xlsx ещё не был загружен.")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    file_id = message.document.file_id
    file_info = bot.get_file(file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    
    #проверка файлов
    if message.document.file_name == "s1.xlsx":
        save_path = FILE_PATH
    elif message.document.file_name == "s2.xlsx":
        save_path = FILE_PATH2
    else:
        bot.reply_to(message, "Некорректное имя файла, загрузите 's1.xlsx' или 's2.xlsx'.")
        return
    
    if save_path.exists():
        save_path.unlink()  
    
    
    with open(save_path, 'wb') as new_file:
        new_file.write(downloaded_file)
    
    bot.reply_to(message, f"Файл {message.document.file_name} успешно загружен и обновлён!")

#Функция анализа за месяц
def analyze_monthly_data():
    df = pd.read_excel(FILE_PATH)
    results = []
    for index, row in df.iloc[1:10].iterrows():
        if pd.isnull(row['Unnamed: 4']):
            row['Unnamed: 4'] = '-'
        if pd.isnull(row['Unnamed: 5']):
            row['Unnamed: 5'] = '-'
        if row['Unnamed: 4'] != '-' and row['Unnamed: 5'] != '-' and row['Unnamed: 4'] != 0:
            completion_percentage = int(row['Unnamed: 5']) / int(row['Unnamed: 4']) * 100
            results.append({
                'Teacher': row['ФИО преподавателя'],
                'Issued': row['Unnamed: 4'],
                'Checked': row['Unnamed: 5'],
                'Completion_Percentage': completion_percentage,
                'Period': 'За месяц'
            })
    return results

#Функция анализа за неделю
def analyze_weekly_data():
    df = pd.read_excel(FILE_PATH)
    results = []
    for index, row in df.iloc[1:10].iterrows():
        if pd.isnull(row['Unnamed: 9']):
            row['Unnamed: 9'] = '-'
        if pd.isnull(row['Unnamed: 10']):
            row['Unnamed: 10'] = '-'
        if row['Unnamed: 9'] != '-' and row['Unnamed: 10'] != '-' and row['Unnamed: 9'] != 0:
            completion_percentage = int(row['Unnamed: 10']) / int(row['Unnamed: 9']) * 100
            results.append({
                'Teacher': row['ФИО преподавателя'],
                'Issued': row['Unnamed: 9'],
                'Checked': row['Unnamed: 10'],
                'Completion_Percentage': completion_percentage,
                'Period': 'За неделю'
            })
    return results

#Функция анализа баллов
def analyze_student_scores():
    df = pd.read_excel(FILE_PATH2)
    results = []
    for index, row in df.iloc[0:453].iterrows():
        if row['Average score'] != '-' and row['Average score'] < 4:
            results.append({
                'FIO': row['FIO'],
                'Average score': row['Average score']
            })
    return results

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, "Это бот, предназначенный для выполнения тз№1, напишите /checkTE6 для 1 тз и /checkTE5 для 5 тз, так же, если хотите заменить таблицы-/update_file для 1 тз и /update_file2 для второго тз")


@bot.message_handler(commands=['checkTE6'])
def send_results(message):
    monthly_results = analyze_monthly_data()
    weekly_results = analyze_weekly_data()
    
    if not monthly_results and not weekly_results:
        bot.send_message(message.chat.id, "Все хорошо, у всех преподавателей выполнение выше 75%.")
        return
    
    if monthly_results:
        response_message = "Результаты за месяц:\n"
        for result in monthly_results:
            response_message += (
                f"Преподаватель: {result['Teacher']}\n"
                f"Получено: {result['Issued']} (За месяц)\n"
                f"Проверено: {result['Checked']} (За месяц)\n"
                f"Процент выполнения: {result['Completion_Percentage']}%\n"
            )
            if result['Completion_Percentage'] < 75:
                response_message += "ПРОЦЕНТ ПРОВЕРКИ НИЖЕ 75!!! За месяц\n"
        bot.send_message(message.chat.id, response_message)
    
    if weekly_results:
        response_message = "Результаты за неделю:\n"
        for result in weekly_results:
            response_message += (
                f"Преподаватель: {result['Teacher']}\n"
                f"Получено: {result['Issued']} (За неделю)\n"
                f"Проверено: {result['Checked']} (За неделю)\n"
                f"Процент выполнения: {result['Completion_Percentage']}%\n"
            )
            if result['Completion_Percentage'] < 75:
                response_message += "ПРОЦЕНТ ПРОВЕРКИ НИЖЕ 75!!! За неделю\n"
        bot.send_message(message.chat.id, response_message)

@bot.message_handler(commands=['checkTE5'])
def send_student_scores(message):
    low_scores = analyze_student_scores()
    if not low_scores:
        response_message = "Все студенты имеют общий балл 3 или выше (по 12-бальной системе)."
    else:
        response_message = "Студенты с общим баллом меньше 3 (по 12-бальной системе):\n"
        for result in low_scores:
            response_message += f"ФИО: {result['FIO']}, Общий балл: {result['Average score']}\n"
    bot.send_message(message.chat.id, response_message)

#запуск бота
print("бот запущен")
bot.polling(none_stop=True)
