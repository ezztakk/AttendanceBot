import telebot
import gspread
from google.oauth2.service_account import Credentials
import datetime
import pandas as pd
from io import BytesIO

# ==================== НАСТРОЙКИ ====================
import os
BOT_TOKEN = os.environ.get('BOT_TOKEN')  # ← ТОКЕН ИЗ ПЕРЕМЕННЫХ
SPREADSHEET_NAME = "Посещаемость студентов"
import os
GOOGLE_KEY_FILE = os.path.join(os.path.dirname(__file__), "google_key.json")
GROUP_NAME = "4231133"  # ← ВАША ГРУППА
# ===================================================

# Расписание пар (ваше точное расписание)
LESSON_TIMES = {
    1: "08:00 - 09:30",
    2: "09:40 - 11:10",
    3: "11:50 - 13:20",
    4: "13:30 - 15:00",
    5: "15:40 - 17:10",
    6: "17:20 - 18:50"
}

# Статусы с эмодзи
STATUSES = {
    'present': {'emoji': '✅', 'text': 'Присутствовал'},
    'absent': {'emoji': '❌', 'text': 'Отсутствовал'},
    'sick': {'emoji': '🤒', 'text': 'Болел'},
    'valid': {'emoji': '📄', 'text': 'Уважительная причина'},
    'other': {'emoji': '❓', 'text': 'Иная причина'}
}

# Настройка доступа к Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']
# Исправление для Python 3.14
try:
    from google.oauth2 import service_account
    
    # Создаём credentials с правильными параметрами
    creds = service_account.Credentials.from_service_account_file(
        GOOGLE_KEY_FILE,
        scopes=scope
    )
    client = gspread.authorize(creds)
    print("✅ Google Таблица подключена!")
except Exception as e:
    print(f"❌ Ошибка подключения к Google: {e}")
    exit()

# Открываем таблицу
try:
    spreadsheet = client.open(SPREADSHEET_NAME)
    attendance_sheet = spreadsheet.worksheet("Посещаемость")
    students_sheet = spreadsheet.worksheet("Студенты")
    print("✅ Google Таблица подключена!")
except Exception as e:
    print(f"❌ Ошибка подключения к Google: {e}")
    exit()

# Создаём бота
bot = telebot.TeleBot(BOT_TOKEN)

# ==================== ХРАНЕНИЕ ТЕКУЩЕГО ВЫБОРА ====================
user_data = {}

def get_user_data(user_id):
    if user_id not in user_data:
        user_data[user_id] = {
            'current_date': datetime.date.today().strftime("%d.%m.%Y"),
            'current_lesson': 1,
            'marking_mode': False
        }
    return user_data[user_id]

# ==================== ГЛАВНОЕ МЕНЮ ====================
@bot.message_handler(commands=['start'])
def start(message):
    user = get_user_data(message.chat.id)
    
    markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btn1 = telebot.types.KeyboardButton('📅 Выбрать дату')
    btn2 = telebot.types.KeyboardButton('🔢 Выбрать пару')
    btn3 = telebot.types.KeyboardButton('📝 Отметить студентов')
    btn4 = telebot.types.KeyboardButton('📊 Получить отчёт')
    btn5 = telebot.types.KeyboardButton('👥 Управление студентами')
    btn6 = telebot.types.KeyboardButton('ℹ️ Текущие настройки')
    markup.add(btn1, btn2, btn3, btn4, btn5, btn6)
    
    time_slot = LESSON_TIMES.get(user['current_lesson'], "")
    
    bot.send_message(message.chat.id,
                    f"👋 *Система учёта посещаемости*\n"
                    f"👥 *Группа:* {GROUP_NAME}\n\n"
                    f"📅 *Текущая дата:* {user['current_date']}\n"
                    f"🔢 *Текущая пара:* {user['current_lesson']}\n"
                    f"⏰ *Время:* {time_slot}\n\n"
                    f"Выберите действие:",
                    parse_mode='Markdown',
                    reply_markup=markup)

# ==================== ВЫБОР ДАТЫ ====================
@bot.message_handler(func=lambda message: message.text == '📅 Выбрать дату')
def choose_date(message):
    user = get_user_data(message.chat.id)
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=3)
    
    today = datetime.date.today()
    
    # Сегодня
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"✅ Сегодня ({today.strftime('%d.%m')})",
            callback_data=f"date_today"
        )
    )
    
    # Вчера
    yesterday = today - datetime.timedelta(days=1)
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"📅 Вчера ({yesterday.strftime('%d.%m')})",
            callback_data=f"date_{yesterday.strftime('%d.%m.%Y')}"
        )
    )
    
    # Завтра
    tomorrow = today + datetime.timedelta(days=1)
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"📅 Завтра ({tomorrow.strftime('%d.%m')})",
            callback_data=f"date_{tomorrow.strftime('%d.%m.%Y')}"
        )
    )
    
    # Другие даты
    for i in range(2, 8):
        other_date = today - datetime.timedelta(days=i)
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"{other_date.strftime('%d.%m')}",
                callback_data=f"date_{other_date.strftime('%d.%m.%Y')}"
            )
        )
    
    # Кнопка для ввода произвольной даты
    markup.add(
        telebot.types.InlineKeyboardButton(
            "📝 Ввести другую дату",
            callback_data="date_custom"
        )
    )
    
    bot.send_message(message.chat.id,
                    f"📅 *Выберите дату:*\n\n"
                    f"Сейчас выбрано: *{user['current_date']}*",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('date_'))
def handle_date_selection(call):
    user = get_user_data(call.message.chat.id)
    
    if call.data == 'date_today':
        new_date = datetime.date.today().strftime("%d.%m.%Y")
        user['current_date'] = new_date
        bot.answer_callback_query(call.id, f"✅ Выбрана сегодняшняя дата")
        
    elif call.data == 'date_custom':
        msg = bot.send_message(call.message.chat.id,
                              "📝 *Введите дату в формате ДД.ММ.ГГГГ*\n"
                              "Пример: 25.03.2024")
        bot.register_next_step_handler(msg, process_custom_date)
        return
    else:
        new_date = call.data[5:]  # Убираем "date_"
        user['current_date'] = new_date
        bot.answer_callback_query(call.id, f"✅ Дата выбрана: {new_date}")
    
    # Обновляем сообщение
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📅 *Дата установлена:* {user['current_date']}\n\n"
             f"Теперь можете выбрать пару или сразу отмечать студентов.",
        parse_mode='Markdown'
    )

def process_custom_date(message):
    user = get_user_data(message.chat.id)
    
    try:
        datetime.datetime.strptime(message.text, "%d.%m.%Y")
        user['current_date'] = message.text
        
        bot.send_message(message.chat.id,
                        f"✅ *Дата установлена:* {message.text}",
                        parse_mode='Markdown')
        
    except ValueError:
        bot.send_message(message.chat.id,
                        "❌ *Неверный формат даты!*\n"
                        "Используйте: ДД.ММ.ГГГГ\n"
                        "Пример: 25.03.2024",
                        parse_mode='Markdown')

# ==================== ВЫБОР ПАРЫ ====================
@bot.message_handler(func=lambda message: message.text == '🔢 Выбрать пару')
def choose_lesson(message):
    user = get_user_data(message.chat.id)
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=3)
    
    # Кнопки для всех пар
    for lesson_num in range(1, 7):
        time_slot = LESSON_TIMES.get(lesson_num, "")
        is_current = "✅ " if lesson_num == user['current_lesson'] else ""
        
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"{is_current}{lesson_num} пара",
                callback_data=f"lesson_{lesson_num}"
            )
        )
    
    bot.send_message(message.chat.id,
                    f"🔢 *Выберите номер пары:*\n\n"
                    f"📅 Дата: {user['current_date']}\n"
                    f"Текущая: {user['current_lesson']} пара\n\n"
                    f"*Расписание:*\n"
                    f"1. {LESSON_TIMES[1]}\n"
                    f"2. {LESSON_TIMES[2]}\n"
                    f"3. {LESSON_TIMES[3]}\n"
                    f"4. {LESSON_TIMES[4]}\n"
                    f"5. {LESSON_TIMES[5]}\n"
                    f"6. {LESSON_TIMES[6]}",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('lesson_'))
def handle_lesson_selection(call):
    user = get_user_data(call.message.chat.id)
    
    lesson_num = int(call.data.split('_')[1])
    user['current_lesson'] = lesson_num
    
    time_slot = LESSON_TIMES.get(lesson_num, "")
    
    bot.answer_callback_query(call.id, f"✅ Выбрана {lesson_num} пара")
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"✅ *Настройки установлены:*\n\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"🔢 *Пара:* {lesson_num}\n"
             f"⏰ *Время:* {time_slot}\n\n"
             f"Теперь можно *отметить студентов* 👇",
        parse_mode='Markdown'
    )

# ==================== ОТМЕТКА СТУДЕНТОВ ====================
@bot.message_handler(func=lambda message: message.text == '📝 Отметить студентов')
def mark_students(message):
    user = get_user_data(message.chat.id)
    
    # Получаем список студентов
    try:
        students = students_sheet.get_all_values()
        if len(students) <= 1:
            bot.send_message(message.chat.id, "❌ Сначала добавьте студентов!")
            return
        
        # Проверяем, есть ли уже отметки на эту дату и пару
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        
        user['marking_mode'] = True
        
        show_students_list(message.chat.id, students[1:], existing_marks)
        
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")

def get_existing_marks(date, lesson):
    """Получаем существующие отметки для даты и пары"""
    try:
        records = attendance_sheet.get_all_records()
        existing_marks = {}
        
        for record in records:
            if (str(record.get('Дата', '')) == date and
                str(record.get('Пара', '')) == str(lesson)):
                
                student_name = record.get('Студент', '')
                status = record.get('Статус', '')
                reason = record.get('Причина', '')
                if student_name and status:
                    existing_marks[student_name] = {
                        'status': status,
                        'reason': reason
                    }
        return existing_marks
    except:
        return {}

def show_students_list(chat_id, students, existing_marks):
    user = get_user_data(chat_id)
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=3)
    
    # Заголовок с информацией
    time_slot = LESSON_TIMES.get(user['current_lesson'], "")
    
    # Кнопки быстрой отметки всех
    markup.add(
        telebot.types.InlineKeyboardButton(
            "✅ Отметить всех присутствующими",
            callback_data="mark_all_present"
        )
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton(
            "❌ Отметить всех отсутствующими",
            callback_data="mark_all_absent"
        )
    )
    
    # Разделитель
    markup.add(
        telebot.types.InlineKeyboardButton(
            "─" * 20,
            callback_data="no_action"
        )
    )
    
    # Список студентов с текущими статусами
    for student in students:
        if len(student) >= 2:
            student_name = student[1]
            
            # Получаем текущий статус
            if student_name in existing_marks:
                status_info = existing_marks[student_name]
                status_text = status_info['status']
                
                # Находим эмодзи для статуса
                status_emoji = '❓'
                for status_code, info in STATUSES.items():
                    if info['text'] == status_text:
                        status_emoji = info['emoji']
                        break
                
                # Если есть причина, добавляем значок
                if status_info.get('reason') and status_info['reason'] != '-':
                    status_emoji = f"{status_emoji}📝"
            else:
                status_emoji = '⬜'  # Не отмечен
            
            # Сокращаем длинное имя
            display_name = student_name
            if len(display_name) > 15:
                display_name = display_name[:12] + "..."
            
            markup.add(
                telebot.types.InlineKeyboardButton(
                    f"{status_emoji} {display_name}",
                    callback_data=f"student_{student_name}"
                )
            )
    
    # Кнопки сохранения
    markup.add(
        telebot.types.InlineKeyboardButton(
            "💾 Сохранить и выйти",
            callback_data="save_exit"
        ),
        telebot.types.InlineKeyboardButton(
            "🔄 Обновить список",
            callback_data="refresh_list"
        )
    )
    
    bot.send_message(chat_id,
                    f"📝 *Отметка посещаемости*\n\n"
                    f"👥 *Группа:* {GROUP_NAME}\n"
                    f"📅 *Дата:* {user['current_date']}\n"
                    f"🔢 *Пара:* {user['current_lesson']} ({time_slot})\n\n"
                    f"*Статусы:*\n"
                    f"✅ - присутствовал\n"
                    f"❌ - отсутствовал\n"
                    f"🤒 - болел\n"
                    f"📄 - уважительная причина\n"
                    f"❓ - иная причина\n"
                    f"⬜ - не отмечен\n"
                    f"📝 - с комментарием\n\n"
                    f"*Выберите студента:*",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('student_'))
def mark_single_student(call):
    user = get_user_data(call.message.chat.id)
    student_name = call.data.split('_', 1)[1]
    
    # Получаем текущий статус студента
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    current_status = None
    current_reason = None
    
    if student_name in existing_marks:
        current_status = existing_marks[student_name]['status']
        current_reason = existing_marks[student_name].get('reason', '-')
    
    # Создаём клавиатуру с выбором статуса
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    for status_code, info in STATUSES.items():
        is_current = "✅ " if info['text'] == current_status else ""
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"{is_current}{info['emoji']} {info['text']}",
                callback_data=f"set_{student_name}_{status_code}"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton(
            "↩️ Назад к списку",
            callback_data="back_to_list"
        )
    )
    
    status_info = ""
    if current_status:
        for status_code, info in STATUSES.items():
            if info['text'] == current_status:
                status_info = f"\n📊 *Текущий статус:* {info['emoji']} {current_status}"
                if current_reason and current_reason != '-':
                    status_info += f"\n📝 *Причина:* {current_reason}"
                break
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📝 *Выберите статус для студента:*\n\n"
             f"👤 *{student_name}*\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"🔢 *Пара:* {user['current_lesson']}"
             f"{status_info}",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data.startswith('set_'))
def set_student_status(call):
    user = get_user_data(call.message.chat.id)
    
    # Формат: set_Иванов Алексей_present
    data = call.data.split('_', 2)
    student_name = data[1]
    status_code = data[2]
    
    info = STATUSES[status_code]
    
    # Если статус требует причины, запрашиваем её
    if status_code in ['sick', 'valid', 'other']:
        # Сохраняем временные данные
        user['temp_data'] = {
            'student_name': student_name,
            'status_code': status_code,
            'status_text': info['text'],
            'callback_message_id': call.message.message_id
        }
        
        msg = bot.send_message(call.message.chat.id,
                              f"📝 *Введите причину для {student_name}:*\n"
                              f"Статус: {info['emoji']} {info['text']}\n\n"
                              f"Пример: 'Болел ОРВИ', 'Справка от врача', 'Семейные обстоятельства'")
        bot.register_next_step_handler(msg, save_with_reason)
        return
    
    # Для простых статусов сохраняем сразу
    save_attendance_record(user['current_date'], user['current_lesson'], 
                          student_name, info['text'], "-")
    
    bot.answer_callback_query(call.id, f"✅ {student_name}: {info['text']}")
    
    # Обновляем список
    refresh_students_list(call.message.chat.id, call.message.message_id)

def save_with_reason(message):
    user = get_user_data(message.chat.id)
    
    if not user.get('temp_data'):
        bot.send_message(message.chat.id, "❌ Ошибка: данные не найдены")
        return
    
    temp_data = user['temp_data']
    reason = message.text
    
    save_attendance_record(user['current_date'], user['current_lesson'], 
                          temp_data['student_name'], temp_data['status_text'], reason)
    
    info = STATUSES[temp_data['status_code']]
    
    bot.send_message(message.chat.id,
                    f"✅ *{temp_data['student_name']} отмечен*\n"
                    f"📝 *Причина:* {reason}",
                    parse_mode='Markdown')
    
    # Очищаем временные данные
    user['temp_data'] = None
    
    # Обновляем список
    refresh_students_list(message.chat.id)

def save_attendance_record(date, lesson, student, status, reason):
    """Сохраняет запись о посещении"""
    try:
        # Сначала удаляем старые записи для этого студента на эту дату и пару
        records = attendance_sheet.get_all_values()
        
        rows_to_delete = []
        for i, row in enumerate(records):
            if (i > 0 and len(row) >= 4 and
                str(row[0]) == date and
                str(row[1]) == str(lesson) and
                str(row[3]) == student):
                rows_to_delete.append(i + 1)
        
        for row_num in sorted(rows_to_delete, reverse=True):
            attendance_sheet.delete_rows(row_num)
        
        # Добавляем новую запись
        time_now = datetime.datetime.now().strftime("%H:%M")
        
        attendance_sheet.append_row([
            date,
            lesson,
            GROUP_NAME,
            student,
            status,
            reason,
            time_now
        ])
        
        return True
    except Exception as e:
        print(f"Ошибка сохранения: {e}")
        return False

@bot.callback_query_handler(func=lambda call: call.data == 'back_to_list')
def back_to_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

@bot.callback_query_handler(func=lambda call: call.data == 'refresh_list')
def refresh_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

@bot.callback_query_handler(func=lambda call: call.data in ['mark_all_present', 'mark_all_absent'])
def mark_all_students(call):
    user = get_user_data(call.message.chat.id)
    
    status_code = 'present' if call.data == 'mark_all_present' else 'absent'
    info = STATUSES[status_code]
    
    # Получаем список студентов
    try:
        students = students_sheet.get_all_values()
        
        for student in students[1:]:
            if len(student) >= 2:
                student_name = student[1]
                save_attendance_record(user['current_date'], user['current_lesson'], 
                                      student_name, info['text'], "-")
        
        bot.answer_callback_query(call.id, f"✅ Все студенты отмечены как {info['text']}")
        refresh_students_list(call.message.chat.id, call.message.message_id)
        
    except Exception as e:
        bot.answer_callback_query(call.id, f"❌ Ошибка: {e}")

def refresh_students_list(chat_id, message_id=None):
    """Обновляет список студентов"""
    user = get_user_data(chat_id)
    
    try:
        students = students_sheet.get_all_values()
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        
        if message_id:
            # Удаляем старое сообщение
            try:
                bot.delete_message(chat_id, message_id)
            except:
                pass
        
        show_students_list(chat_id, students[1:], existing_marks)
        
    except Exception as e:
        bot.send_message(chat_id, f"❌ Ошибка обновления: {e}")

@bot.callback_query_handler(func=lambda call: call.data == 'save_exit')
def save_and_exit(call):
    user = get_user_data(call.message.chat.id)
    user['marking_mode'] = False
    
    bot.answer_callback_query(call.id, "✅ Данные сохранены")
    
    time_slot = LESSON_TIMES.get(user['current_lesson'], "")
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"✅ *Данные сохранены!*\n\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"🔢 *Пара:* {user['current_lesson']} ({time_slot})\n"
             f"👥 *Группа:* {GROUP_NAME}\n\n"
             f"Для нового действия нажмите /start",
        parse_mode='Markdown'
    )

# ==================== УПРАВЛЕНИЕ СТУДЕНТАМИ ====================
@bot.message_handler(func=lambda message: message.text == '👥 Управление студентами')
def manage_students(message):
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    markup.add(
        telebot.types.InlineKeyboardButton("➕ Добавить студента", callback_data="add_student"),
        telebot.types.InlineKeyboardButton("🗑️ Удалить студента", callback_data="delete_student"),
        telebot.types.InlineKeyboardButton("📋 Список студентов", callback_data="list_students"),
        telebot.types.InlineKeyboardButton("📤 Импорт из файла", callback_data="import_students")
    )
    
    bot.send_message(message.chat.id,
                    "👥 *Управление списком студентов*\n\n"
                    f"Группа: *{GROUP_NAME}*",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'list_students')
def list_students(call):
    try:
        students = students_sheet.get_all_values()
        
        if len(students) <= 1:
            bot.answer_callback_query(call.id, "📭 Список студентов пуст")
            return
        
        response = f"📋 *Список студентов ({GROUP_NAME}):*\n\n"
        
        for i, student in enumerate(students[1:], 1):
            if len(student) >= 2:
                response += f"{i}. {student[1]}\n"
        
        bot.send_message(call.message.chat.id, response, parse_mode='Markdown')
        
    except Exception as e:
        bot.answer_callback_query(call.id, f"❌ Ошибка: {e}")

@bot.callback_query_handler(func=lambda call: call.data == 'add_student')
def add_student(call):
    msg = bot.send_message(call.message.chat.id,
                          "📝 *Добавление студента*\n\n"
                          "Введите Фамилию и Имя студента:\n\n"
                          "*Пример:*\n"
                          "Иванов Алексей")
    bot.register_next_step_handler(msg, save_new_student)

def save_new_student(message):
    try:
        name = message.text.strip()
        
        if not name:
            bot.send_message(message.chat.id, "❌ Имя не может быть пустым!")
            return
        
        # Проверяем, есть ли уже такой студент
        students = students_sheet.get_all_values()
        for student in students[1:]:
            if len(student) >= 2 and student[1] == name:
                bot.send_message(message.chat.id, f"⚠️ Студент '{name}' уже есть в списке!")
                return
        
        # Добавляем студента
        students_sheet.append_row([GROUP_NAME, name])
        
        bot.send_message(message.chat.id,
                        f"✅ *Студент добавлен!*\n\n"
                        f"👤 *{name}*\n"
                        f"👥 *Группа:* {GROUP_NAME}",
                        parse_mode='Markdown')
        
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")

# ==================== ОТЧЁТЫ ====================
@bot.message_handler(func=lambda message: message.text == '📊 Получить отчёт')
def get_report_menu(message):
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    markup.add(
        telebot.types.InlineKeyboardButton("📅 Отчёт за месяц", callback_data="report_month"),
        telebot.types.InlineKeyboardButton("📆 Отчёт за период", callback_data="report_period"),
        telebot.types.InlineKeyboardButton("👤 Отчёт по студенту", callback_data="report_student"),
        telebot.types.InlineKeyboardButton("📊 Общая статистика", callback_data="report_stats")
    )
    
    bot.send_message(message.chat.id,
                    "📊 *Выберите тип отчёта:*",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'report_month')
def ask_month_for_report(call):
    current_month = datetime.date.today().strftime("%m.%Y")
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📅 *За какой месяц нужен отчёт?*\n\n"
             f"Введите месяц и год в формате:\n"
             f"`ММ.ГГГГ`\n\n"
             f"*Пример:* `{current_month}`\n"
             f"Или введите `текущий` для текущего месяца",
        parse_mode='Markdown'
    )
    
    bot.register_next_step_handler_by_chat_id(call.message.chat.id, generate_monthly_report)

def generate_monthly_report(message):
    try:
        # Определяем месяц
        if message.text.lower() == 'текущий':
            month_year = datetime.date.today().strftime("%m.%Y")
        else:
            month_year = message.text
        
        month, year = map(int, month_year.split('.'))
        
        # Получаем все записи
        records = attendance_sheet.get_all_records()
        
        if not records:
            bot.send_message(message.chat.id, "📭 Нет данных для отчёта")
            return
        
        # Создаём DataFrame
        df = pd.DataFrame(records)
        
        # Преобразуем дату
        df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
        
        # Фильтруем по месяцу и году
        mask = (df['Дата'].dt.month == month) & (df['Дата'].dt.year == year)
        filtered = df[mask]
        
        if filtered.empty:
            bot.send_message(message.chat.id, f"📭 Нет данных за {month_year}")
            return
        
        # Создаём Excel файл
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. Детальная посещаемость
            filtered.to_excel(writer, sheet_name='Посещаемость', index=False)
            
            # 2. Сводка по студентам
            student_stats = []
            
            for student, group_df in filtered.groupby('Студент'):
                total = len(group_df)
                present = len(group_df[group_df['Статус'] == 'Присутствовал'])
                absent = len(group_df[group_df['Статус'] == 'Отсутствовал'])
                sick = len(group_df[group_df['Статус'] == 'Болел'])
                valid = len(group_df[group_df['Статус'] == 'Уважительная причина'])
                other = len(group_df[group_df['Статус'] == 'Иная причина'])
                
                attendance_rate = (present / total * 100) if total > 0 else 0
                
                student_stats.append({
                    'Студент': student,
                    'Всего занятий': total,
                    'Присутствовал': present,
                    'Отсутствовал': absent,
                    'Болел': sick,
                    'Уважительно': valid,
                    'Иные причины': other,
                    '% посещения': round(attendance_rate, 1)
                })
            
            stats_df = pd.DataFrame(student_stats)
            stats_df.to_excel(writer, sheet_name='Статистика', index=False)
            
            # 3. Причины пропусков
            reasons_df = filtered[filtered['Причина'] != '-']
            if not reasons_df.empty:
                reasons_df.to_excel(writer, sheet_name='Причины пропусков', index=False)
            
            # 4. Общая статистика
            summary = {
                'Параметр': ['Всего занятий', 'Всего студентов', 'Средний % посещения',
                            'Всего пропусков', 'По болезни', 'Уважительно', 'Неуважительно'],
                'Значение': [
                    filtered['Дата'].nunique(),
                    filtered['Студент'].nunique(),
                    round(stats_df['% посещения'].mean(), 1),
                    len(filtered) - len(filtered[filtered['Статус'] == 'Присутствовал']),
                    len(filtered[filtered['Статус'] == 'Болел']),
                    len(filtered[filtered['Статус'] == 'Уважительная причина']),
                    len(filtered[filtered['Статус'] == 'Отсутствовал'])
                ]
            }
            
            summary_df = pd.DataFrame(summary)
            summary_df.to_excel(writer, sheet_name='Итоги', index=False)
        
        output.seek(0)
        
        # Формируем текстовую сводку
        total_classes = filtered['Дата'].nunique()
        total_students = filtered['Студент'].nunique()
        total_present = len(filtered[filtered['Статус'] == 'Присутствовал'])
        
        caption = (f"📊 *Отчёт за {month_year}*\n\n"
                  f"👥 Группа: {GROUP_NAME}\n"
                  f"📅 Занятий: {total_classes}\n"
                  f"👤 Студентов: {total_students}\n"
                  f"✅ Присутствовали: {total_present}\n"
                  f"📈 % посещения: {round(total_present/len(filtered)*100, 1)}%")
        
        # Отправляем файл
        bot.send_chat_action(message.chat.id, 'upload_document')
        
        bot.send_document(
            message.chat.id,
            output,
            caption=caption,
            parse_mode='Markdown',
            visible_file_name=f'посещаемость_{GROUP_NAME}_{month_year}.xlsx'
        )
        
    except ValueError:
        bot.send_message(message.chat.id, "❌ Неправильный формат! Используйте ММ.ГГГГ")
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка генерации отчёта: {str(e)}")

# ==================== ТЕКУЩИЕ НАСТРОЙКИ ====================
@bot.message_handler(func=lambda message: message.text == 'ℹ️ Текущие настройки')
def show_current_settings(message):
    user = get_user_data(message.chat.id)
    time_slot = LESSON_TIMES.get(user['current_lesson'], "")
    
    # Получаем количество студентов
    try:
        students = students_sheet.get_all_values()
        student_count = max(0, len(students) - 1)
    except:
        student_count = 0
    
    bot.send_message(message.chat.id,
                    f"⚙️ *Текущие настройки:*\n\n"
                    f"👥 *Группа:* {GROUP_NAME}\n"
                    f"👤 *Студентов:* {student_count}\n\n"
                    f"📅 *Дата:* {user['current_date']}\n"
                    f"🔢 *Пара:* {user['current_lesson']}\n"
                    f"⏰ *Время:* {time_slot}\n\n"
                    f"*Изменить:*\n"
                    f"📅 - выбрать дату\n"
                    f"🔢 - выбрать пару\n"
                    f"📝 - отметить студентов",
                    parse_mode='Markdown')

# ==================== ЗАПУСК ====================
if __name__ == "__main__":
    print("=" * 50)
    print(f"🤖 Бот для учёта посещаемости запущен!")
    print(f"📍 Группа: {GROUP_NAME}")
    print(f"📅 Расписание пар:")
    for i in range(1, 7):
        print(f"   {i}. {LESSON_TIMES[i]}")
    print("=" * 50)
    
    try:
        bot.polling(none_stop=True, interval=0)
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import time
        time.sleep(10)