import telebot
import gspread
from google.oauth2.service_account import Credentials
import datetime
import pandas as pd
from io import BytesIO
import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

# ==================== НАСТРОЙКИ ====================
BOT_TOKEN = os.environ.get('BOT_TOKEN')
SPREADSHEET_NAME = "Посещаемость студентов"
GOOGLE_KEY_FILE = os.path.join(os.path.dirname(__file__), "google_key.json")
GROUP_NAME = "4231133"

# Типы неуважительных пропусков (только они считаются прогулами)
UNRESPECTFUL_STATUSES = ['Отсутствовал']  # ❌

# Количество студентов на одной странице
ITEMS_PER_PAGE = 10
# ===================================================

# Расписание пар
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

try:
    from google.oauth2 import service_account
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
            'marking_mode': False,
            'current_page': 0,
            'students_list': [],
            'selected_students': set()
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
    btn5 = telebot.types.KeyboardButton('ℹ️ Текущие настройки')
    markup.add(btn1, btn2, btn3, btn4, btn5)
    
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
    
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"✅ Сегодня ({today.strftime('%d.%m')})",
            callback_data=f"date_today"
        )
    )
    
    yesterday = today - datetime.timedelta(days=1)
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"📅 Вчера ({yesterday.strftime('%d.%m')})",
            callback_data=f"date_{yesterday.strftime('%d.%m.%Y')}"
        )
    )
    
    tomorrow = today + datetime.timedelta(days=1)
    markup.add(
        telebot.types.InlineKeyboardButton(
            f"📅 Завтра ({tomorrow.strftime('%d.%m')})",
            callback_data=f"date_{tomorrow.strftime('%d.%m.%Y')}"
        )
    )
    
    for i in range(2, 8):
        other_date = today - datetime.timedelta(days=i)
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"{other_date.strftime('%d.%m')}",
                callback_data=f"date_{other_date.strftime('%d.%m.%Y')}"
            )
        )
    
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
        new_date = call.data[5:]
        user['current_date'] = new_date
        bot.answer_callback_query(call.id, f"✅ Дата выбрана: {new_date}")
    
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

# ==================== ПОЛУЧЕНИЕ СУЩЕСТВУЮЩИХ ОТМЕТОК ====================
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

# ==================== СОХРАНЕНИЕ ЗАПИСИ ====================
def save_attendance_record(date, lesson, student, status, reason):
    """Сохраняет запись о посещении"""
    try:
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

# ==================== ОТМЕТКА СТУДЕНТОВ С ЧЕКБОКСАМИ ====================
def show_students_list_with_checkboxes(chat_id, students, existing_marks, page=None):
    """Показывает список студентов с чекбоксами для множественного выбора"""
    user = get_user_data(chat_id)
    
    if 'selected_students' not in user:
        user['selected_students'] = set()
    
    if page is None:
        page = user.get('current_page', 0)
    else:
        user['current_page'] = page
    
    total_students = len(students)
    total_pages = (total_students + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    
    if total_pages == 0:
        page = 0
    elif page < 0:
        page = 0
    elif page >= total_pages:
        page = total_pages - 1
    user['current_page'] = page
    
    start = page * ITEMS_PER_PAGE
    end = min(start + ITEMS_PER_PAGE, total_students)
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    time_slot = LESSON_TIMES.get(user['current_lesson'], "")
    
    selected_count = len(user['selected_students'])
    if selected_count > 0:
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"✅ ПРИМЕНИТЬ К ВЫБРАННЫМ ({selected_count})",
                callback_data="apply_to_selected"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Все присутствуют", callback_data="mark_all_present"),
        telebot.types.InlineKeyboardButton("❌ Все отсутствуют", callback_data="mark_all_absent")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("🤒 Все болеют", callback_data="mark_all_sick"),
        telebot.types.InlineKeyboardButton("📄 Все уважительная", callback_data="mark_all_valid")
    )
    
    for idx_in_list in range(start, end):
        student = students[idx_in_list]
        if len(student) >= 2:
            student_name = student[1]
            
            if student_name in existing_marks:
                status_info = existing_marks[student_name]
                status_text = status_info['status']
                status_emoji = '❓'
                for code, info in STATUSES.items():
                    if info['text'] == status_text:
                        status_emoji = info['emoji']
                        break
                if status_info.get('reason') and status_info['reason'] != '-':
                    status_emoji = f"{status_emoji}📝"
            else:
                status_emoji = '⬜'
            
            checkbox = "☑️" if idx_in_list in user['selected_students'] else "◻️"
            
            display_name = student_name
            if len(display_name) > 12:
                display_name = display_name[:12] + "…"
            
            markup.add(
                telebot.types.InlineKeyboardButton(
                    f"{checkbox} {status_emoji} {display_name}",
                    callback_data=f"toggle_{idx_in_list}"
                )
            )
    
    nav_buttons = []
    if page > 0:
        nav_buttons.append(telebot.types.InlineKeyboardButton("◀ Предыдущая", callback_data="page_prev"))
    if page < total_pages - 1:
        nav_buttons.append(telebot.types.InlineKeyboardButton("Следующая ▶", callback_data="page_next"))
    if nav_buttons:
        markup.add(*nav_buttons)
    
    page_info = f"📄 Страница {page+1} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
    
    markup.add(
        telebot.types.InlineKeyboardButton("❌ Снять все выборы", callback_data="clear_selection"),
        telebot.types.InlineKeyboardButton("🔄 Обновить", callback_data="refresh_list")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("💾 СОХРАНИТЬ И ВЫЙТИ", callback_data="save_exit")
    )
    
    selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
    
    bot.send_message(
        chat_id,
        f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
        f"👥 *Группа:* {GROUP_NAME}\n"
        f"📅 *Дата:* {user['current_date']}\n"
        f"🔢 *Пара:* {user['current_lesson']} ({time_slot})\n"
        f"{selected_text}"
        f"{page_info}\n\n"
        f"*Как отмечать:*\n"
        f"1. Нажмите на студента, чтобы выбрать ☑️\n"
        f"2. Выберите статус для ВСЕХ выбранных\n"
        f"3. Или отметьте всю группу сразу\n\n"
        f"*Статусы:* ✅ ❌ 🤒 📄 ❓\n"
        f"*⬜ - не отмечен, 📝 - есть причина*",
        parse_mode='Markdown',
        reply_markup=markup
    )

# ==================== ОБРАБОТЧИКИ ДЛЯ ОТМЕТКИ ====================
@bot.message_handler(func=lambda message: message.text == '📝 Отметить студентов')
def mark_students(message):
    user = get_user_data(message.chat.id)
    
    try:
        students = students_sheet.get_all_values()
        if len(students) <= 1:
            bot.send_message(message.chat.id, "❌ Сначала добавьте студентов!")
            return
        
        user['students_list'] = students[1:]
        user['selected_students'] = set()
        user['current_page'] = 0
        
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        user['marking_mode'] = True
        
        show_students_list_with_checkboxes(message.chat.id, students[1:], existing_marks, 0)
        
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")

@bot.callback_query_handler(func=lambda call: call.data.startswith('toggle_'))
def toggle_student(call):
    """Выбор/снятие выбора студента"""
    user = get_user_data(call.message.chat.id)
    idx = int(call.data.split('_')[1])
    
    if idx in user['selected_students']:
        user['selected_students'].remove(idx)
        bot.answer_callback_query(call.id, "❌ Выбор снят")
    else:
        user['selected_students'].add(idx)
        bot.answer_callback_query(call.id, "✅ Студент выбран")
    
    students = user.get('students_list', [])
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    
    try:
        bot.delete_message(call.message.chat.id, call.message.message_id)
    except:
        pass
    
    show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'clear_selection')
def clear_selection(call):
    """Снять все выборы"""
    user = get_user_data(call.message.chat.id)
    user['selected_students'] = set()
    bot.answer_callback_query(call.id, "❌ Все выборы сняты")
    
    students = user.get('students_list', [])
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    
    try:
        bot.delete_message(call.message.chat.id, call.message.message_id)
    except:
        pass
    
    show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'apply_to_selected')
def apply_to_selected(call):
    """Применить статус к выбранным студентам"""
    user = get_user_data(call.message.chat.id)
    
    if not user.get('selected_students'):
        bot.answer_callback_query(call.id, "❌ Нет выбранных студентов")
        return
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    for status_code, info in STATUSES.items():
        markup.add(
            telebot.types.InlineKeyboardButton(
                f"{info['emoji']} {info['text']}",
                callback_data=f"apply_status_{status_code}"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton("↩️ Назад", callback_data="back_to_list")
    )
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📝 *Применить статус к выбранным студентам*\n\n"
             f"✅ *Выбрано:* {len(user['selected_students'])} студентов\n\n"
             f"*Выберите статус:*",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data.startswith('apply_status_'))
def apply_status_to_selected(call):
    """Применяет выбранный статус ко всем отмеченным студентам"""
    user = get_user_data(call.message.chat.id)
    status_code = call.data.split('_')[2]
    info = STATUSES[status_code]
    
    if not user.get('selected_students') or not user.get('students_list'):
        bot.answer_callback_query(call.id, "❌ Нет выбранных студентов")
        return
    
    if status_code in ['sick', 'valid', 'other']:
        user['pending_status'] = {
            'status_code': status_code,
            'status_text': info['text'],
            'students': list(user['selected_students']).copy(),
            'callback_message_id': call.message.message_id
        }
        
        msg = bot.send_message(
            call.message.chat.id,
            f"📝 *Введите причину для {len(user['selected_students'])} студентов:*\n"
            f"Статус: {info['emoji']} {info['text']}\n\n"
            f"Причина будет применена ко всем выбранным студентам."
        )
        bot.register_next_step_handler(msg, save_reason_for_selected)
        return
    else:
        for idx in user['selected_students']:
            if idx < len(user['students_list']):
                student_name = user['students_list'][idx][1]
                save_attendance_record(
                    user['current_date'], 
                    user['current_lesson'], 
                    student_name, 
                    info['text'], 
                    "-"
                )
    
    user['selected_students'] = set()
    bot.answer_callback_query(call.id, f"✅ Отмечено {len(user['selected_students'])} студентов")
    
    students = user.get('students_list', [])
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    
    try:
        bot.delete_message(call.message.chat.id, call.message.message_id)
    except:
        pass
    
    show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, user['current_page'])

def save_reason_for_selected(message):
    """Сохраняет причину для всех выбранных студентов"""
    user = get_user_data(message.chat.id)
    reason = message.text
    
    if 'pending_status' not in user:
        bot.send_message(message.chat.id, "❌ Ошибка: данные не найдены")
        return
    
    pending = user['pending_status']
    
    for idx in pending['students']:
        if idx < len(user['students_list']):
            student_name = user['students_list'][idx][1]
            save_attendance_record(
                user['current_date'],
                user['current_lesson'],
                student_name,
                pending['status_text'],
                reason
            )
    
    user['selected_students'] = set()
    del user['pending_status']
    
    bot.send_message(
        message.chat.id,
        f"✅ *Отмечено {len(pending['students'])} студентов*\n"
        f"📝 *Причина:* {reason}"
    )
    
    students = user.get('students_list', [])
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    show_students_list_with_checkboxes(message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data in ['mark_all_present', 'mark_all_absent'])
def mark_all_students(call):
    user = get_user_data(call.message.chat.id)
    
    status_code = 'present' if call.data == 'mark_all_present' else 'absent'
    info = STATUSES[status_code]
    
    try:
        students = user.get('students_list', [])
        
        for student in students:
            if len(student) >= 2:
                student_name = student[1]
                save_attendance_record(user['current_date'], user['current_lesson'], 
                                      student_name, info['text'], "-")
        
        user['selected_students'] = set()
        bot.answer_callback_query(call.id, f"✅ Все студенты отмечены как {info['text']}")
        
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        
        try:
            bot.delete_message(call.message.chat.id, call.message.message_id)
        except:
            pass
        
        show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, user['current_page'])
        
    except Exception as e:
        bot.answer_callback_query(call.id, f"❌ Ошибка: {e}")

@bot.callback_query_handler(func=lambda call: call.data == 'mark_all_sick')
def mark_all_sick(call):
    """Отметить всех студентов как болеющих"""
    user = get_user_data(call.message.chat.id)
    
    markup = telebot.types.InlineKeyboardMarkup()
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Да, все болеют", callback_data="confirm_all_sick"),
        telebot.types.InlineKeyboardButton("❌ Отмена", callback_data="back_to_list")
    )
    
    bot.edit_message_text(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text="⚠️ *Отметить ВСЕХ студентов как болеющих?*\n\n"
             "Это перезапишет текущие отметки на эту дату и пару.",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data == 'confirm_all_sick')
def confirm_all_sick(call):
    """Подтверждение отметки всех как болеющих"""
    user = get_user_data(call.message.chat.id)
    
    msg = bot.send_message(
        call.message.chat.id,
        "📝 *Введите причину болезни для всех студентов:*\n"
        "Например: ОРВИ, Грипп, Температура"
    )
    bot.register_next_step_handler(msg, save_all_sick_with_reason)

def save_all_sick_with_reason(message):
    """Сохраняет отметку болезни для всех студентов"""
    user = get_user_data(message.chat.id)
    reason = message.text
    
    students = user.get('students_list', [])
    for student in students:
        if len(student) >= 2:
            save_attendance_record(
                user['current_date'],
                user['current_lesson'],
                student[1],
                'Болел',
                reason
            )
    
    user['selected_students'] = set()
    
    bot.send_message(
        message.chat.id,
        f"✅ *Все студенты отмечены как болеющие*\n📝 *Причина:* {reason}"
    )
    
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    show_students_list_with_checkboxes(message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'mark_all_valid')
def mark_all_valid(call):
    """Отметить всех студентов с уважительной причиной"""
    user = get_user_data(call.message.chat.id)
    
    msg = bot.send_message(
        call.message.chat.id,
        "📝 *Введите уважительную причину для всех студентов:*\n"
        "Например: Соревнования, Конференция, Мероприятие"
    )
    bot.register_next_step_handler(msg, save_all_valid_with_reason)

def save_all_valid_with_reason(message):
    """Сохраняет отметку уважительной причины для всех студентов"""
    user = get_user_data(message.chat.id)
    reason = message.text
    
    students = user.get('students_list', [])
    for student in students:
        if len(student) >= 2:
            save_attendance_record(
                user['current_date'],
                user['current_lesson'],
                student[1],
                'Уважительная причина',
                reason
            )
    
    user['selected_students'] = set()
    
    bot.send_message(
        message.chat.id,
        f"✅ *Все студенты отмечены с уважительной причиной*\n📝 *Причина:* {reason}"
    )
    
    existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
    show_students_list_with_checkboxes(message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'back_to_list')
def back_to_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

@bot.callback_query_handler(func=lambda call: call.data == 'refresh_list')
def refresh_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

def refresh_students_list(chat_id, message_id=None):
    """Обновляет список студентов с сохранением выбора"""
    user = get_user_data(chat_id)
    
    try:
        all_students = students_sheet.get_all_values()
        students = all_students[1:] if len(all_students) > 1 else []
        
        old_selection = user.get('selected_students', set())
        user['students_list'] = students
        user['selected_students'] = {idx for idx in old_selection if idx < len(students)}
        
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        
        if message_id:
            try:
                bot.delete_message(chat_id, message_id)
            except:
                pass
        
        show_students_list_with_checkboxes(chat_id, students, existing_marks, user.get('current_page', 0))
        
    except Exception as e:
        bot.send_message(chat_id, f"❌ Ошибка обновления: {e}")

@bot.callback_query_handler(func=lambda call: call.data == 'save_exit')
def save_and_exit(call):
    user = get_user_data(call.message.chat.id)
    user['marking_mode'] = False
    user['selected_students'] = set()
    
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

@bot.callback_query_handler(func=lambda call: call.data == 'page_prev')
def page_prev(call):
    user = get_user_data(call.message.chat.id)
    current_page = user.get('current_page', 0)
    if current_page > 0:
        try:
            bot.delete_message(call.message.chat.id, call.message.message_id)
        except:
            pass
        students = user.get('students_list', [])
        if not students:
            all_students = students_sheet.get_all_values()
            students = all_students[1:] if len(all_students) > 1 else []
            user['students_list'] = students
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, page=current_page - 1)
    else:
        bot.answer_callback_query(call.id, "Вы на первой странице")

@bot.callback_query_handler(func=lambda call: call.data == 'page_next')
def page_next(call):
    user = get_user_data(call.message.chat.id)
    current_page = user.get('current_page', 0)
    students = user.get('students_list', [])
    total_pages = (len(students) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    if current_page < total_pages - 1:
        try:
            bot.delete_message(call.message.chat.id, call.message.message_id)
        except:
            pass
        existing_marks = get_existing_marks(user['current_date'], user['current_lesson'])
        show_students_list_with_checkboxes(call.message.chat.id, students, existing_marks, page=current_page + 1)
    else:
        bot.answer_callback_query(call.id, "Вы на последней странице")

# ==================== ДОБАВЛЕНИЕ СТУДЕНТА (ТОЛЬКО ДЛЯ ТЕСТИРОВАНИЯ) ====================
def save_new_student(message):
    """Сохраняет нового студента (вызывается из других частей кода)"""
    try:
        name = message.text.strip()
        
        if not name:
            bot.send_message(message.chat.id, "❌ Имя не может быть пустым!")
            return
        
        students = students_sheet.get_all_values()
        for student in students[1:]:
            if len(student) >= 2 and student[1] == name:
                bot.send_message(message.chat.id, f"⚠️ Студент '{name}' уже есть в списке!")
                return
        
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
    """Упрощённое меню - только отчёт за месяц"""
    current_month = datetime.date.today().strftime("%m.%Y")
    msg = bot.send_message(message.chat.id,
                          f"📅 *Введите месяц и год для отчёта*\n\n"
                          f"Формат: `ММ.ГГГГ`\n"
                          f"*Пример:* `{current_month}`\n"
                          f"Или введите `текущий` для текущего месяца",
                          parse_mode='Markdown')
    bot.register_next_step_handler(msg, generate_monthly_report)

def generate_monthly_report(message):
    """Генерирует отчёт с правильным выделением прогулов"""
    try:
        # Определяем месяц
        if message.text.lower() == 'текущий':
            month_year = datetime.date.today().strftime("%m.%Y")
        else:
            month_year = message.text
        
        month, year = map(int, month_year.split('.'))
        
        # Получаем данные
        records = attendance_sheet.get_all_records()
        if not records:
            bot.send_message(message.chat.id, "📭 Нет данных для отчёта")
            return
        
        df = pd.DataFrame(records)
        df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
        
        # Фильтруем по месяцу
        mask = (df['Дата'].dt.month == month) & (df['Дата'].dt.year == year)
        filtered = df[mask]
        
        if filtered.empty:
            bot.send_message(message.chat.id, f"📭 Нет данных за {month_year}")
            return
        
        # Получаем список студентов
        all_students_data = students_sheet.get_all_values()
        all_students = [s[1] for s in all_students_data[1:] if len(s) >= 2]
        
        # ========== 1. ЛИСТ ПОСЕЩАЕМОСТИ (СТУДЕНТЫ × ДАТЫ) ==========
        all_dates = sorted(filtered['Дата'].dt.strftime('%d.%m.%Y').unique())
        
        attendance_matrix = []
        for student in all_students:
            row = {'Студент': student}
            student_records = filtered[filtered['Студент'] == student]
            
            for date in all_dates:
                day_records = student_records[student_records['Дата'].dt.strftime('%d.%m.%Y') == date]
                if not day_records.empty:
                    status = day_records.iloc[0]['Статус']
                    # Ставим сокращённое обозначение
                    if status == 'Присутствовал':
                        row[date] = '✅'
                    elif status == 'Отсутствовал':
                        row[date] = '❌'  # ПРОГУЛ - красным
                    elif status == 'Болел':
                        row[date] = '🤒'
                    elif status == 'Уважительная причина':
                        row[date] = '📄'
                    elif status == 'Иная причина':
                        row[date] = '❓'
                    else:
                        row[date] = status
                else:
                    row[date] = ''  # Пусто, если не было пары
            attendance_matrix.append(row)
        
        df_attendance = pd.DataFrame(attendance_matrix)
        
        # ========== 2. ЛИСТ СТАТИСТИКИ (ПРАВИЛЬНЫЕ ЗАГОЛОВКИ) ==========
        stats_data = []
        
        for student in all_students:
            student_records = filtered[filtered['Студент'] == student]
            
            total_classes = len(student_records)
            present = len(student_records[student_records['Статус'] == 'Присутствовал'])
            unexcused = len(student_records[student_records['Статус'] == 'Отсутствовал'])  # ТОЛЬКО ЭТО ПРОГУЛЫ
            sick = len(student_records[student_records['Статус'] == 'Болел'])
            excused = len(student_records[student_records['Статус'] == 'Уважительная причина'])
            other = len(student_records[student_records['Статус'] == 'Иная причина'])
            
            attendance_rate = round(present / total_classes * 100, 1) if total_classes > 0 else 0
            
            stats_data.append({
                'Студент': student,
                'Всего занятий': total_classes,
                '✅ Присутствовал': present,
                '❌ ПРОГУЛ (неуваж.)': unexcused,  # ПРАВИЛЬНОЕ НАЗВАНИЕ
                '🤒 Болел': sick,
                '📄 Уважительная причина': excused,
                '❓ Иная причина': other,
                '% посещения': attendance_rate
            })
        
        df_stats = pd.DataFrame(stats_data)
        
        # ========== 3. ЛИСТ ИТОГОВ ==========
        total_unexcused = df_stats['❌ ПРОГУЛ (неуваж.)'].sum()
        students_with_absences = len(df_stats[df_stats['❌ ПРОГУЛ (неуваж.)'] > 0])
        
        summary_data = {
            'Показатель': [
                'Всего занятий в месяце',
                'Всего студентов',
                'Студентов с прогулами',
                'ВСЕГО ПРОГУЛОВ (неуваж.)',
                'Среднее число прогулов',
                'Максимум прогулов у одного студента'
            ],
            'Значение': [
                len(all_dates),
                len(all_students),
                students_with_absences,
                total_unexcused,
                round(total_unexcused / len(all_students), 1) if len(all_students) > 0 else 0,
                df_stats['❌ ПРОГУЛ (неуваж.)'].max() if not df_stats.empty else 0
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        
        # ========== 4. СОЗДАЁМ EXCEL ==========
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Записываем листы
            df_attendance.to_excel(writer, sheet_name='Посещаемость', index=False)
            df_stats.to_excel(writer, sheet_name='Статистика', index=False)
            df_summary.to_excel(writer, sheet_name='Итоги', index=False)
            
            # Причины пропусков
            reasons_df = filtered[filtered['Причина'] != '-']
            if not reasons_df.empty:
                reasons_df = reasons_df[['Дата', 'Пара', 'Студент', 'Статус', 'Причина']]
                reasons_df.to_excel(writer, sheet_name='Причины', index=False)
            
            # ========== ФОРМАТИРОВАНИЕ ==========
            workbook = writer.book
            worksheet_att = writer.sheets['Посещаемость']
            worksheet_stats = writer.sheets['Статистика']
            
            # === ФОРМАТИРОВАНИЕ ЛИСТА СТАТИСТИКИ ===
            # Заголовки (жирные, с фоном)
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True)
            
            for col in range(1, 9):
                col_letter = get_column_letter(col)
                cell = worksheet_stats[f'{col_letter}1']
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')
            
            # Ширина столбцов
            worksheet_stats.column_dimensions['A'].width = 25  # Студент
            worksheet_stats.column_dimensions['B'].width = 15  # Всего занятий
            worksheet_stats.column_dimensions['C'].width = 18  # ✅ Присутствовал
            worksheet_stats.column_dimensions['D'].width = 22  # ❌ ПРОГУЛ - САМЫЙ ВАЖНЫЙ
            worksheet_stats.column_dimensions['E'].width = 12  # 🤒 Болел
            worksheet_stats.column_dimensions['F'].width = 20  # 📄 Уважительная причина
            worksheet_stats.column_dimensions['G'].width = 15  # ❓ Иная причина
            worksheet_stats.column_dimensions['H'].width = 15  # % посещения
            
            # === КРАСНЫЙ ФОН ТОЛЬКО ДЛЯ ПРОГУЛОВ ===
            red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            red_font = Font(color='9C0006', bold=True)
            
            # Применяем красный фон к ячейкам с прогулами (>0) в столбце D
            for row in range(2, len(df_stats) + 2):
                cell = worksheet_stats.cell(row=row, column=4)  # Столбец D - ПРОГУЛЫ
                if cell.value and cell.value > 0:
                    cell.fill = red_fill
                    cell.font = red_font
            
            # === ФОРМАТИРОВАНИЕ ЛИСТА ПОСЕЩАЕМОСТИ ===
            # Ширина столбцов
            worksheet_att.column_dimensions['A'].width = 25  # Студент
            for col in range(2, len(all_dates) + 2):
                col_letter = get_column_letter(col)
                worksheet_att.column_dimensions[col_letter].width = 12  # Даты
            
            # Заголовки дат
            for col in range(2, len(all_dates) + 2):
                col_letter = get_column_letter(col)
                cell = worksheet_att[f'{col_letter}1']
                cell.alignment = Alignment(horizontal='center')
            
            # === ФОРМАТИРОВАНИЕ ЛИСТА ИТОГОВ ===
            worksheet_summary = writer.sheets['Итоги']
            worksheet_summary.column_dimensions['A'].width = 35
            worksheet_summary.column_dimensions['B'].width = 20
            
            # Автофильтр
            worksheet_stats.auto_filter.ref = worksheet_stats.dimensions
            worksheet_att.auto_filter.ref = worksheet_att.dimensions
        
        output.seek(0)
        
        # Текстовая сводка
        caption = (
            f"📊 *ОТЧЁТ ЗА {month_year}*\n\n"
            f"👥 *Группа:* {GROUP_NAME}\n"
            f"📅 *Занятий:* {len(all_dates)}\n"
            f"👤 *Студентов:* {len(all_students)}\n"
            f"❌ *ВСЕГО ПРОГУЛОВ:* {total_unexcused}\n"
            f"⚠️ *Студентов с прогулами:* {students_with_absences}\n\n"
            f"*Прогул = ❌ Отсутствовал (неуважительно)*\n"
            f"*Болезнь и уважительные причины НЕ считаются прогулами*"
        )
        
        # Отправляем файл
        bot.send_chat_action(message.chat.id, 'upload_document')
        bot.send_document(
            message.chat.id,
            output,
            caption=caption,
            parse_mode='Markdown',
            visible_file_name=f'прогулы_{GROUP_NAME}_{month_year}.xlsx'
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
    print(f"🤖 Бот для учёта посещаемости ЗАПУЩЕН!")
    print(f"📍 Группа: {GROUP_NAME}")
    print(f"✅ Множественный выбор студентов - АКТИВЕН")
    print(f"📊 Отчёт: только прогулы выделены красным")
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
