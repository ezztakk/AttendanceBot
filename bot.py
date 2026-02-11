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
import time
from threading import Lock

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

# ==================== БЕЗОПАСНОЕ РЕДАКТИРОВАНИЕ СООБЩЕНИЙ ====================
def safe_edit_message(chat_id, message_id, text, reply_markup=None, parse_mode='Markdown'):
    """Безопасное обновление сообщения - игнорирует ошибку 'message is not modified'"""
    try:
        bot.edit_message_text(
            chat_id=chat_id,
            message_id=message_id,
            text=text,
            parse_mode=parse_mode,
            reply_markup=reply_markup
        )
    except Exception as e:
        if "message is not modified" in str(e).lower():
            # Игнорируем эту ошибку - сообщение уже актуально
            pass
        else:
            print(f"⚠️ Ошибка при редактировании: {e}")
# ====================================================

# ==================== БАЗОВОЕ КЭШИРОВАНИЕ ====================
class SheetsCache:
    """Базовый кэш для данных Google Sheets"""
    def __init__(self):
        self.students_cache = []
        self.students_timestamp = 0
        self.attendance_cache = {}
        self.attendance_timestamp = {}
        self.cache_ttl = 30
        self.lock = Lock()
        self.max_retries = 5
        self.base_delay = 1
    
    def _safe_call(self, func, *args, **kwargs):
        for attempt in range(self.max_retries):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_str = str(e)
                if '429' in error_str or 'RESOURCE_EXHAUSTED' in error_str:
                    if attempt < self.max_retries - 1:
                        delay = self.base_delay * (2 ** attempt)
                        print(f"⚠️ Превышена квота API. Ожидание {delay} сек... (попытка {attempt + 1}/{self.max_retries})")
                        time.sleep(delay)
                    else:
                        print("❌ Исчерпаны все попытки вызова API")
                        raise e
                else:
                    raise e
    
    def get_students(self):
        with self.lock:
            current_time = time.time()
            if not self.students_cache or current_time - self.students_timestamp > self.cache_ttl:
                try:
                    self.students_cache = self._safe_call(students_sheet.get_all_values)
                    self.students_timestamp = current_time
                    print("📥 Загружен список студентов (кэш обновлён)")
                except Exception as e:
                    if self.students_cache:
                        print("⚠️ Используем устаревший кэш студентов")
                        return self.students_cache
                    raise e
            return self.students_cache
    
    def get_attendance(self, date, lesson):
        key = f"{date}_{lesson}"
        with self.lock:
            current_time = time.time()
            if key not in self.attendance_cache or current_time - self.attendance_timestamp.get(key, 0) > self.cache_ttl:
                try:
                    records = self._safe_call(attendance_sheet.get_all_records)
                    filtered = {}
                    for record in records:
                        if (str(record.get('Дата', '')) == date and
                            str(record.get('Пара', '')) == str(lesson)):
                            student_name = record.get('Студент', '')
                            if student_name:
                                filtered[student_name] = {
                                    'status': record.get('Статус', ''),
                                    'reason': record.get('Причина', '')
                                }
                    self.attendance_cache[key] = filtered
                    self.attendance_timestamp[key] = current_time
                    print(f"📥 Загружены отметки для {date} пара {lesson} (кэш обновлён)")
                except Exception as e:
                    if key in self.attendance_cache:
                        print(f"⚠️ Используем устаревший кэш для {date} пара {lesson}")
                        return self.attendance_cache[key]
                    raise e
            return self.attendance_cache[key]
    
    def clear_attendance_cache(self, date=None, lesson=None):
        with self.lock:
            if date and lesson:
                key = f"{date}_{lesson}"
                self.attendance_cache.pop(key, None)
                self.attendance_timestamp.pop(key, None)
                print(f"🗑️ Очищен кэш для {date} пара {lesson}")
            elif date:
                keys_to_remove = [key for key in self.attendance_cache.keys() if key.startswith(f"{date}_")]
                for key in keys_to_remove:
                    self.attendance_cache.pop(key, None)
                    self.attendance_timestamp.pop(key, None)
                print(f"🗑️ Очищен кэш для всех пар {date}")
            else:
                self.attendance_cache.clear()
                self.attendance_timestamp.clear()
                print("🗑️ Очищен весь кэш отметок")
    
    def clear_students_cache(self):
        with self.lock:
            self.students_cache = []
            self.students_timestamp = 0
            print("🗑️ Очищен кэш студентов")

# ==================== УЛУЧШЕННОЕ КЭШИРОВАНИЕ ====================
class ImprovedSheetsCache(SheetsCache):
    """Улучшенный кэш с принудительным ожиданием между запросами"""
    
    def __init__(self):
        super().__init__()
        self.last_request_time = 0
        self.min_request_interval = 1.1  # Минимум 1.1 секунда между запросами (<60 в минуту)
    
    def _wait_for_rate_limit(self):
        """Принудительное ожидание для соблюдения квоты"""
        now = time.time()
        time_since_last = now - self.last_request_time
        if time_since_last < self.min_request_interval:
            wait_time = self.min_request_interval - time_since_last
            time.sleep(wait_time)
        self.last_request_time = time.time()
    
    def _safe_call(self, func, *args, **kwargs):
        """Безопасный вызов API с ожиданием и повторными попытками"""
        self._wait_for_rate_limit()
        
        for attempt in range(self.max_retries):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                error_str = str(e)
                if '429' in error_str or 'RESOURCE_EXHAUSTED' in error_str:
                    if attempt < self.max_retries - 1:
                        delay = self.base_delay * (4 ** attempt)
                        print(f"⚠️ Квота API превышена. Ожидание {delay} сек... (попытка {attempt + 1}/{self.max_retries})")
                        time.sleep(delay)
                        self._wait_for_rate_limit()
                    else:
                        print("❌ Исчерпаны все попытки вызова API")
                        raise
                else:
                    raise
# ====================================================

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
    
    # Инициализируем улучшенный кэш
    cache = ImprovedSheetsCache()
    print("✅ Улучшенная система кэширования запущена")
    
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
            'selected_lessons': set(),  # Множественный выбор пар
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
    btn2 = telebot.types.KeyboardButton('🔢 Выбрать пары')
    btn3 = telebot.types.KeyboardButton('📝 Отметить студентов')
    btn4 = telebot.types.KeyboardButton('📊 Получить отчёт')
    btn5 = telebot.types.KeyboardButton('ℹ️ Текущие настройки')
    markup.add(btn1, btn2, btn3, btn4, btn5)
    
    # Формируем текст о выбранных парах
    if user.get('selected_lessons'):
        selected = sorted(user['selected_lessons'])
        lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected))}"
    else:
        lessons_text = "🔢 *Пары:* не выбраны"
    
    bot.send_message(message.chat.id,
                    f"👋 *Система учёта посещаемости*\n"
                    f"👥 *Группа:* {GROUP_NAME}\n\n"
                    f"📅 *Дата:* {user['current_date']}\n"
                    f"{lessons_text}\n\n"
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
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📅 *Дата установлена:* {user['current_date']}\n\n"
             f"Теперь выберите пары и отмечайте студентов.",
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

# ==================== ВЫБОР ПАР (МНОЖЕСТВЕННЫЙ) ====================
@bot.message_handler(func=lambda message: message.text == '🔢 Выбрать пары')
def choose_lessons(message):
    user = get_user_data(message.chat.id)
    
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    # Кнопки для всех пар
    for lesson_num in range(1, 7):
        time_slot = LESSON_TIMES.get(lesson_num, "")
        
        # Отмечаем выбранные пары
        if lesson_num in user.get('selected_lessons', set()):
            btn_text = f"✅ {lesson_num} пара ({time_slot})"
        else:
            btn_text = f"{lesson_num} пара ({time_slot})"
        
        markup.add(
            telebot.types.InlineKeyboardButton(
                btn_text,
                callback_data=f"toggle_lesson_{lesson_num}"
            )
        )
    
    # Кнопки управления
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Выбрать все", callback_data="lessons_all"),
        telebot.types.InlineKeyboardButton("❌ Очистить все", callback_data="lessons_clear")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("📌 Готово", callback_data="lessons_done")
    )
    
    selected = user.get('selected_lessons', set())
    selected_text = f"✅ *Выбрано пар:* {len(selected)}" if selected else "❌ *Ничего не выбрано*"
    
    bot.send_message(message.chat.id,
                    f"🔢 *ВЫБОР ПАР*\n\n"
                    f"{selected_text}\n\n"
                    f"*Расписание:*\n"
                    f"1. {LESSON_TIMES[1]}\n"
                    f"2. {LESSON_TIMES[2]}\n"
                    f"3. {LESSON_TIMES[3]}\n"
                    f"4. {LESSON_TIMES[4]}\n"
                    f"5. {LESSON_TIMES[5]}\n"
                    f"6. {LESSON_TIMES[6]}\n\n"
                    f"*Нажимайте на пары, чтобы выбрать/снять выбор*",
                    parse_mode='Markdown',
                    reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('toggle_lesson_'))
def toggle_lesson(call):
    """Выбор/снятие выбора пары"""
    user = get_user_data(call.message.chat.id)
    lesson_num = int(call.data.split('_')[2])
    
    if 'selected_lessons' not in user:
        user['selected_lessons'] = set()
    
    if lesson_num in user['selected_lessons']:
        user['selected_lessons'].remove(lesson_num)
        bot.answer_callback_query(call.id, f"❌ Пара {lesson_num} снята")
    else:
        user['selected_lessons'].add(lesson_num)
        bot.answer_callback_query(call.id, f"✅ Пара {lesson_num} выбрана")
    
    # Обновляем меню
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    for num in range(1, 7):
        time_slot = LESSON_TIMES.get(num, "")
        if num in user['selected_lessons']:
            btn_text = f"✅ {num} пара ({time_slot})"
        else:
            btn_text = f"{num} пара ({time_slot})"
        
        markup.add(
            telebot.types.InlineKeyboardButton(
                btn_text,
                callback_data=f"toggle_lesson_{num}"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Выбрать все", callback_data="lessons_all"),
        telebot.types.InlineKeyboardButton("❌ Очистить все", callback_data="lessons_clear")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("📌 Готово", callback_data="lessons_done")
    )
    
    selected = user['selected_lessons']
    selected_text = f"✅ *Выбрано пар:* {len(selected)}" if selected else "❌ *Ничего не выбрано*"
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"🔢 *ВЫБОР ПАР*\n\n"
             f"{selected_text}\n\n"
             f"*Расписание:*\n"
             f"1. {LESSON_TIMES[1]}\n"
             f"2. {LESSON_TIMES[2]}\n"
             f"3. {LESSON_TIMES[3]}\n"
             f"4. {LESSON_TIMES[4]}\n"
             f"5. {LESSON_TIMES[5]}\n"
             f"6. {LESSON_TIMES[6]}\n\n"
             f"*Нажимайте на пары, чтобы выбрать/снять выбор*",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data == 'lessons_all')
def lessons_all(call):
    """Выбрать все пары - БЕЗ РЕКУРСИИ"""
    user = get_user_data(call.message.chat.id)
    user['selected_lessons'] = {1, 2, 3, 4, 5, 6}
    bot.answer_callback_query(call.id, "✅ Выбраны все пары")
    
    # Создаём новую клавиатуру без вызова toggle_lesson
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    for num in range(1, 7):
        time_slot = LESSON_TIMES.get(num, "")
        btn_text = f"✅ {num} пара ({time_slot})"
        markup.add(
            telebot.types.InlineKeyboardButton(
                btn_text,
                callback_data=f"toggle_lesson_{num}"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Выбрать все", callback_data="lessons_all"),
        telebot.types.InlineKeyboardButton("❌ Очистить все", callback_data="lessons_clear")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("📌 Готово", callback_data="lessons_done")
    )
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"🔢 *ВЫБОР ПАР*\n\n"
             f"✅ *Выбрано пар:* 6\n\n"
             f"*Расписание:*\n"
             f"1. {LESSON_TIMES[1]}\n"
             f"2. {LESSON_TIMES[2]}\n"
             f"3. {LESSON_TIMES[3]}\n"
             f"4. {LESSON_TIMES[4]}\n"
             f"5. {LESSON_TIMES[5]}\n"
             f"6. {LESSON_TIMES[6]}\n\n"
             f"*Нажимайте на пары, чтобы выбрать/снять выбор*",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data == 'lessons_clear')
def lessons_clear(call):
    """Очистить выбор всех пар - БЕЗ РЕКУРСИИ"""
    user = get_user_data(call.message.chat.id)
    user['selected_lessons'] = set()
    bot.answer_callback_query(call.id, "❌ Выбор очищен")
    
    # Создаём новую клавиатуру без вызова toggle_lesson
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    for num in range(1, 7):
        time_slot = LESSON_TIMES.get(num, "")
        btn_text = f"{num} пара ({time_slot})"
        markup.add(
            telebot.types.InlineKeyboardButton(
                btn_text,
                callback_data=f"toggle_lesson_{num}"
            )
        )
    
    markup.add(
        telebot.types.InlineKeyboardButton("✅ Выбрать все", callback_data="lessons_all"),
        telebot.types.InlineKeyboardButton("❌ Очистить все", callback_data="lessons_clear")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("📌 Готово", callback_data="lessons_done")
    )
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"🔢 *ВЫБОР ПАР*\n\n"
             f"❌ *Ничего не выбрано*\n\n"
             f"*Расписание:*\n"
             f"1. {LESSON_TIMES[1]}\n"
             f"2. {LESSON_TIMES[2]}\n"
             f"3. {LESSON_TIMES[3]}\n"
             f"4. {LESSON_TIMES[4]}\n"
             f"5. {LESSON_TIMES[5]}\n"
             f"6. {LESSON_TIMES[6]}\n\n"
             f"*Нажимайте на пары, чтобы выбрать/снять выбор*",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data == 'lessons_done')
def lessons_done(call):
    """Завершить выбор пар"""
    user = get_user_data(call.message.chat.id)
    
    if not user.get('selected_lessons'):
        bot.answer_callback_query(call.id, "❌ Выберите хотя бы одну пару!")
        return
    
    selected = sorted(user['selected_lessons'])
    selected_text = ", ".join(map(str, selected))
    
    bot.answer_callback_query(call.id, f"✅ Выбраны пары: {selected_text}")
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"✅ *Настройки установлены*\n\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"🔢 *Выбранные пары:* {selected_text}\n\n"
             f"Теперь можно *отметить студентов* 👇",
        parse_mode='Markdown'
    )

# ==================== ПОЛУЧЕНИЕ СУЩЕСТВУЮЩИХ ОТМЕТОК (С КЭШЕМ) ====================
def get_existing_marks(date, lesson):
    """Получаем существующие отметки для даты и пары с кэшированием"""
    try:
        return cache.get_attendance(date, lesson)
    except Exception as e:
        print(f"❌ Ошибка получения отметок: {e}")
        return {}

# ==================== СОХРАНЕНИЕ ЗАПИСИ (БАТЧОВОЕ) ====================
def save_attendance_record(date, lessons, student, status, reason):
    """Сохраняет запись о посещении для одной или нескольких пар (батч-операция)"""
    try:
        if isinstance(lessons, (list, set)):
            lesson_list = list(lessons)
        else:
            lesson_list = [lessons]
        
        # Получаем все записи ОДИН РАЗ с задержкой
        time.sleep(1.1)
        records = attendance_sheet.get_all_values()
        
        # Собираем строки для удаления
        rows_to_delete = []
        rows_to_add = []
        
        for lesson in lesson_list:
            # Ищем существующие записи
            for i, row in enumerate(records):
                if (i > 0 and len(row) >= 4 and
                    str(row[0]) == date and
                    str(row[1]) == str(lesson) and
                    str(row[3]) == student):
                    rows_to_delete.append(i + 1)
            
            # Добавляем новую запись
            time_now = datetime.datetime.now().strftime("%H:%M")
            rows_to_add.append([
                date,
                lesson,
                GROUP_NAME,
                student,
                status,
                reason,
                time_now
            ])
        
        # Батчевое удаление
        if rows_to_delete:
            for row_num in sorted(rows_to_delete, reverse=True):
                attendance_sheet.delete_rows(row_num)
            print(f"🗑️ Удалено {len(rows_to_delete)} записей")
        
        # Батчевое добавление
        if rows_to_add:
            for row in rows_to_add:
                attendance_sheet.append_row(row)
            print(f"📝 Добавлено {len(rows_to_add)} записей")
        
        # Очищаем кэш для затронутых дат и пар
        for lesson in lesson_list:
            cache.clear_attendance_cache(date, lesson)
        
        return len(rows_to_add)
    except Exception as e:
        print(f"❌ Ошибка сохранения: {e}")
        return 0

# ==================== СОЗДАНИЕ КЛАВИАТУРЫ СТУДЕНТОВ ====================
def create_students_markup(students, existing_marks, page, selected_students):
    """Создаёт клавиатуру со списком студентов (без отправки сообщения)"""
    markup = telebot.types.InlineKeyboardMarkup(row_width=2)
    
    selected_count = len(selected_students)
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
    
    total_students = len(students)
    total_pages = (total_students + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    
    if page < 0:
        page = 0
    elif page >= total_pages:
        page = total_pages - 1
    
    start = page * ITEMS_PER_PAGE
    end = min(start + ITEMS_PER_PAGE, total_students)
    
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
            
            checkbox = "☑️" if idx_in_list in selected_students else "◻️"
            
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
    
    markup.add(
        telebot.types.InlineKeyboardButton("❌ Снять все выборы", callback_data="clear_selection"),
        telebot.types.InlineKeyboardButton("🔄 Обновить", callback_data="refresh_list")
    )
    
    markup.add(
        telebot.types.InlineKeyboardButton("💾 СОХРАНИТЬ И ВЫЙТИ", callback_data="save_exit")
    )
    
    return markup

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
    
    markup = create_students_markup(students, existing_marks, page, user['selected_students'])
    
    selected_count = len(user['selected_students'])
    selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
    
    # Информация о выбранных парах
    lessons_text = ""
    if user.get('selected_lessons'):
        selected_lessons = sorted(user['selected_lessons'])
        lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
    
    page_info = f"📄 Страница {page+1} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
    
    bot.send_message(
        chat_id,
        f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
        f"👥 *Группа:* {GROUP_NAME}\n"
        f"📅 *Дата:* {user['current_date']}\n"
        f"{lessons_text}"
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

# ==================== БЕЗОПАСНОЕ ПОЛУЧЕНИЕ СТУДЕНТА ====================
def get_student_by_index(user, idx):
    """Безопасное получение студента по индексу"""
    if 'students_list' not in user:
        return None
    if idx >= len(user['students_list']):
        return None
    if len(user['students_list'][idx]) < 2:
        return None
    return user['students_list'][idx][1]

# ==================== ОБРАБОТЧИКИ ДЛЯ ОТМЕТКИ ====================
@bot.message_handler(func=lambda message: message.text == '📝 Отметить студентов')
def mark_students(message):
    user = get_user_data(message.chat.id)
    
    # Проверяем, выбраны ли пары
    if not user.get('selected_lessons'):
        bot.send_message(message.chat.id, 
                        "❌ *Сначала выберите пары!*\n"
                        "Нажмите 🔢 Выбрать пары",
                        parse_mode='Markdown')
        return
    
    try:
        # Используем кэшированный список студентов
        all_students = cache.get_students()
        students = all_students[1:] if len(all_students) > 1 else []
        
        if len(students) <= 0:
            bot.send_message(message.chat.id, "❌ Сначала добавьте студентов!")
            return
        
        user['students_list'] = students
        user['selected_students'] = set()
        user['current_page'] = 0
        
        # Получаем отметки для ВСЕХ выбранных пар (с кэшированием)
        existing_marks = {}
        for lesson in user['selected_lessons']:
            marks = get_existing_marks(user['current_date'], lesson)
            for student, data in marks.items():
                if student not in existing_marks:
                    existing_marks[student] = data
        
        user['marking_mode'] = True
        
        selected_lessons = sorted(user['selected_lessons'])
        lessons_text = ", ".join(map(str, selected_lessons))
        
        bot.send_message(message.chat.id,
                        f"📌 *Отметка для нескольких пар*\n"
                        f"🔢 *Пары:* {lessons_text}\n"
                        f"📅 *Дата:* {user['current_date']}\n\n"
                        f"*Отметки будут применены ко ВСЕМ выбранным парам!*",
                        parse_mode='Markdown')
        
        show_students_list_with_checkboxes(message.chat.id, students, existing_marks, 0)
        
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")

@bot.callback_query_handler(func=lambda call: call.data.startswith('toggle_'))
def toggle_student(call):
    """Выбор/снятие выбора студента (обновление сообщения без удаления)"""
    user = get_user_data(call.message.chat.id)
    idx = int(call.data.split('_')[1])
    
    # Защита от невалидного индекса
    if idx >= len(user.get('students_list', [])):
        bot.answer_callback_query(call.id, "❌ Данные устарели, обновите список")
        refresh_students_list(call.message.chat.id, call.message.message_id)
        return
    
    if idx in user['selected_students']:
        user['selected_students'].remove(idx)
        bot.answer_callback_query(call.id, "❌ Выбор снят")
    else:
        user['selected_students'].add(idx)
        bot.answer_callback_query(call.id, "✅ Студент выбран")
    
    students = user.get('students_list', [])
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
    
    # Обновляем существующее сообщение
    markup = create_students_markup(students, existing_marks, user['current_page'], user['selected_students'])
    selected_count = len(user['selected_students'])
    selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
    
    # Информация о выбранных парах
    lessons_text = ""
    if user.get('selected_lessons'):
        selected_lessons = sorted(user['selected_lessons'])
        lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
    
    page = user['current_page']
    total_students = len(students)
    total_pages = (total_students + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    page_info = f"📄 Страница {page+1} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
             f"👥 *Группа:* {GROUP_NAME}\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"{lessons_text}"
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

@bot.callback_query_handler(func=lambda call: call.data == 'clear_selection')
def clear_selection(call):
    """Снять все выборы (обновление сообщения без удаления)"""
    user = get_user_data(call.message.chat.id)
    user['selected_students'] = set()
    bot.answer_callback_query(call.id, "❌ Все выборы сняты")
    
    students = user.get('students_list', [])
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
    
    # Обновляем существующее сообщение
    markup = create_students_markup(students, existing_marks, user['current_page'], user['selected_students'])
    
    # Информация о выбранных парах
    lessons_text = ""
    if user.get('selected_lessons'):
        selected_lessons = sorted(user['selected_lessons'])
        lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
    
    page = user['current_page']
    total_students = len(students)
    total_pages = (total_students + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    page_info = f"📄 Страница {page+1} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
             f"👥 *Группа:* {GROUP_NAME}\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"{lessons_text}"
             f"✅ *Выбрано:* 0 студентов\n"
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
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"📝 *Применить статус к выбранным студентам*\n\n"
             f"✅ *Выбрано:* {len(user['selected_students'])} студентов\n\n"
             f"*Отметка будет применена ко всем выбранным парам:*\n"
             f"{', '.join(map(str, sorted(user['selected_lessons'])))}",
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
            f"Причина будет применена ко всем выбранным студентам и всем выбранным парам."
        )
        bot.register_next_step_handler(msg, save_reason_for_selected)
        return
    else:
        for idx in user['selected_students']:
            student_name = get_student_by_index(user, idx)
            if student_name:
                save_attendance_record(
                    user['current_date'], 
                    user['selected_lessons'],
                    student_name, 
                    info['text'], 
                    "-"
                )
    
    user['selected_students'] = set()
    bot.answer_callback_query(call.id, f"✅ Отмечено студентов")
    
    students = user.get('students_list', [])
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
    
    back_to_list_with_data(call.message.chat.id, call.message.message_id, students, existing_marks)

def back_to_list_with_data(chat_id, message_id, students, existing_marks):
    """Возврат к списку студентов с обновлением сообщения"""
    user = get_user_data(chat_id)
    
    markup = create_students_markup(students, existing_marks, user['current_page'], user['selected_students'])
    selected_count = len(user['selected_students'])
    selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
    
    lessons_text = ""
    if user.get('selected_lessons'):
        selected_lessons = sorted(user['selected_lessons'])
        lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
    
    page = user['current_page']
    total_students = len(students)
    total_pages = (total_students + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    page_info = f"📄 Страница {page+1} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
    
    safe_edit_message(
        chat_id=chat_id,
        message_id=message_id,
        text=f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
             f"👥 *Группа:* {GROUP_NAME}\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"{lessons_text}"
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

def save_reason_for_selected(message):
    """Сохраняет причину для всех выбранных студентов"""
    user = get_user_data(message.chat.id)
    reason = message.text
    
    if 'pending_status' not in user:
        bot.send_message(message.chat.id, "❌ Ошибка: данные не найдены")
        return
    
    pending = user['pending_status']
    
    for idx in pending['students']:
        student_name = get_student_by_index(user, idx)
        if student_name:
            save_attendance_record(
                user['current_date'],
                user['selected_lessons'],
                student_name,
                pending['status_text'],
                reason
            )
    
    user['selected_students'] = set()
    del user['pending_status']
    
    bot.send_message(
        message.chat.id,
        f"✅ *Отмечено {len(pending['students'])} студентов*\n"
        f"📝 *Причина:* {reason}\n"
        f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}"
    )
    
    students = user.get('students_list', [])
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
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
                save_attendance_record(
                    user['current_date'], 
                    user['selected_lessons'],
                    student_name, 
                    info['text'], 
                    "-"
                )
        
        user['selected_students'] = set()
        bot.answer_callback_query(call.id, f"✅ Все студенты отмечены как {info['text']}")
        
        existing_marks = {}
        for lesson in user['selected_lessons']:
            marks = get_existing_marks(user['current_date'], lesson)
            for student, data in marks.items():
                if student not in existing_marks:
                    existing_marks[student] = data
        
        back_to_list_with_data(call.message.chat.id, call.message.message_id, students, existing_marks)
        
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
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"⚠️ *Отметить ВСЕХ студентов как болеющих?*\n\n"
             f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}\n"
             f"📅 *Дата:* {user['current_date']}\n\n"
             f"Это перезапишет текущие отметки на выбранные пары.",
        parse_mode='Markdown',
        reply_markup=markup
    )

@bot.callback_query_handler(func=lambda call: call.data == 'confirm_all_sick')
def confirm_all_sick(call):
    """Подтверждение отметки всех как болеющих"""
    user = get_user_data(call.message.chat.id)
    
    msg = bot.send_message(
        call.message.chat.id,
        f"📝 *Введите причину болезни для всех студентов:*\n"
        f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}\n"
        f"📅 *Дата:* {user['current_date']}\n\n"
        f"Например: ОРВИ, Грипп, Температура"
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
                user['selected_lessons'],
                student[1],
                'Болел',
                reason
            )
    
    user['selected_students'] = set()
    
    bot.send_message(
        message.chat.id,
        f"✅ *Все студенты отмечены как болеющие*\n"
        f"📝 *Причина:* {reason}\n"
        f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}"
    )
    
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
    show_students_list_with_checkboxes(message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'mark_all_valid')
def mark_all_valid(call):
    """Отметить всех студентов с уважительной причиной"""
    user = get_user_data(call.message.chat.id)
    
    msg = bot.send_message(
        call.message.chat.id,
        f"📝 *Введите уважительную причину для всех студентов:*\n"
        f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}\n"
        f"📅 *Дата:* {user['current_date']}\n\n"
        f"Например: Соревнования, Конференция, Мероприятие"
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
                user['selected_lessons'],
                student[1],
                'Уважительная причина',
                reason
            )
    
    user['selected_students'] = set()
    
    bot.send_message(
        message.chat.id,
        f"✅ *Все студенты отмечены с уважительной причиной*\n"
        f"📝 *Причина:* {reason}\n"
        f"🔢 *Пары:* {', '.join(map(str, sorted(user['selected_lessons'])))}"
    )
    
    existing_marks = {}
    for lesson in user['selected_lessons']:
        marks = get_existing_marks(user['current_date'], lesson)
        for student, data in marks.items():
            if student not in existing_marks:
                existing_marks[student] = data
    show_students_list_with_checkboxes(message.chat.id, students, existing_marks, user['current_page'])

@bot.callback_query_handler(func=lambda call: call.data == 'back_to_list')
def back_to_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

@bot.callback_query_handler(func=lambda call: call.data == 'refresh_list')
def refresh_list(call):
    refresh_students_list(call.message.chat.id, call.message.message_id)

def refresh_students_list(chat_id, message_id=None):
    """Обновляет список студентов с сохранением выбора (с кэшированием)"""
    user = get_user_data(chat_id)
    
    try:
        all_students = cache.get_students()
        students = all_students[1:] if len(all_students) > 1 else []
        
        old_selection = user.get('selected_students', set())
        user['students_list'] = students
        user['selected_students'] = {idx for idx in old_selection if idx < len(students)}
        
        existing_marks = {}
        for lesson in user['selected_lessons']:
            marks = get_existing_marks(user['current_date'], lesson)
            for student, data in marks.items():
                if student not in existing_marks:
                    existing_marks[student] = data
        
        if message_id:
            back_to_list_with_data(chat_id, message_id, students, existing_marks)
        else:
            show_students_list_with_checkboxes(chat_id, students, existing_marks, user.get('current_page', 0))
        
    except Exception as e:
        bot.send_message(chat_id, f"❌ Ошибка обновления: {e}")

@bot.callback_query_handler(func=lambda call: call.data == 'save_exit')
def save_and_exit(call):
    user = get_user_data(call.message.chat.id)
    user['marking_mode'] = False
    user['selected_students'] = set()
    
    bot.answer_callback_query(call.id, "✅ Данные сохранены")
    
    selected_lessons = sorted(user['selected_lessons'])
    lessons_text = ", ".join(map(str, selected_lessons)) if selected_lessons else "не выбраны"
    
    safe_edit_message(
        chat_id=call.message.chat.id,
        message_id=call.message.message_id,
        text=f"✅ *Данные сохранены!*\n\n"
             f"📅 *Дата:* {user['current_date']}\n"
             f"🔢 *Пары:* {lessons_text}\n"
             f"👥 *Группа:* {GROUP_NAME}\n\n"
             f"Для нового действия нажмите /start",
        parse_mode='Markdown'
    )

@bot.callback_query_handler(func=lambda call: call.data == 'page_prev')
def page_prev(call):
    user = get_user_data(call.message.chat.id)
    current_page = user.get('current_page', 0)
    if current_page > 0:
        students = user.get('students_list', [])
        if not students:
            all_students = cache.get_students()
            students = all_students[1:] if len(all_students) > 1 else []
            user['students_list'] = students
        
        existing_marks = {}
        for lesson in user['selected_lessons']:
            marks = get_existing_marks(user['current_date'], lesson)
            for student, data in marks.items():
                if student not in existing_marks:
                    existing_marks[student] = data
        
        user['current_page'] = current_page - 1
        markup = create_students_markup(students, existing_marks, current_page - 1, user['selected_students'])
        
        selected_count = len(user['selected_students'])
        selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
        
        lessons_text = ""
        if user.get('selected_lessons'):
            selected_lessons = sorted(user['selected_lessons'])
            lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
        
        total_pages = (len(students) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
        page_info = f"📄 Страница {current_page} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
        
        safe_edit_message(
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            text=f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
                 f"👥 *Группа:* {GROUP_NAME}\n"
                 f"📅 *Дата:* {user['current_date']}\n"
                 f"{lessons_text}"
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
    else:
        bot.answer_callback_query(call.id, "Вы на первой странице")

@bot.callback_query_handler(func=lambda call: call.data == 'page_next')
def page_next(call):
    user = get_user_data(call.message.chat.id)
    current_page = user.get('current_page', 0)
    students = user.get('students_list', [])
    total_pages = (len(students) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    
    if current_page < total_pages - 1:
        existing_marks = {}
        for lesson in user['selected_lessons']:
            marks = get_existing_marks(user['current_date'], lesson)
            for student, data in marks.items():
                if student not in existing_marks:
                    existing_marks[student] = data
        
        user['current_page'] = current_page + 1
        markup = create_students_markup(students, existing_marks, current_page + 1, user['selected_students'])
        
        selected_count = len(user['selected_students'])
        selected_text = f"✅ *Выбрано:* {selected_count} студентов\n" if selected_count > 0 else ""
        
        lessons_text = ""
        if user.get('selected_lessons'):
            selected_lessons = sorted(user['selected_lessons'])
            lessons_text = f"🔢 *Пары:* {', '.join(map(str, selected_lessons))}\n"
        
        page_info = f"📄 Страница {current_page + 2} из {total_pages}" if total_pages > 0 else "📄 Нет студентов"
        
        safe_edit_message(
            chat_id=call.message.chat.id,
            message_id=call.message.message_id,
            text=f"📝 *ОТМЕТКА ПОСЕЩАЕМОСТИ*\n\n"
                 f"👥 *Группа:* {GROUP_NAME}\n"
                 f"📅 *Дата:* {user['current_date']}\n"
                 f"{lessons_text}"
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
    else:
        bot.answer_callback_query(call.id, "Вы на последней странице")

# ==================== ДОБАВЛЕНИЕ СТУДЕНТА ====================
def save_new_student(message):
    """Сохраняет нового студента"""
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
        cache.clear_students_cache()
        
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
        if message.text.lower() == 'текущий':
            month_year = datetime.date.today().strftime("%m.%Y")
        else:
            month_year = message.text
        
        month, year = map(int, month_year.split('.'))
        
        time.sleep(1.1)
        records = attendance_sheet.get_all_records()
        if not records:
            bot.send_message(message.chat.id, "📭 Нет данных для отчёта")
            return
        
        df = pd.DataFrame(records)
        df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
        
        mask = (df['Дата'].dt.month == month) & (df['Дата'].dt.year == year)
        filtered = df[mask]
        
        if filtered.empty:
            bot.send_message(message.chat.id, f"📭 Нет данных за {month_year}")
            return
        
        all_students_data = cache.get_students()
        all_students = [s[1] for s in all_students_data[1:] if len(s) >= 2]
        
        all_dates = sorted(filtered['Дата'].dt.strftime('%d.%m.%Y').unique())
        
        attendance_matrix = []
        for student in all_students:
            row = {'Студент': student}
            student_records = filtered[filtered['Студент'] == student]
            
            for date in all_dates:
                day_records = student_records[student_records['Дата'].dt.strftime('%d.%m.%Y') == date]
                if not day_records.empty:
                    status = day_records.iloc[0]['Статус']
                    if status == 'Присутствовал':
                        row[date] = '✅'
                    elif status == 'Отсутствовал':
                        row[date] = '❌'
                    elif status == 'Болел':
                        row[date] = '🤒'
                    elif status == 'Уважительная причина':
                        row[date] = '📄'
                    elif status == 'Иная причина':
                        row[date] = '❓'
                    else:
                        row[date] = status
                else:
                    row[date] = ''
            attendance_matrix.append(row)
        
        df_attendance = pd.DataFrame(attendance_matrix)
        
        stats_data = []
        for student in all_students:
            student_records = filtered[filtered['Студент'] == student]
            
            total_classes = len(student_records)
            present = len(student_records[student_records['Статус'] == 'Присутствовал'])
            unexcused = len(student_records[student_records['Статус'] == 'Отсутствовал'])
            sick = len(student_records[student_records['Статус'] == 'Болел'])
            excused = len(student_records[student_records['Статус'] == 'Уважительная причина'])
            other = len(student_records[student_records['Статус'] == 'Иная причина'])
            
            attendance_rate = round(present / total_classes * 100, 1) if total_classes > 0 else 0
            
            stats_data.append({
                'Студент': student,
                'Всего занятий': total_classes,
                '✅ Присутствовал': present,
                '❌ ПРОГУЛ (неуваж.)': unexcused,
                '🤒 Болел': sick,
                '📄 Уважительная причина': excused,
                '❓ Иная причина': other,
                '% посещения': attendance_rate
            })
        
        df_stats = pd.DataFrame(stats_data)
        
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
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_attendance.to_excel(writer, sheet_name='Посещаемость', index=False)
            df_stats.to_excel(writer, sheet_name='Статистика', index=False)
            df_summary.to_excel(writer, sheet_name='Итоги', index=False)
            
            reasons_df = filtered[filtered['Причина'] != '-']
            if not reasons_df.empty:
                reasons_df = reasons_df[['Дата', 'Пара', 'Студент', 'Статус', 'Причина']]
                reasons_df.to_excel(writer, sheet_name='Причины', index=False)
            
            workbook = writer.book
            worksheet_att = writer.sheets['Посещаемость']
            worksheet_stats = writer.sheets['Статистика']
            
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True)
            
            for col in range(1, 9):
                col_letter = get_column_letter(col)
                cell = worksheet_stats[f'{col_letter}1']
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')
            
            worksheet_stats.column_dimensions['A'].width = 25
            worksheet_stats.column_dimensions['B'].width = 15
            worksheet_stats.column_dimensions['C'].width = 18
            worksheet_stats.column_dimensions['D'].width = 22
            worksheet_stats.column_dimensions['E'].width = 12
            worksheet_stats.column_dimensions['F'].width = 20
            worksheet_stats.column_dimensions['G'].width = 15
            worksheet_stats.column_dimensions['H'].width = 15
            
            red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            red_font = Font(color='9C0006', bold=True)
            
            for row in range(2, len(df_stats) + 2):
                cell = worksheet_stats.cell(row=row, column=4)
                if cell.value and cell.value > 0:
                    cell.fill = red_fill
                    cell.font = red_font
            
            worksheet_att.column_dimensions['A'].width = 25
            for col in range(2, len(all_dates) + 2):
                col_letter = get_column_letter(col)
                worksheet_att.column_dimensions[col_letter].width = 12
            
            worksheet_summary = writer.sheets['Итоги']
            worksheet_summary.column_dimensions['A'].width = 35
            worksheet_summary.column_dimensions['B'].width = 20
            
            worksheet_stats.auto_filter.ref = worksheet_stats.dimensions
            worksheet_att.auto_filter.ref = worksheet_att.dimensions
        
        output.seek(0)
        
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
    
    if user.get('selected_lessons'):
        selected = sorted(user['selected_lessons'])
        lessons_text = ", ".join(map(str, selected))
        time_slots = "\n".join([f"   {i}. {LESSON_TIMES[i]}" for i in selected])
    else:
        lessons_text = "не выбраны"
        time_slots = "   не выбраны"
    
    try:
        all_students = cache.get_students()
        student_count = max(0, len(all_students) - 1)
    except:
        student_count = 0
    
    bot.send_message(message.chat.id,
                    f"⚙️ *Текущие настройки:*\n\n"
                    f"👥 *Группа:* {GROUP_NAME}\n"
                    f"👤 *Студентов:* {student_count}\n\n"
                    f"📅 *Дата:* {user['current_date']}\n"
                    f"🔢 *Выбранные пары:* {lessons_text}\n"
                    f"⏰ *Время пар:*\n{time_slots}\n\n"
                    f"*Изменить:*\n"
                    f"📅 - выбрать дату\n"
                    f"🔢 - выбрать пары\n"
                    f"📝 - отметить студентов",
                    parse_mode='Markdown')

# ==================== ЗАПУСК ====================
if __name__ == "__main__":
    print("=" * 60)
    print("🤖 Бот для учёта посещаемости ЗАПУЩЕН!")
    print("=" * 60)
    print(f"📍 Группа: {GROUP_NAME}")
    print(f"✅ Множественный выбор пар - АКТИВЕН")
    print(f"✅ Множественный выбор студентов - АКТИВЕН")
    print(f"✅ Обновление сообщений без удаления - АКТИВНО")
    print(f"✅ УЛУЧШЕННОЕ КЭШИРОВАНИЕ - АКТИВНО")
    print(f"✅ Кнопки 'Выбрать все' и 'Очистить все' - ИСПРАВЛЕНЫ")
    print(f"✅ Батчевые операции - АКТИВНЫ")
    print(f"📊 Отчёт: только прогулы выделены красным")
    print(f"📅 Расписание пар:")
    for i in range(1, 7):
        print(f"   {i}. {LESSON_TIMES[i]}")
    print("=" * 60)
    print("⚡ Статус: Ожидание команд...")
    print("=" * 60)
    
    try:
        while True:
    try:
        print("🔄 Запуск бота...")
        bot.polling(none_stop=False, interval=1, timeout=30)
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        print("🔄 Перезапуск через 10 секунд...")
        time.sleep(10)
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import time
        time.sleep(10)

