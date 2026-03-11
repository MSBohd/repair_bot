import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
import openpyxl
import os
import asyncio

# Логування
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.environ.get("BOT_TOKEN")

# ПАРОЛЬ ДЛЯ ОПТУ
OPT_PASSWORD = "4343"

# ШЛЯХ ДО ПАПКИ З ФАЙЛАМИ
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Файли з цінами
FILES = {
    "📱 iPhone": "iphones.xlsx",
    "📱 iPad": "ipads.xlsx",
    "💻 MacBook": "macbooks.xlsx",
    "⌚ Apple Watch": "watches.xlsx"
}

# ДЛЯ ЯКИХ КАТЕГОРІЙ ДОСТУПНІ ОПТОВІ ЦІНИ
OPT_AVAILABLE = {
    "📱 iPhone": True,
    "📱 iPad": False,
    "💻 MacBook": False,
    "⌚ Apple Watch": False,
}

# Тестові дані
TEST_DATA = {
    "📱 iPhone": {
        "iPhone X": [
            {"name": "Екран", "client": 3800, "opt": 2000},
            {"name": "Батарея", "client": 1400, "opt": 900},
            {"name": "Камера", "client": 2000, "opt": 1500},
        ],
        "iPhone 11": [
            {"name": "Екран", "client": 4200, "opt": 2200},
            {"name": "Батарея", "client": 1600, "opt": 1000},
            {"name": "Камера", "client": 2200, "opt": 1700},
        ]
    }
}

# Збереження стану користувача
user_data_store = {}

def get_user_data(user_id):
    if user_id not in user_data_store:
        user_data_store[user_id] = {
            "category": None,
            "model": None,
            "client_type": None,
            "selected_repairs": [],
            "repairs_list": [],
            "state": "start",
            "opt_unlocked": False,  # ТЕПЕР ЗБЕРІГАЄТЬСЯ НАЗАВЖДИ
            "message_ids": []
        }
    return user_data_store[user_id]

def get_currency(client_type):
    """Повертає валюту в залежності від типу клієнта"""
    return "$" if client_type == "opt" else "грн"

def load_models(category):
    filename = FILES.get(category)
    
    if filename:
        filepath = os.path.join(BASE_DIR, filename)
        if os.path.exists(filepath):
            try:
                wb = openpyxl.load_workbook(filepath)
                models = wb.sheetnames
                wb.close()
                logger.info(f"Завантажено моделі з {filename}: {models}")
                return models
            except Exception as e:
                logger.error(f"Помилка читання {filename}: {e}")
    
    if category in TEST_DATA:
        models = list(TEST_DATA[category].keys())
        logger.info(f"Використано тестові дані для {category}: {models}")
        return models
    
    return []

def load_repairs(category, model):
    filename = FILES.get(category)
    
    if filename:
        filepath = os.path.join(BASE_DIR, filename)
        if os.path.exists(filepath):
            try:
                wb = openpyxl.load_workbook(filepath)
                ws = wb[model]
                
                repairs = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[0]:
                        repairs.append({
                            "name": str(row[0]).strip(),  # ВИПРАВЛЕНО - додав .strip()
                            "client": float(row[1]) if row[1] else 0,
                            "opt": float(row[2]) if row[2] else 0
                        })
                
                wb.close()
                logger.info(f"Завантажено {len(repairs)} ремонтів для {model}")
                return repairs
            except Exception as e:
                logger.error(f"Помилка читання ремонтів: {e}")
    
    if category in TEST_DATA and model in TEST_DATA[category]:
        return TEST_DATA[category][model]
    
    return []

def get_available_repairs(repairs, client_type):
    """Повертає тільки ремонти з ціною > 0"""
    available = []
    for idx, repair in enumerate(repairs):
        price = repair[client_type]
        if price > 0:
            available.append((idx, repair))
    return available

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    
    user_data["category"] = None
    user_data["model"] = None
    user_data["client_type"] = None
    user_data["selected_repairs"] = []
    user_data["repairs_list"] = []
    user_data["state"] = "category"
    user_data["message_ids"] = []
    
    keyboard = [
        ["📱 iPhone", "📱 iPad"],
        ["💻 MacBook", "⌚ Apple Watch"],
        ["🗑️ Очистити чат"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    msg = await update.message.reply_text(
        "🍎 *Калькулятор ремонтів Apple*\n\n"
        "Оберіть категорію пристрою:",
        reply_markup=reply_markup,
        parse_mode='Markdown'
    )
    user_data["message_ids"].append(msg.message_id)

async def show_category_keyboard(update):
    keyboard = [
        ["📱 iPhone", "📱 iPad"],
        ["💻 MacBook", "⌚ Apple Watch"],
        ["🗑️ Очистити чат"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Оберіть категорію:", reply_markup=reply_markup)

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    text = update.message.text
    
    logger.info(f"Користувач {user_id} написав: {text}, стан: {user_data['state']}")
    
    # ========================
    # ОЧИЩЕННЯ ЧАТУ
    # ========================
    if text == "🗑️ Очистити чат":
        # Видаляємо всі збережені повідомлення бота
        for msg_id in user_data.get("message_ids", []):
            try:
                await context.bot.delete_message(
                    chat_id=update.effective_chat.id,
                    message_id=msg_id
                )
            except Exception as e:
                logger.warning(f"Не вдалось видалити повідомлення {msg_id}: {e}")
        
        # Видаляємо повідомлення користувача з командою
        try:
            await context.bot.delete_message(
                chat_id=update.effective_chat.id,
                message_id=update.message.message_id
            )
        except:
            pass
        
        # ПОВНІСТЮ ОЧИЩАЄМО ВСІ ДАНІ
        user_data.clear()
        user_data.update({
            "category": None,
            "model": None,
            "client_type": None,
            "selected_repairs": [],
            "repairs_list": [],
            "state": "category",
            "opt_unlocked": False,
            "message_ids": []
        })
        
        keyboard = [
            ["📱 iPhone", "📱 iPad"],
            ["💻 MacBook", "⌚ Apple Watch"],
            ["🗑️ Очистити чат"]
        ]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        
        msg = await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="🧹 *Чат очищено!*\n\n"
                 "🍎 *Калькулятор ремонтів Apple*\n\n"
                 "Оберіть категорію пристрою:",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )
        user_data["message_ids"].append(msg.message_id)
        return
    
    # ========================
    # ВВЕДЕННЯ ПАРОЛЯ ДЛЯ ОПТУ
    # ========================
    if user_data["state"] == "waiting_password":
        if text == OPT_PASSWORD:
            user_data["opt_unlocked"] = True
            user_data["client_type"] = "opt"
            user_data["state"] = "repairs"
            user_data["selected_repairs"] = []
            
            msg = await update.message.reply_text(
                "✅ *Пароль прийнято!*\n\nЗараз покажу ремонти...",
                parse_mode='Markdown',
                reply_markup=ReplyKeyboardRemove()
            )
            user_data["message_ids"].append(msg.message_id)
            await show_repairs_menu_message(update, user_data)
        else:
            # НЕПРАВИЛЬНИЙ ПАРОЛЬ - ПОВЕРТАЄМО НА ПОЧАТОК
            user_data.clear()
            user_data.update({
                "category": None,
                "model": None,
                "client_type": None,
                "selected_repairs": [],
                "repairs_list": [],
                "state": "category",
                "opt_unlocked": False,
                "message_ids": []
            })
            
            keyboard = [
                ["📱 iPhone", "📱 iPad"],
                ["💻 MacBook", "⌚ Apple Watch"],
                ["🗑️ Очистити чат"]
            ]
            reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            
            msg = await update.message.reply_text(
                "❌ *Невірний пароль!*\n\n"
                "Повертаємось на початок.\n\n"
                "🍎 *Калькулятор ремонтів Apple*\n\n"
                "Оберіть категорію пристрою:",
                reply_markup=reply_markup,
                parse_mode='Markdown'
            )
            user_data["message_ids"].append(msg.message_id)
        return
    
    # ========================
    # ВИБІР КАТЕГОРІЇ
    # ========================
    if user_data["state"] == "category" and text in FILES.keys():
        user_data["category"] = text
        user_data["model"] = None
        user_data["selected_repairs"] = []
        user_data["state"] = "model"
        
        models = load_models(text)
        
        if not models:
            msg = await update.message.reply_text(f"❌ Не знайдено моделей для {text}")
            user_data["message_ids"].append(msg.message_id)
            return
        
        keyboard = []
        for i in range(0, len(models), 2):
            keyboard.append(models[i:i+2])
        keyboard.append(["◀️ Назад", "🗑️ Очистити чат"])
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        msg = await update.message.reply_text(
            f"*{text}*\n\nОберіть модель:",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )
        user_data["message_ids"].append(msg.message_id)
    
    # ========================
    # ВИБІР МОДЕЛІ
    # ========================
    elif user_data["state"] == "model":
        if text == "◀️ Назад":
            user_data["state"] = "category"
            await show_category_keyboard(update)
            return
        
        user_data["model"] = text
        user_data["state"] = "client_type"
        
        repairs = load_repairs(user_data["category"], text)
        user_data["repairs_list"] = repairs
        
        if not repairs:
            msg = await update.message.reply_text("❌ Не знайдено ремонтів для цієї моделі.")
            user_data["message_ids"].append(msg.message_id)
            return
        
        opt_available = OPT_AVAILABLE.get(user_data["category"], False)
        
        if opt_available:
            keyboard = [
                ["👤 РОЗДРІБ", "🏢 ОПТ"],
                ["◀️ Назад", "🗑️ Очистити чат"]
            ]
        else:
            keyboard = [
                ["👤 РОЗДРІБ"],
                ["◀️ Назад", "🗑️ Очистити чат"]
            ]
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        msg = await update.message.reply_text(
            f"*{user_data['category']}: {text}*\n\n"
            "Оберіть тип клієнта:",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )
        user_data["message_ids"].append(msg.message_id)
    
    # ========================
    # ВИБІР ТИПУ КЛІЄНТА
    # ========================
    elif user_data["state"] == "client_type":
        if text == "◀️ Назад":
            user_data["state"] = "model"
            models = load_models(user_data["category"])
            keyboard = []
            for i in range(0, len(models), 2):
                keyboard.append(models[i:i+2])
            keyboard.append(["◀️ Назад", "🗑️ Очистити чат"])
            reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            msg = await update.message.reply_text("Оберіть модель:", reply_markup=reply_markup)
            user_data["message_ids"].append(msg.message_id)
            return
        
        if text == "🏢 ОПТ":
            opt_available = OPT_AVAILABLE.get(user_data["category"], False)
            if not opt_available:
                msg = await update.message.reply_text(
                    "⚠️ *Оптові ціни для цього пристрою поки що не доступні*\n\n"
                    "Зверніться до менеджера для отримання оптових цін.",
                    parse_mode='Markdown'
                )
                user_data["message_ids"].append(msg.message_id)
                return
            
            # ПЕРЕВІРКА ПАРОЛЯ
            if not user_data["opt_unlocked"]:
                user_data["state"] = "waiting_password"
                
                keyboard = [["🗑️ Очистити чат"]]
                reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
                
                msg = await update.message.reply_text(
                    "🔐 *Доступ до оптових цін*\n\n"
                    "Введіть пароль:",
                    reply_markup=reply_markup,
                    parse_mode='Markdown'
                )
                user_data["message_ids"].append(msg.message_id)
                return
            
            user_data["client_type"] = "opt"
        
        elif text == "👤 РОЗДРІБ":
            user_data["client_type"] = "client"
        
        else:
            return
        
        user_data["state"] = "repairs"
        user_data["selected_repairs"] = []
        
        msg = await update.message.reply_text(
            "Зараз покажу ремонти...",
            reply_markup=ReplyKeyboardRemove()
        )
        user_data["message_ids"].append(msg.message_id)
        
        await show_repairs_menu_message(update, user_data)

async def show_repairs_menu_message(update, user_data):
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    available_repairs = get_available_repairs(repairs, client_type)
    
    if not available_repairs:
        msg = await update.message.reply_text("❌ Немає доступних ремонтів для цього типу клієнта.")
        user_data["message_ids"].append(msg.message_id)
        return
    
    # ЗАГОЛОВОК + СПИСОК З ЦІНАМИ
    text = f"*{user_data['category']}: {user_data['model']}*\n"
    text += f"Тип: *{client_label}*\n\n"
    
    # Список ремонтів з цінами
    text += "💰 *Прайс-лист:*\n"
    text += "━━━━━━━━━━━━━━━━━━\n"
    for original_idx, repair in available_repairs:
        price = repair[client_type]
        check = "☑️" if original_idx in user_data["selected_repairs"] else "☐"
        
        # ПЕРЕНІС ДОВГИХ НАЗВ - ВИПРАВЛЕНО
        repair_name = repair['name'].strip()  # Прибираємо зайві пробіли
        if len(repair_name) > 30:
            # Розбиваємо по словах для кращого вигляду
            words = repair_name.split()
            line1 = ""
            line2 = ""
            for word in words:
                if len(line1 + word) < 30:
                    line1 += word + " "
                else:
                    line2 += word + " "
            text += f"{check} {line1.strip()}\n"
            if line2:
                text += f"    {line2.strip()}\n"
        else:
            text += f"{check} {repair_name}\n"
        
        text += f"    💵 {price:.0f} {currency}\n"
        text += "━━━━━━━━━━━━━━━━━━\n"
    
    text += "\n✅ Оберіть потрібні ремонти:"
    
    # КНОПКИ БЕЗ ЦІН
    keyboard = []
    for original_idx, repair in available_repairs:
        check = "✓" if original_idx in user_data["selected_repairs"] else "○"
        repair_name = repair['name'].strip()
        # Скорочуємо назву в кнопці якщо дуже довга
        if len(repair_name) > 35:
            repair_name = repair_name[:32] + "..."
        button_text = f"{check} {repair_name}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"repair_{original_idx}")])
    
    control_buttons = []
    if user_data["selected_repairs"]:
        control_buttons.append(InlineKeyboardButton("💰 Розрахувати", callback_data="calculate"))
        control_buttons.append(InlineKeyboardButton("🔄 Скинути", callback_data="reset_repairs"))
    
    if control_buttons:
        keyboard.append(control_buttons)
    
    keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_client")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    msg = await update.message.reply_text(text, reply_markup=reply_markup, parse_mode='Markdown')
    user_data["message_ids"].append(msg.message_id)

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    data = query.data
    
    if data.startswith("repair_"):
        repair_idx = int(data.replace("repair_", ""))
        
        repairs = user_data["repairs_list"]
        client_type = user_data["client_type"]
        
        if repair_idx < len(repairs):
            price = repairs[repair_idx][client_type]
            if price > 0:
                if repair_idx in user_data["selected_repairs"]:
                    user_data["selected_repairs"].remove(repair_idx)
                else:
                    user_data["selected_repairs"].append(repair_idx)
        
        await show_repairs_menu_inline(query, user_data)
    
    elif data == "calculate":
        await calculate_result(query, user_data, context)
    
    elif data == "reset_repairs":
        user_data["selected_repairs"] = []
        await show_repairs_menu_inline(query, user_data)
    
    elif data == "back_to_client":
        user_data["state"] = "client_type"
        user_data["selected_repairs"] = []
        
        opt_available = OPT_AVAILABLE.get(user_data["category"], False)
        if opt_available:
            keyboard = [
                ["👤 РОЗДРІБ", "🏢 ОПТ"],
                ["◀️ Назад", "🗑️ Очистити чат"]
            ]
        else:
            keyboard = [
                ["👤 РОЗДРІБ"],
                ["◀️ Назад", "🗑️ Очистити чат"]
            ]
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        msg = await query.message.reply_text("Оберіть тип клієнта:", reply_markup=reply_markup)
        user_data["message_ids"].append(msg.message_id)
        await query.message.delete()

async def show_repairs_menu_inline(query, user_data):
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    available_repairs = get_available_repairs(repairs, client_type)
    
    # ЗАГОЛОВОК + СПИСОК З ЦІНАМИ
    text = f"*{user_data['category']}: {user_data['model']}*\n"
    text += f"Тип: *{client_label}*\n\n"
    
    # Список ремонтів з цінами
    text += "💰 *Прайс-лист:*\n"
    text += "━━━━━━━━━━━━━━━━━━\n"
    for original_idx, repair in available_repairs:
        price = repair[client_type]
        check = "☑️" if original_idx in user_data["selected_repairs"] else "☐"
        
        # ПЕРЕНІС ДОВГИХ НАЗВ - ВИПРАВЛЕНО
        repair_name = repair['name'].strip()
        if len(repair_name) > 30:
            words = repair_name.split()
            line1 = ""
            line2 = ""
            for word in words:
                if len(line1 + word) < 30:
                    line1 += word + " "
                else:
                    line2 += word + " "
            text += f"{check} {line1.strip()}\n"
            if line2:
                text += f"    {line2.strip()}\n"
        else:
            text += f"{check} {repair_name}\n"
        
        text += f"    💵 {price:.0f} {currency}\n"
        text += "━━━━━━━━━━━━━━━━━━\n"
    
    text += "\n✅ Оберіть потрібні ремонти:"
    
    # КНОПКИ БЕЗ ЦІН
    keyboard = []
    for original_idx, repair in available_repairs:
        check = "✓" if original_idx in user_data["selected_repairs"] else "○"
        repair_name = repair['name'].strip()
        if len(repair_name) > 35:
            repair_name = repair_name[:32] + "..."
        button_text = f"{check} {repair_name}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"repair_{original_idx}")])
    
    control_buttons = []
    if user_data["selected_repairs"]:
        control_buttons.append(InlineKeyboardButton("💰 Розрахувати", callback_data="calculate"))
        control_buttons.append(InlineKeyboardButton("🔄 Скинути", callback_data="reset_repairs"))
    
    if control_buttons:
        keyboard.append(control_buttons)
    
    keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_client")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    try:
        await query.edit_message_text(text, reply_markup=reply_markup, parse_mode='Markdown')
    except Exception as e:
        logger.warning(f"Throttling при оновленні меню: {e}")

async def calculate_result(query, user_data, context):
    if not user_data["selected_repairs"]:
        await query.answer("⚠️ Оберіть хоча б один ремонт!", show_alert=True)
        return
    
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    total = 0
    details = []
    
    for idx in user_data["selected_repairs"]:
        repair = repairs[idx]
        price = repair[client_type]
        total += price
        details.append(f"• {repair['name']}: {price:.0f} {currency}")
    
    num_repairs = len(user_data["selected_repairs"])
    discount_percent = 0
    if num_repairs == 2:
        discount_percent = 5
    elif num_repairs == 3:
        discount_percent = 10
    elif num_repairs >= 4:
        discount_percent = 15
    
    discount_amount = total * (discount_percent / 100)
    final = total - discount_amount
    
    text = "━━━━━━━━━━━━━━━━━━━━\n"
    text += "💰 *РЕЗУЛЬТАТ РОЗРАХУНКУ*\n"
    text += "━━━━━━━━━━━━━━━━━━━━\n\n"
    text += f"*{user_data['model']}* | {client_label}\n\n"
    text += f"Обрано ремонтів: *{num_repairs}*\n\n"
    
    for detail in details:
        text += f"{detail}\n"
    
    text += "\n━━━━━━━━━━━━━━━━━━━━\n"
    text += f"Сума: *{total:.0f} {currency}*\n"
    
    if discount_percent > 0:
        text += f"Знижка {discount_percent}%: *-{discount_amount:.0f} {currency}*\n"
    
    text += "━━━━━━━━━━━━━━━━━━━━\n"
    text += f"💰 *ВСЬОГО: {final:.0f} {currency}*\n"
    text += "━━━━━━━━━━━━━━━━━━━━"
    
    await query.message.delete()
    
    keyboard = [
        ["📱 iPhone", "📱 iPad"],
        ["💻 MacBook", "⌚ Apple Watch"],
        ["🗑️ Очистити чат"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    msg1 = await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=text,
        parse_mode='Markdown'
    )
    msg2 = await context.bot.send_message(
        chat_id=query.message.chat_id,
        text="🔄 Новий розрахунок?\nОберіть категорію:",
        reply_markup=reply_markup
    )
    
    user_data["message_ids"] = [msg1.message_id, msg2.message_id]
    user_data.update({
        "state": "category",
        "category": None,
        "model": None,
        "client_type": None,
        "selected_repairs": [],
        "repairs_list": []
    })

def main():
    application = Application.builder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, message_handler))
    application.add_handler(CallbackQueryHandler(button_handler))
    
    print("🤖 Бот запущено!")
    print(f"📁 Робоча папка: {BASE_DIR}")
    for category, filename in FILES.items():
        filepath = os.path.join(BASE_DIR, filename)
        opt = "✅ опт" if OPT_AVAILABLE.get(category) else "❌ без опту"
        status = "✅" if os.path.exists(filepath) else "⚠️ "
        print(f"  {status} {filename} ({opt})")
    
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    asyncio.set_event_loop(asyncio.new_event_loop())
    main()

