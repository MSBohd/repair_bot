import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
import openpyxl
import os

# Логування
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

TOKEN = os.getenv("TOKEN")
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
            "state": "start"
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
                            "name": str(row[0]),
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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    
    user_data["category"] = None
    user_data["model"] = None
    user_data["client_type"] = None
    user_data["selected_repairs"] = []
    user_data["repairs_list"] = []
    user_data["state"] = "category"
    
    keyboard = [
        ["📱 iPhone", "📱 iPad"],
        ["💻 MacBook", "⌚ Apple Watch"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await update.message.reply_text(
        "🍎 *Калькулятор ремонтів Apple*\n\n"
        "Оберіть категорію пристрою:",
        reply_markup=reply_markup,
        parse_mode='Markdown'
    )

async def show_category_keyboard(update):
    keyboard = [
        ["📱 iPhone", "📱 iPad"],
        ["💻 MacBook", "⌚ Apple Watch"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Оберіть категорію:", reply_markup=reply_markup)

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    text = update.message.text
    
    logger.info(f"Користувач {user_id} написав: {text}, стан: {user_data['state']}")
    
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
            await update.message.reply_text(f"❌ Не знайдено моделей для {text}")
            return
        
        keyboard = []
        for i in range(0, len(models), 2):
            keyboard.append(models[i:i+2])
        keyboard.append(["◀️ Назад"])
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        await update.message.reply_text(
            f"*{text}*\n\nОберіть модель:",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )
    
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
            await update.message.reply_text("❌ Не знайдено ремонтів для цієї моделі.")
            return
        
        opt_available = OPT_AVAILABLE.get(user_data["category"], False)
        
        if opt_available:
            keyboard = [["👤 РОЗДРІБ", "🏢 ОПТ"], ["◀️ Назад"]]
        else:
            keyboard = [["👤 РОЗДРІБ"], ["◀️ Назад"]]
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        await update.message.reply_text(
            f"*{user_data['category']}: {text}*\n\n"
            "Оберіть тип клієнта:",
            reply_markup=reply_markup,
            parse_mode='Markdown'
        )
    
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
            keyboard.append(["◀️ Назад"])
            reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            await update.message.reply_text("Оберіть модель:", reply_markup=reply_markup)
            return
        
        if text == "🏢 ОПТ":
            opt_available = OPT_AVAILABLE.get(user_data["category"], False)
            if not opt_available:
                await update.message.reply_text(
                    "⚠️ *Оптові ціни для цього пристрою поки що не доступні*\n\n"
                    "Зверніться до менеджера для отримання оптових цін.",
                    parse_mode='Markdown'
                )
                return
            user_data["client_type"] = "opt"
        
        elif text == "👤 РОЗДРІБ":
            user_data["client_type"] = "client"
        
        else:
            return
        
        user_data["state"] = "repairs"
        user_data["selected_repairs"] = []
        
        await update.message.reply_text(
            "Зараз покажу ремонти...",
            reply_markup=ReplyKeyboardRemove()
        )
        
        await show_repairs_menu_message(update, user_data)

async def show_repairs_menu_message(update, user_data):
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    # ТІЛЬКИ ЗАГОЛОВОК - без списку ремонтів
    text = f"*{user_data['category']}: {user_data['model']}*\n"
    text += f"Тип: *{client_label}*\n\n"
    text += "Оберіть ремонти:"
    
    keyboard = []
    for idx, repair in enumerate(repairs):
        price = repair[client_type]
        
        # ПРОПУСКАЄМО якщо ціна 0 в опті
        if price == 0:
            continue
        
        check = "✓" if idx in user_data["selected_repairs"] else "○"
        # ЦІНА В КНОПЦІ
        button_text = f"{check} {repair['name']}: {price:.0f} {currency}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"repair_{idx}")])
    
    control_buttons = []
    if user_data["selected_repairs"]:
        control_buttons.append(InlineKeyboardButton("💰 Розрахувати", callback_data="calculate"))
        control_buttons.append(InlineKeyboardButton("🔄 Скинути", callback_data="reset_repairs"))
    
    if control_buttons:
        keyboard.append(control_buttons)
    
    keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_client")])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(text, reply_markup=reply_markup, parse_mode='Markdown')


async def show_repairs_menu_inline(query, user_data):
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    # ТІЛЬКИ ЗАГОЛОВОК - без списку ремонтів
    text = f"*{user_data['category']}: {user_data['model']}*\n"
    text += f"Тип: *{client_label}*\n\n"
    text += "Оберіть ремонти:"
    
    keyboard = []
    for idx, repair in enumerate(repairs):
        price = repair[client_type]
        
        # ПРОПУСКАЄМО якщо ціна 0 в опті
        if price == 0:
            continue
        
        check = "✓" if idx in user_data["selected_repairs"] else "○"
        # ЦІНА В КНОПЦІ
        button_text = f"{check} {repair['name']}: {price:.0f} {currency}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"repair_{idx}")])
    
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

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = update.effective_user.id
    user_data = get_user_data(user_id)
    data = query.data
    
    if data.startswith("repair_"):
        repair_idx = int(data.replace("repair_", ""))
        if repair_idx in user_data["selected_repairs"]:
            user_data["selected_repairs"].remove(repair_idx)
        else:
            user_data["selected_repairs"].append(repair_idx)
        await show_repairs_menu_inline(query, user_data)
    
    elif data == "calculate":
        await calculate_result(query, user_data)
    
    elif data == "reset_repairs":
        user_data["selected_repairs"] = []
        await show_repairs_menu_inline(query, user_data)
    
    elif data == "back_to_client":
        user_data["state"] = "client_type"
        user_data["selected_repairs"] = []
        
        opt_available = OPT_AVAILABLE.get(user_data["category"], False)
        if opt_available:
            keyboard = [["👤 РОЗДРІБ", "🏢 ОПТ"], ["◀️ Назад"]]
        else:
            keyboard = [["👤 РОЗДРІБ"], ["◀️ Назад"]]
        
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        await query.message.reply_text("Оберіть тип клієнта:", reply_markup=reply_markup)
        await query.message.delete()

async def show_repairs_menu_inline(query, user_data):
    repairs = user_data["repairs_list"]
    client_type = user_data["client_type"]
    client_label = "РОЗДРІБ" if client_type == "client" else "ОПТ"
    currency = get_currency(client_type)
    
    text = f"*{user_data['category']}: {user_data['model']}*\n"
    text += f"Тип: *{client_label}*\n\n"
    text += "Оберіть ремонти:\n\n"
    
    keyboard = []
    for idx, repair in enumerate(repairs):
        price = repair[client_type]
        check = "☑️" if idx in user_data["selected_repairs"] else "☐"
        text += f"{check} {repair['name']}: {price} {currency}\n"
        button_text = f"{'✓' if idx in user_data['selected_repairs'] else '○'} {repair['name']}"
        keyboard.append([InlineKeyboardButton(button_text, callback_data=f"repair_{idx}")])
    
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

async def calculate_result(query, user_data):
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
        details.append(f"• {repair['name']}: {price} {currency}")
    
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
        ["💻 MacBook", "⌚ Apple Watch"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await query.message.reply_text(text, parse_mode='Markdown')
    await query.message.reply_text(
        "🔄 Новий розрахунок?\nОберіть категорію:",
        reply_markup=reply_markup
    )
    
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

    main()
