"""
CRM-Помощник — Telegram-бот поддержки отдела 1С CRM.
Сбор ошибок и предложений с записью в Excel.
"""

import json
import os
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from config import (
    ADMIN_IDS,
    BOT_TOKEN,
    ERROR_CATEGORIES,
    EXCEL_FILE,
    MODULES,
    USERS_DB_FILE,
)

# ── Состояния ConversationHandler ──────────────────────────────────────
(
    REG_FIO,
    REG_MODULE,
    MAIN_MENU,
    ERROR_CATEGORY,
    ERROR_DESCRIPTION,
    SUGGESTION_TEXT,
) = range(6)

# ── Эмодзи для модулей ────────────────────────────────────────────────
MODULE_EMOJI = {
    "Воронка продаж": "📊",
    "Карточка клиента": "👤",
    "Карточка интереса": "💡",
    "Телефония": "📞",
    "Почта": "📧",
    "Отчёты и аналитика": "📈",
}

# ── Эмодзи для категорий ошибок ───────────────────────────────────────
ERROR_EMOJI = {
    "Воронка продаж": "📊",
    "Проблема с карточкой клиента": "👤",
    "Проблема с карточкой интереса": "💡",
    "Другое": "🔧",
}

# ── Вспомогательные функции для хранения пользователей ─────────────────

def _ensure_data_dir():
    Path(USERS_DB_FILE).parent.mkdir(parents=True, exist_ok=True)
    Path(EXCEL_FILE).parent.mkdir(parents=True, exist_ok=True)


def _load_users() -> dict:
    _ensure_data_dir()
    if os.path.exists(USERS_DB_FILE):
        with open(USERS_DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_users(users: dict):
    _ensure_data_dir()
    with open(USERS_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def _get_user(user_id: int) -> dict | None:
    users = _load_users()
    return users.get(str(user_id))


def _save_user(user_id: int, fio: str, module: str):
    users = _load_users()
    users[str(user_id)] = {
        "fio": fio,
        "module": module,
        "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    _save_users(users)


# ── Excel ──────────────────────────────────────────────────────────────

EXCEL_HEADERS = [
    "Дата и время",
    "Telegram ID",
    "ФИО",
    "Модуль",
    "Тип обращения",
    "Категория ошибки",
    "Описание",
]


def _ensure_excel():
    _ensure_data_dir()
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Обращения"
        ws.append(EXCEL_HEADERS)
        # Ширина столбцов
        widths = [20, 14, 25, 22, 20, 30, 60]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w
        wb.save(EXCEL_FILE)


def _append_to_excel(row: list):
    _ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append(row)
    wb.save(EXCEL_FILE)


# ── Клавиатуры ─────────────────────────────────────────────────────────

def _main_menu_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton("🐞 Сообщить об ошибке", callback_data="report_error")],
        [InlineKeyboardButton("💎 Предложить улучшение", callback_data="suggest")],
    ]
    return InlineKeyboardMarkup(buttons)


def _modules_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(
            f"{MODULE_EMOJI.get(m, '📁')} {m}",
            callback_data=f"module:{m}",
        )]
        for m in MODULES
    ]
    return InlineKeyboardMarkup(buttons)


def _error_categories_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(
            f"{ERROR_EMOJI.get(c, '❓')} {c}",
            callback_data=f"errcat:{c}",
        )]
        for c in ERROR_CATEGORIES
    ]
    buttons.append(
        [InlineKeyboardButton("🔙 Назад в меню", callback_data="back_menu")]
    )
    return InlineKeyboardMarkup(buttons)


def _back_to_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("🏠 Вернуться в меню", callback_data="back_menu")]]
    )


# ── Хендлеры ───────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Точка входа: /start. Проверяем регистрацию."""
    user_id = update.effective_user.id
    user = _get_user(user_id)

    if user:
        return await _show_main_menu(update, context, user)

    await update.message.reply_text(
        "👋 Добро пожаловать в CRM-Помощник!\n\n"
        "🤖 Я помогу вам быстро сообщить об ошибке "
        "или предложить улучшение для 1С CRM.\n\n"
        "📝 Для начала давайте познакомимся.\n"
        "Введите ваше ФИО:"
    )
    return REG_FIO


async def reg_fio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получаем ФИО, предлагаем выбрать модуль."""
    fio = update.message.text.strip()
    if len(fio) < 3:
        await update.message.reply_text(
            "⚠️ Пожалуйста, введите корректное ФИО:"
        )
        return REG_FIO

    context.user_data["reg_fio"] = fio
    await update.message.reply_text(
        f"✅ Отлично, {fio}!\n\n"
        "🔽 Выберите модуль 1С CRM, с которым вы работаете:",
        reply_markup=_modules_keyboard(),
    )
    return REG_MODULE


async def reg_module(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Сохраняем модуль, показываем главное меню."""
    query = update.callback_query
    await query.answer()

    module = query.data.removeprefix("module:")
    fio = context.user_data.pop("reg_fio")
    user_id = update.effective_user.id

    _save_user(user_id, fio, module)

    user = _get_user(user_id)
    return await _show_main_menu_from_callback(query, context, user)


async def _show_main_menu(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    user: dict,
) -> int:
    """Главное меню (из обычного сообщения)."""
    emoji = MODULE_EMOJI.get(user["module"], "📁")
    await update.message.reply_text(
        f"👋 Здравствуйте, {user['fio']}!\n"
        f"{emoji} Ваш модуль: {user['module']}\n\n"
        "⬇️ Чем могу помочь?",
        reply_markup=_main_menu_keyboard(),
    )
    return MAIN_MENU


async def _show_main_menu_from_callback(query, context, user: dict) -> int:
    """Главное меню (из callback-кнопки)."""
    emoji = MODULE_EMOJI.get(user["module"], "📁")
    await query.edit_message_text(
        f"👋 Здравствуйте, {user['fio']}!\n"
        f"{emoji} Ваш модуль: {user['module']}\n\n"
        "⬇️ Чем могу помочь?",
        reply_markup=_main_menu_keyboard(),
    )
    return MAIN_MENU


async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Обработка нажатий главного меню."""
    query = update.callback_query
    await query.answer()

    if query.data == "report_error":
        await query.edit_message_text(
            "🐞 <b>Сообщить об ошибке</b>\n\n"
            "Выберите категорию проблемы:",
            reply_markup=_error_categories_keyboard(),
            parse_mode="HTML",
        )
        return ERROR_CATEGORY

    if query.data == "suggest":
        await query.edit_message_text(
            "💎 <b>Предложить улучшение</b>\n\n"
            "Мы ценим ваши идеи! ✨\n\n"
            "Опишите, что, на ваш взгляд, можно улучшить в системе.\n"
            "Любая деталь может оказаться полезной 👇",
            parse_mode="HTML",
        )
        return SUGGESTION_TEXT

    return MAIN_MENU


async def error_category_handler(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Выбор категории ошибки."""
    query = update.callback_query
    await query.answer()

    if query.data == "back_menu":
        user = _get_user(update.effective_user.id)
        return await _show_main_menu_from_callback(query, context, user)

    category = query.data.removeprefix("errcat:")
    context.user_data["error_category"] = category

    if category == "Другое":
        await query.edit_message_text(
            "🔧 <b>Категория: Другое</b>\n\n"
            "Расскажите подробнее, с какой проблемой вы столкнулись.\n"
            "Постарайтесь описать шаги, которые привели к ошибке — "
            "это поможет нам разобраться быстрее 🔍",
            parse_mode="HTML",
        )
        return ERROR_DESCRIPTION

    cat_emoji = ERROR_EMOJI.get(category, "❓")
    await query.edit_message_text(
        f"{cat_emoji} <b>Категория: {category}</b>\n\n"
        "Опишите проблему:\n"
        "• Что произошло?\n"
        "• При каких действиях?\n"
        "• Есть ли скриншоты? 👇",
        parse_mode="HTML",
    )
    return ERROR_DESCRIPTION


async def error_description_handler(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Получаем описание ошибки, сохраняем в Excel."""
    user_id = update.effective_user.id
    user = _get_user(user_id)
    category = context.user_data.pop("error_category", "—")
    description = update.message.text.strip()

    _append_to_excel([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        user_id,
        user["fio"],
        user["module"],
        "Ошибка",
        category,
        description,
    ])

    await update.message.reply_text(
        "✅ <b>Принято в работу!</b>\n\n"
        "Спасибо, что сообщили о проблеме! 🙏\n"
        "Наша команда уже в курсе — разберёмся "
        "и постараемся исправить как можно скорее. 🚀",
        reply_markup=_back_to_menu_keyboard(),
        parse_mode="HTML",
    )
    return MAIN_MENU


async def suggestion_text_handler(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Получаем предложение, сохраняем в Excel."""
    user_id = update.effective_user.id
    user = _get_user(user_id)
    description = update.message.text.strip()

    _append_to_excel([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        user_id,
        user["fio"],
        user["module"],
        "Предложение",
        "—",
        description,
    ])

    await update.message.reply_text(
        "✅ <b>Предложение принято!</b>\n\n"
        "Благодарим за вашу инициативу! 💪\n"
        "Каждая идея помогает сделать систему удобнее для всех.\n"
        "Мы обязательно рассмотрим ваше предложение. ⭐",
        reply_markup=_back_to_menu_keyboard(),
        parse_mode="HTML",
    )
    return MAIN_MENU


async def back_to_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Возврат в главное меню из callback."""
    query = update.callback_query
    await query.answer()
    user = _get_user(update.effective_user.id)
    if user:
        return await _show_main_menu_from_callback(query, context, user)
    await query.edit_message_text("⚠️ Нажмите /start для начала.")
    return ConversationHandler.END


# ── Админские команды ──────────────────────────────────────────────────

async def cmd_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показать админ-панель."""
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("🔒 У вас нет доступа к этой команде.")
        return

    buttons = [
        [InlineKeyboardButton("📥 Выгрузить Excel", callback_data="admin:export")],
        [InlineKeyboardButton("📊 Статистика", callback_data="admin:stats")],
        [InlineKeyboardButton("👥 Список пользователей", callback_data="admin:users")],
    ]
    await update.message.reply_text(
        "⚙️ <b>Панель администратора</b>",
        reply_markup=InlineKeyboardMarkup(buttons),
        parse_mode="HTML",
    )


async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка админских кнопок."""
    query = update.callback_query
    await query.answer()

    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await query.edit_message_text("🔒 У вас нет доступа.")
        return

    action = query.data.removeprefix("admin:")

    if action == "export":
        _ensure_excel()
        if os.path.exists(EXCEL_FILE):
            await query.message.reply_document(
                document=open(EXCEL_FILE, "rb"),
                filename=f"crm_support_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                caption="📥 Выгрузка обращений",
            )
        else:
            await query.edit_message_text("📭 Файл обращений пока пуст.")

    elif action == "stats":
        _ensure_excel()
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            total = ws.max_row - 1
            errors = sum(1 for row in ws.iter_rows(min_row=2) if row[4].value == "Ошибка")
            suggestions = sum(1 for row in ws.iter_rows(min_row=2) if row[4].value == "Предложение")
            users = _load_users()
            text = (
                "📊 <b>Статистика обращений</b>\n\n"
                f"📋 Всего обращений: <b>{total}</b>\n"
                f"🐞 Ошибок: <b>{errors}</b>\n"
                f"💎 Предложений: <b>{suggestions}</b>\n"
                f"👥 Пользователей: <b>{len(users)}</b>"
            )
        else:
            text = "📭 Данных пока нет."
        await query.edit_message_text(text, parse_mode="HTML")

    elif action == "users":
        users = _load_users()
        if not users:
            await query.edit_message_text("📭 Зарегистрированных пользователей нет.")
            return
        lines = []
        for uid, info in users.items():
            m_emoji = MODULE_EMOJI.get(info["module"], "📁")
            lines.append(f"👤 {info['fio']} | {m_emoji} {info['module']} | ID: <code>{uid}</code>")
        text = "👥 <b>Пользователи</b>\n\n" + "\n".join(lines)
        if len(text) > 4000:
            text = text[:4000] + "\n\n... (список обрезан)"
        await query.edit_message_text(text, parse_mode="HTML")


# ── Запуск ─────────────────────────────────────────────────────────────

def main():
    _ensure_data_dir()

    app = Application.builder().token(BOT_TOKEN).build()

    # Основной диалог
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            REG_FIO: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, reg_fio),
            ],
            REG_MODULE: [
                CallbackQueryHandler(reg_module, pattern=r"^module:"),
            ],
            MAIN_MENU: [
                CallbackQueryHandler(menu_handler, pattern=r"^(report_error|suggest)$"),
                CallbackQueryHandler(back_to_menu, pattern=r"^back_menu$"),
            ],
            ERROR_CATEGORY: [
                CallbackQueryHandler(error_category_handler, pattern=r"^errcat:"),
                CallbackQueryHandler(back_to_menu, pattern=r"^back_menu$"),
            ],
            ERROR_DESCRIPTION: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, error_description_handler),
            ],
            SUGGESTION_TEXT: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, suggestion_text_handler),
            ],
        },
        fallbacks=[CommandHandler("start", cmd_start)],
    )

    app.add_handler(conv)

    # Админка (вне ConversationHandler, чтобы работала всегда)
    app.add_handler(CommandHandler("admin", cmd_admin))
    app.add_handler(CallbackQueryHandler(admin_callback, pattern=r"^admin:"))

    print("🤖 CRM-Помощник запущен...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
