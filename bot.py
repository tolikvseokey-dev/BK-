import os
from pathlib import Path

from dotenv import load_dotenv
import telebot

# =========================
# Загрузка переменных среды
# =========================
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME")

if not BOT_TOKEN:
    raise ValueError("Не найден BOT_TOKEN в .env")

bot = telebot.TeleBot(BOT_TOKEN)

# =========================
# Папка для временных файлов
# =========================
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)


# =========================
# Команда /start
# =========================
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(
        message.chat.id,
        "Бот запущен 🚀\n\n"
        "Отправь Excel-файл проверки питания в формате .xlsx"
    )


# =========================
# Обработка документов
# =========================
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        document = message.document

        if not document:
            bot.reply_to(message, "Не удалось получить файл.")
            return

        file_name = document.file_name or ""
        file_name_lower = file_name.lower()

        if not file_name_lower.endswith(".xlsx"):
            bot.reply_to(
                message,
                "Пожалуйста, отправь Excel-файл в формате .xlsx"
            )
            return

        bot.reply_to(message, "Файл получен, начинаю загрузку ⏳")

        file_info = bot.get_file(document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        safe_file_name = Path(file_name).name
        save_path = UPLOADS_DIR / safe_file_name

        with open(save_path, "wb") as new_file:
            new_file.write(downloaded_file)

        bot.reply_to(
            message,
            f"Файл получен ✅\n\n"
            f"Имя файла: {safe_file_name}"
        )

    except Exception as e:
        bot.reply_to(
            message,
            f"Ошибка при обработке файла: {e}"
        )


# =========================
# Запуск бота
# =========================
print("Бот запущен...")
bot.infinity_polling()
