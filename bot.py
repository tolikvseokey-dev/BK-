import os
import re
from pathlib import Path
from typing import Optional, Tuple, Dict, List

import pandas as pd
import telebot
from dotenv import load_dotenv

# =========================
# Загрузка переменных среды
# =========================
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "").strip()

if not BOT_TOKEN:
    raise ValueError("Не найден BOT_TOKEN в .env")

bot = telebot.TeleBot(BOT_TOKEN)

# =========================
# Пути
# =========================
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)

CRITERIA_FILE = BASE_DIR / "criteria.xlsx"

# =========================
# Настройки
# =========================
MAX_VIOLATIONS_PER_MESSAGE = 15


# =========================
# Вспомогательные функции
# =========================
def is_admin(message) -> bool:
    """
    Если ADMIN_USERNAME задан, работать с ботом может только этот username.
    Если не задан — доступ открыт всем.
    """
    if not ADMIN_USERNAME:
        return True

    username = (message.from_user.username or "").strip()
    return username.lower() == ADMIN_USERNAME.lower()


def normalize_text(value) -> str:
    """
    Нормализация текста для сравнения:
    - lower
    - ё -> е
    - убрать лишние пробелы
    - убрать служебные знаки
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    text = str(value).strip().lower().replace("ё", "е")
    text = text.replace("\n", " ").replace("\t", " ")
    text = re.sub(r"[\"'`]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_store_name(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    if text in {"итого", "всего"}:
        return ""
    return str(value).strip()


def normalize_unit(unit: str) -> str:
    unit = normalize_text(unit)

    replacements = {
        "гр": "г",
        "грамм": "г",
        "грамма": "г",
        "граммов": "г",
        "кг": "кг",
        "мл": "мл",
        "л": "л",
        "шт.": "шт",
        "штука": "шт",
        "штуки": "шт",
        "штук": "шт",
        "порция": "порц",
        "порции": "порц",
        "порц.": "порц",
    }

    if unit in replacements:
        return replacements[unit]

    return unit


def extract_number_and_unit(value) -> Tuple[Optional[float], str]:
    """
    Примеры:
    '150 г' -> (150, 'г')
    '150гр' -> (150, 'г')
    '2 шт' -> (2, 'шт')
    '1 порц' -> (1, 'порц')
    200 -> (200, '')
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, ""

    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value), ""

    text = str(value).strip().lower().replace(",", ".")
    text = text.replace("ё", "е")

    number_match = re.search(r"(\d+(?:\.\d+)?)", text)
    if not number_match:
        return None, ""

    number = float(number_match.group(1))
    unit = text[number_match.end():].strip()

    # Убираем лишние символы
    unit = re.sub(r"[^a-zа-я.%]+", " ", unit, flags=re.IGNORECASE).strip()
    unit = normalize_unit(unit)

    # Дополнительная нормализация вариантов без пробелов
    if not unit:
        if "гр" in text or re.search(r"\d+\s*г\b", text):
            unit = "г"
        elif "шт" in text:
            unit = "шт"
        elif "порц" in text or "порц" in text:
            unit = "порц"
        elif "мл" in text:
            unit = "мл"

    return number, unit


def format_number(value: float) -> str:
    if value is None:
        return "—"
    if float(value).is_integer():
        return str(int(value))
    return str(round(value, 2)).replace(".", ",")


def format_date_for_message(value) -> str:
    if pd.isna(value):
        return "—"

    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return str(value)

    return dt.strftime("%d.%m %H:%M")


def format_period(series: pd.Series) -> str:
    dates = pd.to_datetime(series, errors="coerce").dropna()
    if dates.empty:
        return "Период не определён"

    date_min = dates.min()
    date_max = dates.max()

    if date_min.date() == date_max.date():
        return date_min.strftime("%d.%m.%Y")

    return f"{date_min.strftime('%d.%m.%Y')} — {date_max.strftime('%d.%m.%Y')}"


def get_safe_file_name(file_name: str) -> str:
    if not file_name:
        return "uploaded_report.xlsx"
    return Path(file_name).name


def find_header_row(raw_df: pd.DataFrame) -> Optional[int]:
    """
    Ищем строку заголовков в отчёте проверки питания.
    Ищем по наличию нескольких ключевых признаков.
    """
    for idx in range(min(len(raw_df), 30)):
        row_values = [normalize_text(v) for v in raw_df.iloc[idx].tolist()]
        row_text = " | ".join(row_values)

        has_store = "торговое предприятие" in row_text or "лавка" in row_text
        has_date = "дата" in row_text or "время" in row_text
        has_check = "номер чека" in row_text or "чек" in row_text
        has_employee = "сотрудник" in row_text
        has_product = "блюдо" in row_text or "пози" in row_text or "продукт" in row_text

        score = sum([has_store, has_date, has_check, has_employee, has_product])

        if score >= 3:
            return idx

    return None


def load_report_dataframe(file_path: Path) -> pd.DataFrame:
    """
    Загружаем отчёт проверки питания.
    Ищем строку заголовков автоматически.
    Потом забираем нужные 7 колонок:
    A, B(не используем), C, D, E, F, G
    """
    excel = pd.ExcelFile(file_path, engine="openpyxl")
    if not excel.sheet_names:
        raise ValueError("В файле нет листов.")

    sheet_name = excel.sheet_names[0]

    raw_df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    header_row = find_header_row(raw_df)
    if header_row is None:
        raise ValueError("Не удалось найти строку заголовков в файле проверки питания.")

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl"
    )

    if len(df.columns) < 7:
        raise ValueError("В файле меньше 7 столбцов. Проверь структуру отчёта.")

    # Берём нужные позиции:
    # A - лавка
    # B - не используем
    # C - дата/время
    # D - номер чека
    # E - сотрудник
    # F - продукт
    # G - факт
    selected = df.iloc[:, 0:7].copy()

    selected.columns = [
        "store",
        "unused",
        "check_datetime",
        "check_number",
        "employee_name",
        "product_name",
        "actual_raw",
    ]

    # Удаляем полностью пустые строки
    selected = selected.dropna(how="all").copy()

    # Протягиваем лавку вниз
    selected["store"] = selected["store"].apply(normalize_store_name)
    selected["store"] = selected["store"].replace("", pd.NA).ffill()

    # Убираем строки без ключевых значений
    selected["product_name"] = selected["product_name"].astype(str).str.strip()
    selected["employee_name"] = selected["employee_name"].astype(str).str.strip()
    selected["check_number"] = selected["check_number"].astype(str).str.strip()

    # Преобразуем "nan" от astype(str) обратно в пусто
    for col in ["product_name", "employee_name", "check_number"]:
        selected[col] = selected[col].replace("nan", "").replace("None", "")

    selected = selected[
        (selected["store"].notna()) &
        (selected["product_name"] != "") &
        (selected["actual_raw"].notna())
    ].copy()

    return selected


def load_criteria_dataframe(criteria_path: Path) -> pd.DataFrame:
    """
    Загружаем критерии.
    Ожидаем структуру, как на твоём скрине:
    A — исходник
    B — продукт
    C — норма текстом
    D — число нормы
    """
    if not criteria_path.exists():
        raise ValueError("Файл criteria.xlsx не найден в корне проекта.")

    excel = pd.ExcelFile(criteria_path, engine="openpyxl")
    if not excel.sheet_names:
        raise ValueError("В criteria.xlsx нет листов.")

    sheet_name = excel.sheet_names[0]

    df = pd.read_excel(
        criteria_path,
        sheet_name=sheet_name,
        engine="openpyxl"
    )

    if len(df.columns) < 4:
        raise ValueError("В criteria.xlsx должно быть минимум 4 столбца.")

    # Жёстко берём B, C, D по позиции, как ты подготовил
    criteria = pd.DataFrame({
        "criteria_product": df.iloc[:, 1],
        "criteria_norm_text": df.iloc[:, 2],
        "criteria_norm_value": df.iloc[:, 3],
    })

    criteria = criteria.dropna(how="all").copy()

    criteria["criteria_product"] = criteria["criteria_product"].astype(str).str.strip()
    criteria["criteria_norm_text"] = criteria["criteria_norm_text"].astype(str).str.strip()

    criteria["criteria_product"] = criteria["criteria_product"].replace("nan", "")
    criteria["criteria_norm_text"] = criteria["criteria_norm_text"].replace("nan", "")

    criteria["criteria_norm_value"] = pd.to_numeric(
        criteria["criteria_norm_value"],
        errors="coerce"
    )

    criteria = criteria[
        (criteria["criteria_product"] != "") &
        (criteria["criteria_norm_value"].notna())
    ].copy()

    # Разбираем единицы из текстовой нормы
    units = criteria["criteria_norm_text"].apply(extract_number_and_unit)
    criteria["criteria_unit"] = units.apply(lambda x: x[1])

    criteria["product_key"] = criteria["criteria_product"].apply(normalize_text)

    return criteria


def build_criteria_lookup(criteria_df: pd.DataFrame) -> Dict[str, dict]:
    lookup = {}
    for _, row in criteria_df.iterrows():
        key = row["product_key"]
        if not key:
            continue

        lookup[key] = {
            "product_name": str(row["criteria_product"]).strip(),
            "norm_text": str(row["criteria_norm_text"]).strip(),
            "norm_value": float(row["criteria_norm_value"]),
            "unit": normalize_unit(str(row["criteria_unit"]).strip()),
        }
    return lookup


def find_criteria_for_product(product_name: str, lookup: Dict[str, dict]) -> Optional[dict]:
    """
    1) точное совпадение
    2) мягкое совпадение contains
    """
    key = normalize_text(product_name)
    if not key:
        return None

    if key in lookup:
        return lookup[key]

    # Мягкий поиск
    candidates = []
    for criteria_key, criteria_data in lookup.items():
        if key in criteria_key or criteria_key in key:
            candidates.append((len(criteria_key), criteria_data))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]

    return None


def analyze_report(report_df: pd.DataFrame, criteria_lookup: Dict[str, dict]) -> dict:
    violations: List[dict] = []
    not_found_count = 0
    unit_mismatch_count = 0
    checked_rows_count = 0

    store_stats: Dict[str, int] = {}

    for _, row in report_df.iterrows():
        store = str(row["store"]).strip()
        check_datetime = row["check_datetime"]
        check_number = str(row["check_number"]).strip()
        employee_name = str(row["employee_name"]).strip()
        product_name = str(row["product_name"]).strip()
        actual_raw = row["actual_raw"]

        actual_value, actual_unit = extract_number_and_unit(actual_raw)
        if actual_value is None:
            continue

        criteria = find_criteria_for_product(product_name, criteria_lookup)
        if not criteria:
            not_found_count += 1
            continue

        norm_value = criteria["norm_value"]
        norm_text = criteria["norm_text"]
        criteria_unit = criteria["unit"]

        # Если обе единицы есть и они не совпадают — не сравниваем
        if actual_unit and criteria_unit and actual_unit != criteria_unit:
            unit_mismatch_count += 1
            continue

        checked_rows_count += 1

        if float(actual_value) > float(norm_value):
            exceed_value = float(actual_value) - float(norm_value)

            violation = {
                "store": store,
                "check_datetime": check_datetime,
                "check_number": check_number,
                "employee_name": employee_name,
                "product_name": product_name,
                "actual_value": actual_value,
                "actual_unit": actual_unit or criteria_unit,
                "norm_value": norm_value,
                "norm_text": norm_text,
                "exceed_value": exceed_value,
                "matched_product_name": criteria["product_name"],
            }
            violations.append(violation)
            store_stats[store] = store_stats.get(store, 0) + 1

    unique_checks = report_df["check_number"].astype(str).replace("nan", "").nunique()
    stores_count = report_df["store"].dropna().astype(str).nunique()
    period_text = format_period(report_df["check_datetime"])

    return {
        "period_text": period_text,
        "stores_count": stores_count,
        "unique_checks": unique_checks,
        "checked_rows_count": checked_rows_count,
        "violations": violations,
        "violations_count": len(violations),
        "store_stats": store_stats,
        "not_found_count": not_found_count,
        "unit_mismatch_count": unit_mismatch_count,
    }


def build_summary_message(result: dict) -> str:
    lines = [
        "🍽 Проверка питания",
        "",
        f"📅 {result['period_text']}",
        f"🏪 Лавок: {result['stores_count']}",
        f"🧾 Чеков: {result['unique_checks']}",
        f"📊 Позиций проверено: {result['checked_rows_count']}",
        "",
        f"⚠️ Нарушений: {result['violations_count']}",
        f"❓ Не найдено в критериях: {result['not_found_count']}",
    ]

    if result["unit_mismatch_count"] > 0:
        lines.append(f"⚖️ Несовпадение единиц: {result['unit_mismatch_count']}")

    return "\n".join(lines)


def build_store_stats_message(store_stats: Dict[str, int]) -> str:
    if not store_stats:
        return "🏪 По лавкам:\n\nНарушений не найдено ✅"

    sorted_items = sorted(store_stats.items(), key=lambda x: (-x[1], x[0]))

    lines = ["🏪 По лавкам:", ""]
    for store, count in sorted_items:
        lines.append(f"{store} — {count}")

    return "\n".join(lines)


def build_violation_text(item: dict) -> str:
    unit = item["actual_unit"] or ""

    actual_part = f"{format_number(item['actual_value'])} {unit}".strip()
    exceed_part = f"+{format_number(item['exceed_value'])} {unit}".strip()

    lines = [
        "⚠️ Нарушение",
        "",
        f"🏪 {item['store']}",
        f"👤 {item['employee_name'] or '—'}",
        f"🕒 {format_date_for_message(item['check_datetime'])}",
        f"🧾 Чек: {item['check_number'] or '—'}",
        "",
        f"🍲 {item['product_name']}",
        f"📦 Факт: {actual_part}",
        f"📏 Норма: {item['norm_text']}",
        "",
        f"❗ Превышение: {exceed_part}",
    ]
    return "\n".join(lines)


def chunk_list(items: List[dict], chunk_size: int) -> List[List[dict]]:
    return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]


def send_analysis_result(chat_id: int, result: dict):
    bot.send_message(chat_id, build_summary_message(result))
    bot.send_message(chat_id, build_store_stats_message(result["store_stats"]))

    violations = result["violations"]

    if not violations:
        bot.send_message(chat_id, "Нарушений не найдено ✅")
        return

    parts = chunk_list(violations, MAX_VIOLATIONS_PER_MESSAGE)

    for index, part in enumerate(parts, start=1):
        header = f"⚠️ Нарушения ({index}/{len(parts)})"
        messages = [header, ""]
        messages.extend([build_violation_text(item) for item in part])
        bot.send_message(chat_id, "\n\n".join(messages))


# =========================
# Команды
# =========================
@bot.message_handler(commands=["start"])
def start(message):
    if not is_admin(message):
        bot.send_message(message.chat.id, "У тебя нет доступа к этому боту.")
        return

    text = (
        "Бот запущен 🚀\n\n"
        "Отправь Excel-файл проверки питания в формате .xlsx\n"
        "Файл критериев должен лежать в репозитории под именем criteria.xlsx"
    )
    bot.send_message(message.chat.id, text)


@bot.message_handler(commands=["help"])
def help_command(message):
    if not is_admin(message):
        bot.send_message(message.chat.id, "У тебя нет доступа к этому боту.")
        return

    text = (
        "Как использовать бота:\n\n"
        "1. В репозитории должен лежать файл criteria.xlsx\n"
        "2. Отправь в бот файл проверки питания .xlsx\n"
        "3. Бот проверит продукты и покажет нарушения"
    )
    bot.send_message(message.chat.id, text)


# =========================
# Обработка документов
# =========================
@bot.message_handler(content_types=["document"])
def handle_document(message):
    if not is_admin(message):
        bot.reply_to(message, "У тебя нет доступа к этому боту.")
        return

    try:
        document = message.document
        if not document:
            bot.reply_to(message, "Не удалось получить файл.")
            return

        file_name = document.file_name or ""
        file_name_lower = file_name.lower()

        if not file_name_lower.endswith(".xlsx"):
            bot.reply_to(message, "Пожалуйста, отправь Excel-файл в формате .xlsx")
            return

        if not CRITERIA_FILE.exists():
            bot.reply_to(
                message,
                "Не найден файл criteria.xlsx в корне проекта."
            )
            return

        bot.reply_to(message, "Файл получен, начинаю проверку ⏳")

        file_info = bot.get_file(document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        safe_file_name = get_safe_file_name(file_name)
        save_path = UPLOADS_DIR / safe_file_name

        with open(save_path, "wb") as new_file:
            new_file.write(downloaded_file)

        criteria_df = load_criteria_dataframe(CRITERIA_FILE)
        criteria_lookup = build_criteria_lookup(criteria_df)

        report_df = load_report_dataframe(save_path)
        result = analyze_report(report_df, criteria_lookup)

        send_analysis_result(message.chat.id, result)

    except Exception as e:
        bot.reply_to(message, f"Ошибка при обработке файла: {e}")


# =========================
# Запуск
# =========================
print("Бот запущен...")
bot.infinity_polling(timeout=60, long_polling_timeout=60)
