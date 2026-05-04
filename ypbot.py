import os
import json
import logging
from datetime import datetime
from io import BytesIO

from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, CallbackQueryHandler, filters, ContextTypes
)
from telegram.request import HTTPXRequest
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ─── SOZLAMALAR ───────────────────────────────────────────────
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"   # @BotFather dan olingan token
ADMIN_IDS = [123456789]              # Admin Telegram ID lari (bir nechta bo'lsa: [111, 222])
DB_FILE   = "users.json"            # Ma'lumotlar saqlanadigan fayl
# ──────────────────────────────────────────────────────────────

DISTRICTS = {
    "01": "Toshkent sh.", "02": "Toshkent v.",  "03": "Andijon",
    "04": "Farg'ona",     "05": "Namangan",      "06": "Samarqand",
    "07": "Buxoro",       "08": "Navoiy",        "09": "Qashqadaryo",
    "10": "Surxondaryo",  "11": "Jizzax",        "12": "Sirdaryo",
    "13": "Xorazm",       "14": "Qoraqalpog'iston",
}

# ConversationHandler holatlari
NAME, DISTRICT, AGE, PHONE, POSITION = range(5)


# ─── DATABASE ─────────────────────────────────────────────────
def load_db() -> list:
    if not os.path.exists(DB_FILE):
        return []
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(data: list):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def next_id() -> str:
    return str(len(load_db()) + 1).zfill(4)

def add_user(record: dict):
    data = load_db()
    data.append(record)
    save_db(data)

def already_registered(telegram_id: int) -> bool:
    return any(u.get("TelegramID") == telegram_id for u in load_db())


# ─── EXCEL YARATISH ───────────────────────────────────────────
def build_excel() -> BytesIO:
    users = load_db()
    wb = openpyxl.Workbook()

    # ── Sheet 1: Foydalanuvchilar ──
    ws = wb.active
    ws.title = "Foydalanuvchilar"

    headers = ["ID", "FullName", "District", "Age", "Phone", "Role", "InitialRating", "Lavozim"]
    col_widths = [7, 24, 10, 6, 17, 10, 14, 22]

    header_fill = PatternFill("solid", fgColor="1D9E75")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")
    thin   = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center
        cell.border    = thin
        ws.column_dimensions[cell.column_letter].width = w

    ws.row_dimensions[1].height = 22

    even_fill = PatternFill("solid", fgColor="E8F7F2")
    for row_i, u in enumerate(users, 2):
        row_data = [
            u.get("ID"), u.get("FullName"), u.get("District"),
            u.get("Age"), u.get("Phone", ""), u.get("Role"),
            u.get("InitialRating"), u.get("Lavozim", "")
        ]
        fill = even_fill if row_i % 2 == 0 else None
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_i, column=col, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.alignment = center if col in (1, 3, 4, 6, 7) else left
            cell.border    = thin
            if fill:
                cell.fill = fill
        ws.row_dimensions[row_i].height = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(len(users)+1, 2)}"

    # ── Sheet 2: Ko'rsatmalar ──
    ws2 = wb.create_sheet("Ko'rsatmalar")
    instructions = [
        ("Maydon",        "Tavsif",                        "Namuna"),
        ("ID",            "4 ta raqam (masalan: 0001)",    "0001"),
        ("FullName",      "To'liq ismi",                   "Ali Valiyev"),
        ("District",      "Hudud raqami (01-14)",          "01"),
        ("Age",           "Yosh (18-35)",                  "22"),
        ("Phone",         "Telefon raqami (ixtiyoriy)",    "+998901234567"),
        ("Role",          "user/admin/debugger",           "user"),
        ("InitialRating", "Boshlang'ich reyting (0-20)",   "10"),
        ("Lavozim",       "Foydalanuvchi lavozimi",        "Matbuot kotibi"),
    ]
    for r, row in enumerate(instructions, 1):
        for c, val in enumerate(row, 1):
            cell = ws2.cell(row=r, column=c, value=val)
            cell.font   = Font(bold=(r==1), name="Arial", size=10,
                               color="FFFFFF" if r==1 else "000000")
            cell.fill   = PatternFill("solid", fgColor="1D9E75") if r==1 else PatternFill("solid", fgColor="FFFFFF")
            cell.border = thin
            cell.alignment = left
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 32
    ws2.column_dimensions["C"].width = 20

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── HUDUD KLAVIATURASI ───────────────────────────────────────
def district_keyboard() -> InlineKeyboardMarkup:
    buttons = []
    items = list(DISTRICTS.items())
    for i in range(0, len(items), 2):
        row = []
        for code, name in items[i:i+2]:
            row.append(InlineKeyboardButton(f"{code} · {name}", callback_data=f"dist_{code}"))
        buttons.append(row)
    return InlineKeyboardMarkup(buttons)


# ─── HANDLERLAR ───────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    if already_registered(user.id):
        await update.message.reply_text(
            "✅ Siz allaqachon ro'yxatdan o'tgansiz!\n\n"
            "Yordam uchun /help buyrug'ini yuboring."
        )
        return ConversationHandler.END

    await update.message.reply_text(
        f"Salom, {user.first_name}! 👋\n\n"
        "Ro'yxatdan o'tish uchun bir nechta savolga javob bering.\n\n"
        "❌ Bekor qilish uchun /cancel buyrug'ini yuboring.\n\n"
        "━━━━━━━━━━━━━━━━━\n"
        "📝 *Ismingiz va familiyangizni kiriting:*",
        parse_mode="Markdown"
    )
    return NAME


async def get_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = update.message.text.strip()
    if len(name) < 3:
        await update.message.reply_text(
            "⚠️ Iltimos, to'liq ism familiyangizni kiriting (kamida 3 belgi)."
        )
        return NAME

    context.user_data["name"] = name
    await update.message.reply_text(
        f"✅ *{name}*\n\n🗺 *Yashash hududingizni tanlang:*",
        parse_mode="Markdown",
        reply_markup=district_keyboard()
    )
    return DISTRICT


async def get_district(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    code = query.data.replace("dist_", "")
    name = DISTRICTS.get(code, code)
    context.user_data["district"] = code

    await query.edit_message_text(
        f"✅ *{code} · {name}*\n\n"
        "━━━━━━━━━━━━━━━━━\n"
        "🎂 *Yoshingizni kiriting* (18–35 oralig'ida):",
        parse_mode="Markdown"
    )
    return AGE


async def get_age(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        age = int(update.message.text.strip())
        if not (18 <= age <= 35):
            raise ValueError
    except ValueError:
        await update.message.reply_text("⚠️ Yosh 18 dan 35 gacha bo'lishi kerak. Qaytadan kiriting:")
        return AGE

    context.user_data["age"] = age
    await update.message.reply_text(
        f"✅ *{age} yosh*\n\n"
        "━━━━━━━━━━━━━━━━━\n"
        "📱 *Telefon raqamingiz?*\n"
        "_Ixtiyoriy — o'tkazish uchun «–» yozing_",
        parse_mode="Markdown"
    )
    return PHONE


async def get_phone(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    phone = "" if text in ("–", "-", "yo'q", "skip") else text
    context.user_data["phone"] = phone

    await update.message.reply_text(
        f"✅ *{phone or '—'}*\n\n"
        "━━━━━━━━━━━━━━━━━\n"
        "💼 *Lavozimingizni kiriting:*\n"
        "_Masalan: Matbuot kotibi, Direktor, O'qituvchi..._",
        parse_mode="Markdown"
    )
    return POSITION


async def get_position(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    position = update.message.text.strip()
    if len(position) < 2:
        await update.message.reply_text("⚠️ Lavozimni to'liq kiriting (kamida 2 belgi).")
        return POSITION

    context.user_data["position"] = position

    uid = next_id()
    dist_code = context.user_data["district"]
    dist_name = DISTRICTS.get(dist_code, dist_code)
    phone     = context.user_data["phone"]

    record = {
        "ID":            uid,
        "FullName":      context.user_data["name"],
        "District":      dist_code,
        "Age":           context.user_data["age"],
        "Phone":         phone,
        "Role":          "user",
        "InitialRating": 10,
        "Lavozim":       position,
        "TelegramID":    update.effective_user.id,
        "RegisteredAt":  datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    add_user(record)

    await update.message.reply_text(
        f"🎉 *Ro'yxatdan o'tish muvaffaqiyatli yakunlandi!*\n\n"
        f"🪪 ID: `{uid}`\n"
        f"👤 Ism: {record['FullName']}\n"
        f"🗺 Hudud: {dist_code} · {dist_name}\n"
        f"🎂 Yosh: {record['Age']}\n"
        f"📱 Telefon: {phone or '—'}\n"
        f"⭐ Reyting: {record['InitialRating']}\n"
        f"💼 Lavozim: {position}\n\n"
        "Sizning ma'lumotlaringiz saqlandi. Rahmat!",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    logger.info(f"Yangi foydalanuvchi: {record['FullName']} (ID: {uid}), Lavozim: {position}")
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Ro'yxatdan o'tish bekor qilindi.\n"
        "Qaytadan boshlash uchun /start buyrug'ini yuboring.",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


# ─── ADMIN HANDLERLAR ─────────────────────────────────────────
def is_admin(update: Update) -> bool:
    return update.effective_user.id in ADMIN_IDS


async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Ruxsat yo'q.")
        return

    users = load_db()
    total    = len(users)
    w_phone  = sum(1 for u in users if u.get("Phone"))
    by_dist  = {}
    for u in users:
        d = u.get("District", "?")
        by_dist[d] = by_dist.get(d, 0) + 1

    top = sorted(by_dist.items(), key=lambda x: -x[1])[:5]
    top_txt = "\n".join(f"  {DISTRICTS.get(c,c)}: {n} ta" for c, n in top) or "  —"

    await update.message.reply_text(
        f"📊 *Statistika*\n\n"
        f"👥 Jami: *{total}* ta\n"
        f"📱 Telefonli: *{w_phone}* ta\n\n"
        f"🗺 Top hududlar:\n{top_txt}",
        parse_mode="Markdown"
    )


async def admin_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Ruxsat yo'q.")
        return

    users = load_db()
    if not users:
        await update.message.reply_text("ℹ️ Hali ro'yxatdan o'tgan foydalanuvchi yo'q.")
        return

    msg = await update.message.reply_text("⏳ Excel tayyorlanmoqda...")
    buf = build_excel()
    filename = f"foydalanuvchilar_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    await update.message.reply_document(
        document=buf,
        filename=filename,
        caption=f"📋 Jami: {len(users)} ta foydalanuvchi\n📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )
    await msg.delete()


async def admin_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Ruxsat yo'q.")
        return
    await update.message.reply_text(
        "🔧 *Admin buyruqlar:*\n\n"
        "/stats — statistika ko'rish\n"
        "/excel — Excel faylni yuklab olish\n"
        "/help — bu yordam xabari",
        parse_mode="Markdown"
    )


async def user_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ℹ️ *Yordam*\n\n"
        "/start — Ro'yxatdan o'tish\n"
        "/cancel — Bekor qilish",
        parse_mode="Markdown"
    )


# ─── MAIN ─────────────────────────────────────────────────────
def main():
    # ── Proxy sozlamasi (agar Telegram blok bo'lsa) ──────────────
    # Proxy ishlatmoqchi bo'lsangiz, quyidagi qatorni yoching va
    # proxy manzilini kiriting (masalan: "socks5://127.0.0.1:1080"):
    PROXY_URL = None  # masalan: "http://127.0.0.1:8080"
    # ─────────────────────────────────────────────────────────────

    request = HTTPXRequest(
        connect_timeout=30.0,
        read_timeout=30.0,
        write_timeout=30.0,
        pool_timeout=30.0,
        proxy=PROXY_URL,
    )

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .request(request)
        .build()
    )

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            DISTRICT: [CallbackQueryHandler(get_district, pattern=r"^dist_")],
            AGE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_age)],
            PHONE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            POSITION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_position)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("stats", admin_stats))
    app.add_handler(CommandHandler("excel", admin_excel))
    app.add_handler(CommandHandler("help",  user_help))

    logger.info("Bot ishga tushdi...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
