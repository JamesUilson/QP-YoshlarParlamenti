# Ro'yxatdan o'tish Telegram Boti

## O'rnatish

### 1. Token olish
1. Telegramda @BotFather ga yozing
2. /newbot buyrug'ini yuboring
3. Bot nomini va username'ini kiriting
4. Token'ni nusxalang (masalan: 1234567890:AAF...)

### 2. bot.py'ni sozlang
Faylni oching va quyidagilarni o'zgartiring:

```python
BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"   # <-- shu yerga tokeningizni qo'ying
ADMIN_IDS = [123456789]              # <-- o'zingizning Telegram ID'ingiz
```

**Telegram ID'ingizni bilish uchun:** @userinfobot ga /start yuboring

### 3. Kutubxonalarni o'rnatish

```bash
pip install -r requirements.txt
```

### 4. Botni ishga tushirish

```bash
python bot.py
```

---

## Foydalanuvchi uchun buyruqlar

| Buyruq     | Tavsif                    |
|------------|---------------------------|
| `/start`   | Ro'yxatdan o'tishni boshlash |
| `/cancel`  | Bekor qilish              |
| `/help`    | Yordam                    |

## Admin uchun buyruqlar

| Buyruq     | Tavsif                    |
|------------|---------------------------|
| `/stats`   | Statistikani ko'rish      |
| `/excel`   | Excel faylni yuklab olish |

---

## Excel formatı

Yuklab olinadigan Excel ikki sheet'dan iborat:

**Sheet 1 — Foydalanuvchilar:**
| ID | FullName | District | Age | Phone | Role | InitialRating |
|----|----------|----------|-----|-------|------|---------------|

**Sheet 2 — Ko'rsatmalar:** maydonlar tavsifi

---

## Fayl tuzilishi

```
📁 loyiha/
├── bot.py           — asosiy bot kodi
├── requirements.txt — kutubxonalar
├── users.json       — ma'lumotlar bazasi (avtomatik yaratiladi)
└── README.md        — ushbu fayl
```

## Serverda ishlatish (ixtiyoriy)

VPS yoki serverda doim ishlashi uchun `systemd` yoki `screen` ishlatish mumkin:

```bash
# screen bilan
screen -S mybot
python bot.py
# Ctrl+A, D — screen'dan chiqish (bot ishlayveradi)
```