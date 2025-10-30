import telebot
from openpyxl import load_workbook
from datetime import datetime, timedelta, time
import threading
import os
from dotenv import load_dotenv
import pytz
import time as t

# ===========================
# LOAD TOKEN AND SETUP BOT
# ===========================
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")  # stored in .env file
bot = telebot.TeleBot(BOT_TOKEN)

# Telegram group and topic info
# CHAT_ID = -1002257320998
# TOPIC_ID = 751
CHAT_ID = -1002635519712
TOPIC_ID = 434

# Singapore timezone
SGT = pytz.timezone("Asia/Singapore")

# ===========================
# HELPER FUNCTIONS
# ===========================
def is_weekend(date_obj):
    """Return True if the given date is Saturday (5) or Sunday (6)."""
    return date_obj.weekday() in [5, 6]


# ===========================
# FUNCTION TO READ SCHEDULE
# ===========================
def get_schedule(for_date, file_path="schedule.xlsx"):
    """
    Reads schedule.xlsx and returns schedule for the given date (datetime.date)
    """
    try:
        wb = load_workbook(file_path)
        sheet = wb.active

        schedule_data = {"Morning": [], "Afternoon": [], "Night": []}
        date_str, day_str = None, None

        for row in sheet.iter_rows(min_row=2, values_only=True):
            date_val, day, level, shift, student_id, name, building = row

            # Convert Excel date to datetime if not already
            if not isinstance(date_val, datetime):
                try:
                    date_val = datetime.strptime(str(date_val), "%d/%m/%Y")
                except Exception:
                    continue

            if date_val.date() == for_date:
                date_str = date_val.strftime("%Y/%m/%d")
                day_str = str(day)
                if name and name.strip().upper() != "NIL":
                    shift_name = shift.strip().capitalize()
                    schedule_data[shift_name].append(f"{name.strip()} ({level.strip()})")

        if not any(schedule_data.values()):
            return None, None, None

        return date_str, day_str, schedule_data

    except Exception as e:
        print(f"❌ Error reading schedule.xlsx: {e}")
        return None, None, None


# ===========================
# FORMAT MESSAGE
# ===========================
def format_schedule_message(date_str, day_str, schedule_data, label):
    """
    Format the schedule message for Telegram display.
    """
    message = f"📅 *Schedule for {label}*\n"
    message += f"Date: {date_str}\n"
    message += f"Day: {day_str}\n\n"

    for shift in ["Morning", "Afternoon", "Night"]:
        names = schedule_data.get(shift, [])
        message += f"*{shift} shift*\n"
        if names:
            message += "\n".join(names) + "\n\n"
        else:
            message += "_No one scheduled_\n\n"
    return message


# ===========================
# FUNCTION TO SEND MESSAGE
# ===========================
def send_schedule_message(target_date, label):
    """
    Sends the schedule message for the specified date to Telegram topic.
    Skips if the date is weekend.
    """
    if is_weekend(target_date):
        print(f"⏭️ Skipping {label} ({target_date.strftime('%A')}) — weekend.")
        return

    date_str, day_str, schedule_data = get_schedule(target_date)

    if not date_str:
        print(f"⚠️ No schedule found for {label}.")
        return

    msg = format_schedule_message(date_str, day_str, schedule_data, label)

    bot.send_message(
        chat_id=CHAT_ID,
        text=msg,
        parse_mode="Markdown",
        message_thread_id=TOPIC_ID
    )

    print(f"✅ {label.capitalize()} schedule message sent to topic successfully.")
    print(msg)


# ===========================
# DAILY SCHEDULER (17:50 SGT)
# ===========================
def schedule_daily_notification():
    """Schedules tomorrow’s schedule message every day at 17:50 SGT."""
    while True:
        now = datetime.now(SGT)
        target = datetime.combine(now.date(), time(17, 50, 0, tzinfo=SGT))

        # If it's already past 17:50 today → schedule for tomorrow
        if now > target:
            target += timedelta(days=1)

        wait_seconds = (target - now).total_seconds()
        print(f"⏰ Next auto notification at: {target.strftime('%Y-%m-%d %H:%M:%S')} (SGT)")
        t.sleep(wait_seconds)

        # Send for tomorrow (if not weekend)
        tomorrow = (datetime.now(SGT) + timedelta(days=1)).date()
        send_schedule_message(tomorrow, "tomorrow")


# ===========================
# TELEGRAM COMMANDS
# ===========================
@bot.message_handler(commands=['today_schedule'])
def cmd_today_schedule(message):
    today = datetime.now(SGT).date()
    send_schedule_message(today, "today")
    bot.reply_to(message, "✅ Checked today’s schedule (skips weekends automatically).")


@bot.message_handler(commands=['tomorrow_schedule'])
def cmd_tomorrow_schedule(message):
    tomorrow = (datetime.now(SGT) + timedelta(days=1)).date()
    send_schedule_message(tomorrow, "tomorrow")
    bot.reply_to(message, "✅ Checked tomorrow’s schedule (skips weekends automatically).")


# ===========================
# MAIN BOT START
# ===========================
if __name__ == "__main__":
    # Start background thread for auto daily notifications
    threading.Thread(target=schedule_daily_notification, daemon=True).start()

    print("Bot is running...")
    print("Auto notification set for 17:50 SGT daily (tomorrow’s schedule).")
    print("Weekends (Saturday & Sunday) are skipped automatically.")
    print("Commands available:")
    print("   /today_schedule — Send today’s schedule (skip weekends)")
    print("   /tomorrow_schedule — Send tomorrow’s schedule (skip weekends)")

    bot.infinity_polling()
