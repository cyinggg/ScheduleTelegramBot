import telebot
import os
import threading
import time
from dotenv import load_dotenv
from openpyxl import load_workbook
from datetime import datetime, timedelta
import pytz

# ======================
# Load environment
# ======================
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise Exception("BOT_TOKEN not found. Please check .env file!")

bot = telebot.TeleBot(BOT_TOKEN)

# Excel files
STUDENTS_FILE = "studentCoach.xlsx"
STAFF_FILE = "staff.xlsx"
LOG_FILE = "log.xlsx"
SCHEDULE_FILE = "schedule.xlsx"

# Telegram group/channel IDs
SUPERVISOR_CHAT_ID = -1001234567890     # Replace with your supervisor group ID
PROJECTHUB_CHAT_ID = -1002257320998     # ProjectHub Team group for daily notifications

# Singapore timezone
SGT = pytz.timezone("Asia/Singapore")

# ======================
# Helper Functions
# ======================

def get_student_name(student_id):
    wb = load_workbook(STUDENTS_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        cell_id = str(row[0]).split(".")[0]
        if str(student_id).strip() == cell_id:
            return row[1]
    return None

def is_valid_staff(staff_name):
    wb = load_workbook(STAFF_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]).strip().lower() == staff_name.lower():
            return True
    return False

def log_action(student_id, action):
    name = get_student_name(student_id)
    if not name:
        return False

    wb = load_workbook(LOG_FILE)
    ws = wb.active
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([timestamp, student_id, name, action, "Pending", "", ""])
    wb.save(LOG_FILE)

    bot.send_message(
        SUPERVISOR_CHAT_ID,
        f"📌 *{action} request pending verification*\n👤 Student: {name} ({student_id})\n🕒 Time: {timestamp}",
        parse_mode="Markdown"
    )
    return True

def get_pending_records():
    wb = load_workbook(LOG_FILE)
    ws = wb.active
    records = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if str(row[4]).strip().lower() == "pending":
            records.append((idx, row[2], row[3], row[0]))
    return records

def update_verification(rows_to_verify, staff_name, signature):
    wb = load_workbook(LOG_FILE)
    ws = wb.active
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verified_students = []

    for row_num in rows_to_verify:
        ws.cell(row=row_num, column=5, value="Verified")
        ws.cell(row=row_num, column=6, value=signature)
        ws.cell(row=row_num, column=7, value=timestamp)
        verified_students.append(ws.cell(row=row_num, column=3).value)

    wb.save(LOG_FILE)

    student_list = "\n".join([f"- {name}" for name in verified_students])
    bot.send_message(
        SUPERVISOR_CHAT_ID,
        f"✅ *Verification Completed*\n🗓 Date: {timestamp}\n👨‍🏫 Verified by: {staff_name}\n📋 Students Verified:\n{student_list}",
        parse_mode="Markdown"
    )

# ======================
# Daily Schedule Notification
# ======================

def send_daily_schedule():
    try:
        tomorrow = (datetime.now(SGT) + timedelta(days=1)).strftime("%Y-%m-%d")
        wb = load_workbook(SCHEDULE_FILE)
        ws = wb.active

        message = f"📅 *Next Day Schedule ({tomorrow})*\n\n"
        found = False

        for row in ws.iter_rows(min_row=2, values_only=True):
            row_date = None
            if isinstance(row[0], datetime):
                row_date = row[0].date()
            elif isinstance(row[0], str):
                try:
                    row_date = datetime.strptime(row[0], "%Y-%m-%d").date()
                except:
                    continue
            if row_date and row_date == (datetime.now(SGT).date() + timedelta(days=1)):
                found = True
                student = row[1] if row[1] else "-"
                shift = row[2] if len(row) > 2 and row[2] else "-"
                remarks = row[3] if len(row) > 3 and row[3] else ""
                message += f"👤 {student}\n🕒 Shift: {shift}\n💬 {remarks}\n\n"

        if not found:
            message += "No scheduled entries found."

        bot.send_message(PROJECTHUB_CHAT_ID, message, parse_mode="Markdown")
        print(f"[{datetime.now()}] ✅ Daily schedule sent to ProjectHub Team")

    except Exception as e:
        print(f"❌ Error sending schedule: {e}")

def schedule_daily_notification():
    while True:
        now_sgt = datetime.now(SGT)
        if now_sgt.hour == 17 and now_sgt.minute == 50:
            try:
                send_daily_schedule()
                time.sleep(60)
            except Exception as e:
                print(f"❌ Error sending schedule: {e}")
        time.sleep(20)

# ======================
# Telegram Commands
# ======================

@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(message, "👋 Welcome! Use /clockin or /clockout to record your shift.\nStaff can use /verify to confirm student records.")

@bot.message_handler(commands=['clockin'])
def clock_in(message):
    msg = bot.reply_to(message, "Please enter your Student ID to Clock In:")
    bot.register_next_step_handler(msg, process_clock_in)

def process_clock_in(message):
    student_id = message.text.strip()
    if log_action(student_id, "Clock In"):
        bot.reply_to(message, f"✅ Clock In recorded for {student_id}")
    else:
        bot.reply_to(message, "❌ Student ID not found. Please try again.")

@bot.message_handler(commands=['clockout'])
def clock_out(message):
    msg = bot.reply_to(message, "Please enter your Student ID to Clock Out:")
    bot.register_next_step_handler(msg, process_clock_out)

def process_clock_out(message):
    student_id = message.text.strip()
    if log_action(student_id, "Clock Out"):
        bot.reply_to(message, f"✅ Clock Out recorded for {student_id}")
    else:
        bot.reply_to(message, "❌ Student ID not found. Please try again.")

@bot.message_handler(commands=['verify'])
def verify_records(message):
    pending = get_pending_records()
    if not pending:
        bot.reply_to(message, "✅ No pending records for verification.")
        return

    msg_text = "📋 *Pending Clock In/Out Records:*\n"
    for idx, (row_num, name, action, time_) in enumerate(pending, start=1):
        msg_text += f"{idx}. {name} — {action} at {time_}\n"
    msg_text += "\nEnter the record numbers to verify (e.g. `1,3,5` or `all`):"

    bot.send_message(message.chat.id, msg_text, parse_mode="Markdown")
    bot.register_next_step_handler(message, lambda m: select_records_to_verify(m, pending))

def select_records_to_verify(message, pending):
    selection = message.text.strip().lower()
    if selection == "all":
        rows_to_verify = [r[0] for r in pending]
    else:
        try:
            indexes = [int(i) for i in selection.split(",")]
            rows_to_verify = [pending[i - 1][0] for i in indexes]
        except:
            bot.reply_to(message, "❌ Invalid selection. Please try again.")
            return

    msg = bot.reply_to(message, "Please enter your *staff name* for verification:", parse_mode="Markdown")
    bot.register_next_step_handler(msg, lambda m: verify_staff_identity(m, rows_to_verify))

def verify_staff_identity(message, rows_to_verify):
    staff_name = message.text.strip()
    if not is_valid_staff(staff_name):
        bot.reply_to(message, "❌ Staff name not found in records.")
        return

    msg = bot.reply_to(message, "✍️ Please enter your *signature* to confirm verification:", parse_mode="Markdown")
    bot.register_next_step_handler(msg, lambda m: finalize_verification(m, staff_name, rows_to_verify))

def finalize_verification(message, staff_name, rows_to_verify):
    signature = message.text.strip()
    update_verification(rows_to_verify, staff_name, signature)
    bot.reply_to(message, f"✅ Verified successfully by {staff_name}.\nSignature saved for selected records.")

@bot.message_handler(commands=['getid'])
def get_id(message):
    bot.reply_to(message, f"💬 Chat ID: `{message.chat.id}`", parse_mode="Markdown")

# ======================
# Run Bot + Scheduler
# ======================

print("🤖 WorkBot is running...")

# Run scheduler in background thread
threading.Thread(target=schedule_daily_notification, daemon=True).start()

# Start Telegram polling
bot.infinity_polling()
