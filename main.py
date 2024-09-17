import openpyxl
from datetime import datetime
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext
from flask import Flask, request
import logging
import os
from difflib import get_close_matches
from threading import Thread
import asyncio

TOKEN = os.getenv("YOUR_BOT_TOKEN")
WEBHOOK_URL = os.getenv("YOUR_WEBHOOK_URL")

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

app = Flask(__name__)
bot = None
application = None

# Global dictionary to store student selection by chat_id
student_selection = {}

def format_date_column(date: datetime) -> str:
    return date.strftime("%d%m")

async def update_attendance(roll_number: int, specific_date: str = None, attendance_status: str = 'P') -> str:
    file_path = "FAAtt.xlsx"
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    today_date = datetime.today() if not specific_date else datetime.strptime(specific_date, "%d-%b")
    formatted_date = format_date_column(today_date)

    # Check for existing date column
    date_column = None
    for col in range(5, sheet.max_column + 1):
        cell_value = sheet.cell(row=1, column=col).value
        if cell_value == formatted_date:
            date_column = col
            break

    if date_column is None:
        total_column = sheet.max_column
        if sheet.cell(row=1, column=total_column).value == 'Total':
            total_column -= 1
        sheet.insert_cols(total_column + 1)
        sheet.cell(row=1, column=total_column + 1, value=formatted_date)
        sheet.cell(row=1, column=total_column + 2, value='Total')
        date_column = total_column + 1

    student_row = None
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=2).value == roll_number:
            student_row = row
            break

    if student_row is None:
        return f"Error: Roll number {roll_number} not found."

    current_status = sheet.cell(row=student_row, column=date_column).value
    if current_status == attendance_status:
        return f"Already {attendance_status}"

    sheet.cell(row=student_row, column=date_column).value = attendance_status

    total_column = sheet.max_column
    count_range = f"E{student_row}:{get_column_letter(total_column - 1)}{student_row}"
    sheet.cell(row=student_row, column=total_column).value = f'=COUNTIF({count_range}, "P")'

    workbook.save(file_path)
    return f"Set {attendance_status} success"

def find_students_by_name(name: str):
    file_path = "FAAtt.xlsx"
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    student_names = [sheet.cell(row=row, column=1).value for row in range(2, sheet.max_row + 1)]
    matched_names = get_close_matches(name, student_names, n=5, cutoff=0.5)

    matched_students = []
    for row in range(2, sheet.max_row + 1):
        student_name = sheet.cell(row=row, column=1).value
        if student_name in matched_names:
            roll_number = sheet.cell(row=row, column=2).value
            matched_students.append((student_name, roll_number, row))

    return matched_students

async def handle_message(update: Update, context: CallbackContext) -> None:
    text = update.message.text.strip()

    if update.message.chat_id in student_selection:
        try:
            selection_index = int(text)
            student_list = student_selection[update.message.chat_id]
            if 0 <= selection_index < len(student_list):
                student_name, roll_number, _ = student_list[selection_index]
                res = await update_attendance(roll_number)
                await update.message.reply_text(res)
            else:
                await update.message.reply_text("Invalid selection. Try again.")
        except ValueError:
            await update.message.reply_text("Invalid selection. Please enter a number.")
        finally:
            del student_selection[update.message.chat_id]
        return

    try:
        parts = text.split()
        roll_number = None
        specific_date = None
        attendance_status = 'P'

        if len(parts) == 3:
            try:
                specific_date = datetime.strptime(parts[1], "%d-%b").strftime("%d-%b")
                roll_number = int(parts[0])
                if parts[2].lower() in ['a', 'absent']:
                    attendance_status = 'A'
            except ValueError:
                roll_number = None

        elif len(parts) == 2:
            try:
                specific_date = datetime.strptime(parts[1], "%d-%b").strftime("%d-%b")
                roll_number = int(parts[0])
            except ValueError:
                roll_number = int(parts[0]) if parts[0].isdigit() else None
                if parts[1].lower() in ['a', 'absent']:
                    attendance_status = 'A'

        elif len(parts) == 1:
            try:
                roll_number = int(parts[0])
            except ValueError:
                roll_number = None

        if roll_number is not None:
            res = await update_attendance(roll_number, specific_date, attendance_status)
            await update.message.reply_text(res)
        else:
            name = text if len(parts) == 1 else parts[0]
            matched_students = find_students_by_name(name)
            if matched_students:
                if len(matched_students) == 1:
                    student_name, roll_number, _ = matched_students[0]
                    res = await update_attendance(roll_number, specific_date, attendance_status)
                    await update.message.reply_text(res)
                else:
                    student_selection[update.message.chat_id] = matched_students
                    student_list = "\n".join([f"{idx}. {name} (Roll: {roll})" for idx, (name, roll, _) in enumerate(matched_students)])
                    await update.message.reply_text(f"Multiple students found:\n{student_list}\nPlease select a student by number.")
            else:
                await update.message.reply_text(f"No students found matching '{name}'.")
    except ValueError:
        await update.message.reply_text("Invalid input. Please enter a valid roll number or use the /sendfile command.")

async def error(update: Update, context: CallbackContext) -> None:
    logging.error(f'Update {update} caused error {context.error}')

async def start(update: Update, context: CallbackContext) -> None:
    await update.message.reply_text('Hello! Send me a message and I will respond. Use /sendfile to get the file.')

async def send_file(update: Update, context: CallbackContext) -> None:
    chat_id = update.message.chat_id
    file_path = './FAAtt.xlsx'
    try:
        with open(file_path, 'rb') as file:
            await context.bot.send_document(chat_id=chat_id, document=file)
            await update.message.reply_text("Here is the file.")
    except Exception as e:
        logging.error(f"Failed to send file: {e}")
        await update.message.reply_text("Failed to send the file.")

async def help_command(update: Update, context: CallbackContext) -> None:
    help_text = (
        "/start - Start the bot and receive a welcome message.\n"
        "/sendfile - Request the attendance file.\n"
        "/help - Get a list of available commands and their descriptions.\n"
        "To update attendance, you can send a message in the format:\n"
        " - Roll number [DATE (optional)] [P/A]\n"
        " - Name [DATE (optional)] [P/A]\n"
        "Examples:\n"
        " - 12213071 17-Sep P\n"
        " - John Doe 17-Sep A\n"
        " - 12213071\n"
        " - John Doe\n"
    )
    await update.message.reply_text(help_text)

@app.route(f'/{TOKEN}', methods=['POST'])
def telegram_webhook():
    json_str = request.get_data(as_text=True)
    update = Update.de_json(json_str, bot)
    application.update_handler(update)
    return 'ok'

def start_flask():
    app.run(host='0.0.0.0', port=5000)

async def set_webhook():
    webhook_url = f"{WEBHOOK_URL}/{TOKEN}"
    await bot.set_webhook(webhook_url)

def main():
    global bot, application

    bot = Application.builder().token(TOKEN).build()

    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("sendfile", send_file))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error)

    asyncio.run(set_webhook())

    Thread(target=start_flask).start()

if __name__ == '__main__':
    main()
