from flask import Flask, request
from telegram import Update, Bot
import openpyxl
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Flask application
app = Flask(__name__)

# Initialize the Telegram bot
TOKEN = "YOUR_BOT_TOKEN"
WEBHOOK_URL = "YOUR_WEBHOOK_URL"
bot = Bot(token=TOKEN)

def update_attendance(roll_number: int) -> str:
    file_path = "FAAtt.xlsx"
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    today_date = datetime.today().strftime("%d-%b")

    date_column = None
    for col in range(5, sheet.max_column + 1):
        cell_value = sheet.cell(row=1, column=col).value
        if cell_value:
            if isinstance(cell_value, datetime):
                cell_date_str = cell_value.strftime("%d-%b")
                if cell_date_str.lower() == today_date.lower():
                    date_column = col
                    break
            elif isinstance(cell_value, str):
                if cell_value.lower() == today_date.lower():
                    date_column = col
                    break

    if date_column is None:
        total_column = sheet.max_column
        if sheet.cell(row=1, column=total_column).value == 'Total':
            total_column -= 1
        sheet.insert_cols(total_column + 1)
        sheet.cell(row=1, column=total_column + 1, value=today_date)
        sheet.cell(row=1, column=total_column + 2, value='Total')
        date_column = total_column + 1

    student_row = None
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=2).value == roll_number:
            student_row = row
            break

    if student_row is None:
        return f"Error: Roll number {roll_number} not found."

    if sheet.cell(row=student_row, column=date_column).value == 'P':
        return f"Attendance for roll number {roll_number} is already marked as Present for {today_date}."

    sheet.cell(row=student_row, column=date_column).value = 'P'

    total_column = sheet.max_column
    current_total_value = sheet.cell(row=student_row, column=total_column).value
    try:
        current_total = int(current_total_value) if current_total_value is not None else 0
    except ValueError:
        current_total = 0

    sheet.cell(row=student_row, column=total_column).value = current_total + 1

    workbook.save(file_path)
    return f"Attendance updated for roll number {roll_number} on {today_date}."

@app.route(f'/{TOKEN}', methods=['POST'])
def respond():
    json_str = request.get_data().decode('UTF-8')
    update = Update.de_json(json_str, bot)
    chat_id = update.message.chat_id
    text = update.message.text.lower()

    if text.startswith('/sendfile'):
        return send_file(chat_id)
    else:
        try:
            roll_number = int(text)
            res = update_attendance(roll_number)
            bot.send_message(chat_id=chat_id, text=res)
        except ValueError:
            bot.send_message(
                chat_id=chat_id,
                text="Invalid input. Please enter a valid roll number or use the /sendfile command."
            )
    return 'ok'

def send_file(chat_id: int) -> str:
    file_path = './FAAtt.xlsx'
    try:
        with open(file_path, 'rb') as file:
            bot.send_document(chat_id=chat_id, document=file)
            bot.send_message(chat_id=chat_id, text="Here is the file.")
    except Exception as e:
        logging.error(f"Failed to send file: {e}")
        bot.send_message(chat_id=chat_id, text="Failed to send the file.")
    return 'ok'

@app.route('/keep_alive', methods=['GET'])
def keep_alive():
    return "Bot is running."

def set_webhook():
    webhook_url = f"{WEBHOOK_URL}/{TOKEN}"
    bot.set_webhook(url=webhook_url)

if __name__ == '__main__':
    set_webhook()
    app.run(host='0.0.0.0', port=8000)
