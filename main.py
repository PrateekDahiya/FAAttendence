from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext
import openpyxl
from datetime import datetime
import logging
import asyncio

TOKEN = "YOUR_BOT_TOKEN"
# Configure logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


async def update_attendance(roll_number: int) -> str:
    file_path = "FAAtt.xlsx"
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    today_date = datetime.today().strftime("%d-%b")

    # Check for existing date column
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

    # Add new date column if not exists
    if date_column is None:
        total_column = sheet.max_column
        if sheet.cell(row=1, column=total_column).value == 'Total':
            total_column -= 1
        sheet.insert_cols(total_column + 1)
        sheet.cell(row=1, column=total_column + 1, value=today_date)
        sheet.cell(row=1, column=total_column + 2, value='Total')
        date_column = total_column + 1

    # Find the student row
    student_row = None
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=2).value == roll_number:
            student_row = row
            break

    if student_row is None:
        return f"Error: Roll number {roll_number} not found."

    # Update attendance
    if sheet.cell(row=student_row, column=date_column).value == 'P':
        return "success"

    sheet.cell(row=student_row, column=date_column).value = 'P'

    # Update total attendance
    total_column = sheet.max_column
    current_total_value = sheet.cell(row=student_row, column=total_column).value
    try:
        current_total = int(current_total_value) if current_total_value is not None else 0
    except ValueError:
        current_total = 0

    sheet.cell(row=student_row, column=total_column).value = current_total + 1

    workbook.save(file_path)
    return "success"


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


async def handle_message(update: Update, context: CallbackContext) -> None:
    text = update.message.text.lower()
    if text.startswith('/sendfile'):
        await send_file(update, context)
    else:
        try:
            roll_number = int(text)
            res = await update_attendance(roll_number)
            if res != "success":
                await update.message.reply_text(res)
        except ValueError:
            await update.message.reply_text(
                "Invalid input. Please enter a valid roll number or use the /sendfile command."
            )


async def error(update: Update, context: CallbackContext) -> None:
    logging.error(f'Update {update} caused error {context.error}')


def main() -> None:
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("sendfile", send_file))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    application.add_error_handler(error)

    # Use the asyncio event loop directly
    loop = asyncio.get_event_loop()
    loop.run_until_complete(application.run_polling())


if __name__ == '__main__':
    main()
